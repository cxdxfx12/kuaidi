"""
派费计算服务 - 超高性能版（支持300万行级别）
- 规则字典化 O(1) 命中（客户编码/客户名称/区域关键词直接查字典）
- SQLite WAL + 200MB cache + 1GB mmap
- multiprocessing 多进程并行计算（子进程只负责计算，主进程统一入库）
"""
import os
import json
import math
from decimal import Decimal
from datetime import datetime
from typing import Dict, List, Tuple, Optional
import pandas as pd

from app.services.excel_parser import ExcelParser
from app.services.rule_service import RuleService, Rule, apply_weight_rounding
from app.models.database import get_session, Base
from app.models.fee_record import FeeRecord
from app.models.fee_detail import FeeDetail

# 数据库表名常量
TABLE_FEE_DETAIL = "fee_details"

# 性能参数（调大有利于300万行数据）
BATCH_SIZE = 25000                  # 每批计算/入库行数（之前1万，改2.5万）
CHUNK_SIZE_FOR_MP = 200000          # 多进程时每进程负责的chunk大小（约20万行/进程）
MIN_ROWS_FOR_MULTIPROCESS = 100000  # 小于10万行时不启多进程（启动开销比收益大）


# ============================================
# 全局规则索引：所有进程共享（只读）
# ============================================
# 设计：把规则预先构建成多个字典，实现 O(1) 命中
# _STATION_CODE_MAP["ST001"] -> rule_tuple
# _STATION_NAME_MAP["上海总部"] -> rule_tuple
# _REGION_MAP["江苏"] -> rule_tuple
# _GLOBAL_RULES: [rule_tuple, ...]
#
# rule_tuple = (min_w, max_w, first_f, continued_f, min_f, name)
#
# 注意：对于station/region规则，同一key可能有多个规则（不同重量区间），
# 所以 value 是 List[rule_tuple]，按 min_weight 升序排列，查找时顺序匹配第一个命中的。
# 99%场景下其实只有1条规则命中（0-999全区间），所以列表长度几乎总是1。
#
# 另外：客户级规则可能同时包含"区域限制"（rule_kw），匹配时必须同时满足区域关键词。
# 为了保持 O(1) 的命中速度，station 字典存的是"纯客户+无区域限制"的规则，
# "客户+有区域限制"的规则放在 _STATION_WITH_REGION_LIST，做线性扫描（通常很少）。

_STATION_CODE_MAP = None            # type: Dict[str, List[Tuple[float, float, float, float, float, str, str, str, List[str]]]]
_STATION_NAME_MAP = None            # type: Dict[str, List[Tuple[float, float, float, float, float, str, str, str, List[str]]]]
_STATION_WITH_REGION_LIST = None    # type: List[Tuple[List[str], List[str], float, float, float, float, float, str, str, str]]
_REGION_MAP = None                  # type: Dict[str, List[Tuple[float, float, float, float, float, str, str, str]]]
_GLOBAL_RULES = None                # type: List[Tuple[float, float, float, float, float, str, str, str]]
_EMPTY_WEIGHT_FEE = 3.0
_RULES_LOADED = False


def _build_rule_indexes(force_reload: bool = False):
    """
    构建全局规则索引（只需要做1次）
    返回值：不返回，而是设置模块级变量供各函数直接访问（避免函数参数传递开销）
    """
    global _STATION_CODE_MAP, _STATION_NAME_MAP, _STATION_WITH_REGION_LIST
    global _REGION_MAP, _GLOBAL_RULES, _EMPTY_WEIGHT_FEE, _RULES_LOADED

    if _RULES_LOADED and not force_reload:
        return

    # 加载规则
    from app.services.rule_service import RuleService
    rs = RuleService()
    raw_rules = rs.load_rules()
    _EMPTY_WEIGHT_FEE = rs._load_empty_weight_fee()

    _STATION_CODE_MAP = {}
    _STATION_NAME_MAP = {}
    _STATION_WITH_REGION_LIST = []
    _REGION_MAP = {}
    _GLOBAL_RULES = []

    for r in raw_rules:
        # 解析区域关键词列表
        region_kw = []
        if r.regions and r.regions.strip():
            region_kw = [k.strip() for k in r.regions.split(",") if k.strip()]

        # 解析客户编码列表
        station_codes = []
        station_names = []
        if r.stations and r.stations.strip():
            for s in r.stations.split(","):
                s = s.strip()
                if not s:
                    continue
                # 简单启发式：纯数字或"字母+数字"组合当编码，包含中文的当名称
                # 但更稳妥的方式是同时放入两个字典，因为配置时可能混用
                station_codes.append(s)
                station_names.append(s)

        rule_core = (
            float(r.min_weight), float(r.max_weight),
            float(r.first_fee), float(r.continued_fee), float(r.min_fee),
            r.name,
            r.continued_unit or "kg",
            r.weight_rounding or "actual"
        )

        if r.rule_type == "station":
            if not region_kw:
                for code in station_codes:
                    _STATION_CODE_MAP.setdefault(code, []).append(
                        (*rule_core, [])
                    )
                for name in station_names:
                    _STATION_NAME_MAP.setdefault(name, []).append(
                        (*rule_core, [])
                    )
            else:
                _STATION_WITH_REGION_LIST.append(
                    (station_codes, region_kw, *rule_core)
                )

        elif r.rule_type == "region":
            for kw in region_kw:
                _REGION_MAP.setdefault(kw, []).append(rule_core)

        elif r.rule_type == "global":
            _GLOBAL_RULES.append(rule_core)

    # 对每个字典 value（规则列表）按 min_weight 升序排序，便于查找时按顺序匹配
    for v in _STATION_CODE_MAP.values():
        v.sort(key=lambda x: x[0])
    for v in _STATION_NAME_MAP.values():
        v.sort(key=lambda x: x[0])
    for v in _REGION_MAP.values():
        v.sort(key=lambda x: x[0])
    _GLOBAL_RULES.sort(key=lambda x: x[0])

    _RULES_LOADED = True


def _calc_fee_from_rules(weight: float, rules_list: List[Tuple]) -> Optional[Tuple[float, str]]:
    """
    从规则列表中找到第一条重量匹配的规则，计算费用
    :return: (fee, rule_name) 或 None（没命中）
    """
    for rule in rules_list:
        min_w, max_w = rule[0], rule[1]
        if min_w <= weight <= max_w:
            first_f, continued_f, min_f = rule[2], rule[3], rule[4]
            continued_unit = rule[6] if len(rule) > 6 else "kg"
            weight_rounding = rule[7] if len(rule) > 7 else "actual"

            rounded_weight = apply_weight_rounding(weight, weight_rounding, None)

            if rounded_weight <= 1.0:
                fee = first_f
            else:
                continued_weight = rounded_weight - 1.0
                if continued_unit == "100g":
                    units = math.ceil(continued_weight / 0.1)
                    fee = first_f + units * continued_f
                else:
                    fee = first_f + continued_weight * continued_f
            fee = round(max(fee, min_f), 2)
            return fee, rule[5]
    return None


def match_rule_fast(weight: float, region: str, station_code: str = "", station_name: str = "") -> Tuple[float, str, bool]:
    """
    O(1) 速度的规则匹配核心函数（模块级，方便多进程调用）
    返回: (fee, rule_name, is_exception)
    """
    # 确保索引已构建（子进程首次调用会构建一次）
    if not _RULES_LOADED:
        _build_rule_indexes()

    # 1) 无重量 → 用默认价
    if weight is None or weight <= 0:
        return _EMPTY_WEIGHT_FEE, "无重量默认价", False

    region_str = region or ""
    code_str = station_code or ""
    name_str = station_name or ""

    # 2) 客户专属（无区域限制）→ O(1) 字典查找
    if code_str:
        station_rules = _STATION_CODE_MAP.get(code_str)
        if station_rules:
            result = _calc_fee_from_rules(weight, station_rules)
            if result is not None:
                return result[0], result[1], False

    if name_str:
        station_rules = _STATION_NAME_MAP.get(name_str)
        if station_rules:
            result = _calc_fee_from_rules(weight, station_rules)
            if result is not None:
                return result[0], result[1], False

    # 3) 客户专属（有区域限制）→ 线性扫描（数量通常≤5，可忽略）
    if _STATION_WITH_REGION_LIST and (code_str or name_str):
        for item in _STATION_WITH_REGION_LIST:
            codes, region_kws, min_w, max_w, first_f, continued_f, min_f, name, continued_unit, weight_rounding = item
            customer_match = (code_str and code_str in codes) or (name_str and name_str in codes)
            if not customer_match:
                continue
            if region_kws and not any(k in region_str for k in region_kws):
                continue
            if min_w <= weight <= max_w:
                rounded_weight = apply_weight_rounding(weight, weight_rounding, None)
                if rounded_weight <= 1.0:
                    fee = first_f
                else:
                    continued_weight = rounded_weight - 1.0
                    if continued_unit == "100g":
                        units = math.ceil(continued_weight / 0.1)
                        fee = first_f + units * continued_f
                    else:
                        fee = first_f + continued_weight * continued_f
                fee = round(max(fee, min_f), 2)
                return fee, name, False

    # 4) 区域级 → O(1) 字典查找（遍历规则的region关键词，而不是遍历所有规则）
    if region_str:
        # 从 region_str 中提取所有已知关键词
        # 方法：遍历 _REGION_MAP 的所有key，看是否在文本中出现
        # *注意*：_REGION_MAP 的key数量通常≤40（30+省+自治区），所以这是 O(40) ≈ O(1)
        for kw in _REGION_MAP.keys():
            if kw in region_str:
                result = _calc_fee_from_rules(weight, _REGION_MAP[kw])
                if result is not None:
                    return result[0], result[1], False

    # 5) 全局兜底 → 列表长度通常为 1
    if _GLOBAL_RULES:
        result = _calc_fee_from_rules(weight, _GLOBAL_RULES)
        if result is not None:
            return result[0], result[1], False

    return 0.0, "无匹配规则", True


# ============================================
# 多进程 worker：只负责计算（纯函数，无IO，无GUI）
# ============================================
def _process_chunk(args):
    """
    子进程的工作函数：接收一块（行数据列表），计算后返回结果列表
    args = (chunk_rows, idx_map)
    idx_map 是 {"字段名": column_index} 的简化版 —— 但为了速度，直接用位置元组传递
    """
    chunk_rows, idx_list = args
    # idx_list = [idx_tracking, idx_station_code, idx_station_name, idx_courier_code, idx_courier_name,
    #             idx_region_code, idx_region_name, idx_weight, idx_quantity, idx_service,
    #             idx_date, idx_customer_code, idx_customer, idx_remark]
    # chunk_rows = List[Tuple[str|None, ...]] —— 每行的原始列值（预提取的Python list切片）

    # 构建本 chunk 的结果
    results = []
    for row_vals in chunk_rows:
        # 按 idx_list 顺序的字段
        tracking_no = row_vals[0] or ""
        station_code = row_vals[1] or ""
        station_name = row_vals[2] or ""
        courier_code = row_vals[3] or ""
        courier_name = row_vals[4] or ""
        region_code = row_vals[5] or ""
        region_name = row_vals[6] or ""
        weight_str = row_vals[7] or ""
        quantity_str = row_vals[8] or ""
        service_type = row_vals[9] or ""
        raw_date = row_vals[10] or ""
        raw_customer_code = row_vals[11] or ""
        raw_customer = row_vals[12] or ""
        remark = row_vals[13] or ""
        excel_row_index = row_vals[14]  # 预先计算好的 Excel 行号

        # 重量解析
        try:
            weight = float(weight_str) if weight_str else 0.0
        except (ValueError, TypeError):
            weight = 0.0

        # 数量解析
        try:
            quantity = int(float(quantity_str)) if quantity_str else 1
        except (ValueError, TypeError):
            quantity = 1

        fee, rule_name, is_exc = match_rule_fast(weight, region_name, raw_customer_code, raw_customer)

        # 组装成数据库需要的 tuple 结构
        extra_data = None
        if raw_date or raw_customer_code or raw_customer:
            extra_data = json.dumps({
                "business_date": raw_date,
                "customer_code": raw_customer_code,
                "customer_name": raw_customer
            }, ensure_ascii=False)

        results.append((
            excel_row_index,   # row_index (Excel行号)
            tracking_no,
            station_code,
            station_name,
            courier_code,
            courier_name,
            region_code,
            region_name,
            weight,
            quantity,
            service_type,
            extra_data,
            fee,
            rule_name,
            1 if is_exc else 0,
            "invalid_data" if is_exc else None,
            remark or (f"无效重量:{weight_str}" if is_exc and weight <= 0 else "")
        ))
    return results


class CalculateService:
    """计算服务 - 支持300万行大数据量"""

    def __init__(self):
        self.parser = ExcelParser()
        self.rule_service = RuleService()
        # 预先构建规则索引（一次）
        _build_rule_indexes()
        self._empty_weight_fee = _EMPTY_WEIGHT_FEE

    def _load_empty_weight_fee(self) -> float:
        """保留兼容（实际值来自 _build_rule_indexes）"""
        return self._empty_weight_fee

    def _match_rule(self, weight: float, region: str, station_code: str = "", station_name: str = "") -> Tuple[float, str, bool]:
        """兼容方法：底层调用模块级的 match_rule_fast"""
        return match_rule_fast(weight, region, station_code, station_name)

    def _load_rules_list(self) -> List[Tuple]:
        """保留兼容（实际上规则已字典化，此处返回空列表即可，仅占位）"""
        return []

    def import_and_calculate(self, file_path: str, sheet_name=None,
                             progress_callback=None) -> Dict:
        """
        一键导入并计算（支持300万行超大数据量）
        - 单进程模式：< 10万行
        - 多进程模式：>= 10万行，自动启动 multiprocessing 并行计算
        - 支持多Sheet：如果未指定sheet_name，自动合并所有Sheet的数据
        """
        def report(percent, stage):
            if progress_callback:
                try:
                    progress_callback(int(percent), str(stage))
                except Exception:
                    pass

        report(1, "正在读取Excel文件...")

        # ============ 阶段1：读取Excel ============
        if file_path.endswith(".csv"):
            report(3, "正在读取CSV...")
            df = pd.read_csv(file_path, dtype=str)
        else:
            report(5, "正在读取Excel...")
            excel_file = pd.ExcelFile(file_path)
            
            if sheet_name:
                df = excel_file.parse(sheet_name=sheet_name, dtype=str)
                report(15, f"已读取Sheet: {sheet_name}")
            else:
                sheets = excel_file.sheet_names
                report(8, f"发现 {len(sheets)} 个Sheet，正在合并...")
                
                dfs = []
                for i, s in enumerate(sheets):
                    df_sheet = excel_file.parse(sheet_name=s, dtype=str)
                    dfs.append(df_sheet)
                    report(8 + int((i + 1) / len(sheets) * 7), f"读取Sheet {i+1}/{len(sheets)}: {s}")
                
                if len(dfs) > 1:
                    df = pd.concat(dfs, ignore_index=True)
                    report(15, f"合并完成，共 {len(df)} 行")
                else:
                    df = dfs[0]

        row_count = len(df)
        columns = list(df.columns)

        # 列名匹配
        from app.services.column_matcher import ColumnMatcher
        matcher = ColumnMatcher()
        matched = matcher.match_columns(columns)

        report(15, f"文件读取完成，共 {row_count} 行数据，正在初始化...")

        # 预计算列索引
        col_map = matched.get("matched", {})

        def _get_col_idx(std_name: str) -> int:
            actual = col_map.get(std_name)
            if actual and actual in columns:
                return columns.index(actual)
            return -1

        idx_tracking = _get_col_idx("tracking_no")
        idx_station_code = _get_col_idx("station_code")
        idx_station_name = _get_col_idx("station_name")
        idx_courier_code = _get_col_idx("courier_code")
        idx_courier_name = _get_col_idx("courier_name")
        idx_region_code = _get_col_idx("region_code")
        idx_region_name = _get_col_idx("region_name")
        idx_weight = _get_col_idx("weight")
        idx_quantity = _get_col_idx("quantity")
        idx_service = _get_col_idx("service_type")
        idx_date = _get_col_idx("business_date")
        idx_customer_code = _get_col_idx("customer_code")
        idx_customer = _get_col_idx("customer_name")
        idx_remark = _get_col_idx("remark")

        # 把列索引按固定顺序打包（传给多进程 worker 用）
        column_indices = [
            idx_tracking, idx_station_code, idx_station_name,
            idx_courier_code, idx_courier_name, idx_region_code, idx_region_name,
            idx_weight, idx_quantity, idx_service,
            idx_date, idx_customer_code, idx_customer, idx_remark
        ]

        # ============ 关键优化：把需要的列提取为Python list ============
        def _extract_col(idx: int):
            if idx < 0:
                return None
            raw_col = df.iloc[:, idx].tolist()
            return ["" if (v is None or (isinstance(v, float) and v != v))
                    else str(v).strip() for v in raw_col]

        # 把所有需要的列放进一个列表（便于 multiprocessing 传递）
        raw_cols = []
        for idx in column_indices:
            raw_cols.append(_extract_col(idx))
        # 释放大 DataFrame 内存（重要！对于几百万行，这一步释放约几百MB）
        del df
        import gc
        gc.collect()

        # ============ 阶段2：创建任务记录 ============
        record = FeeRecord(
            file_name=os.path.basename(file_path),
            file_path=file_path,
            file_size=os.path.getsize(file_path),
            total_rows=row_count,
            status="processing"
        )

        session = get_session()
        try:
            session.add(record)
            session.commit()
            record_id = record.id

            # ============ 阶段3：计算并入库 ============
            # 决策：10万行以下走单进程（省启动开销），10万行以上自动多进程
            use_multiprocess = row_count >= MIN_ROWS_FOR_MULTIPROCESS

            success_count = 0
            exception_count = 0
            total_fee = Decimal("0")

            if use_multiprocess:
                import multiprocessing as mp
                cpu_count = mp.cpu_count()
                # 留1核给GUI/系统，避免卡死；至少1核
                pool_size = max(1, min(cpu_count - 1, 8))

                report(18, f"大数据模式：启动 {pool_size} 个进程并行计算 "
                           f"（CPU {cpu_count}核 · 共 {row_count} 行）")

                # 把原始列数据组装成固定格式的 row_tuples 列表
                # 每个 row_tuple = (tracking, station_code, station_name, courier_code, courier_name,
                #                   region_code, region_name, weight_str, quantity_str, service_type,
                #                   date_str, customer_code_str, customer_name, remark, excel_row_idx)
                report(20, "正在准备分块数据...")
                all_row_tuples = []
                for pos in range(row_count):
                    excel_row = pos + 2  # Excel行号从1开始，+1表头
                    row_tup = []
                    for col in raw_cols:
                        row_tup.append(col[pos] if col is not None else "")
                    row_tup.append(excel_row)  # 附加excel行号
                    all_row_tuples.append(row_tup)

                # 分chunk
                num_chunks = pool_size * 3  # 每进程大约处理3个chunk，更均匀
                chunk_size = max(CHUNK_SIZE_FOR_MP // 4,  # 每chunk约5万行，更灵敏
                                 (row_count + num_chunks - 1) // num_chunks)
                chunks = []
                for i in range(0, row_count, chunk_size):
                    end = min(i + chunk_size, row_count)
                    chunks.append((all_row_tuples[i:end], None))

                # 释放原始列数据
                del raw_cols
                del all_row_tuples
                gc.collect()

                report(22, f"已分 {len(chunks)} 块，启动并行计算...")

                # 使用 multiprocessing Pool + imap_unordered + 动态进度
                processed = 0
                total_to_process = row_count
                with mp.Pool(processes=pool_size) as pool:
                    for chunk_results in pool.imap_unordered(_process_chunk, chunks, chunksize=1):
                        # 每处理完一个chunk，批量入库（用BATCH_SIZE=25000再细分）
                        for b_start in range(0, len(chunk_results), BATCH_SIZE):
                            b_end = min(b_start + BATCH_SIZE, len(chunk_results))
                            batch_to_insert = []
                            for r in chunk_results[b_start:b_end]:
                                is_exc = r[14]  # 第15个字段（0-based）是 is_exception_flag
                                if is_exc:
                                    exception_count += 1
                                else:
                                    success_count += 1
                                    total_fee += Decimal(str(r[12]))
                                batch_to_insert.append((
                                    record_id,
                                    r[0],    # row_index
                                    r[1],    # tracking
                                    r[2],    # station_code
                                    r[3],    # station_name
                                    r[4],    # courier_code
                                    r[5],    # courier_name
                                    r[6],    # region_code
                                    r[7],    # region_name
                                    r[8],    # weight
                                    r[9],    # quantity
                                    r[10],   # service_type
                                    r[11],   # extra_data (原 json)
                                    r[12],   # fee
                                    r[13],   # rule_name
                                    is_exc,
                                    r[15],   # exception_type
                                    r[16],   # remark
                                ))
                            # 入库
                            self._bulk_insert_details(session, batch_to_insert)

                        processed += len(chunk_results)
                        progress = 22 + processed / total_to_process * 68
                        report(progress,
                               f"并行计算中... {processed}/{total_to_process} 行 "
                               f"({progress:.0f}%)  成功 {success_count}，异常 {exception_count}")

            else:
                # ============ 单进程模式（< 10万行） ============
                report(20, f"单进程模式：开始计算 {row_count} 行...")
                all_row_tuples = []
                for pos in range(row_count):
                    excel_row = pos + 2
                    row_tup = []
                    for col in raw_cols:
                        row_tup.append(col[pos] if col is not None else "")
                    row_tup.append(excel_row)
                    all_row_tuples.append(row_tup)

                # 直接复用 _process_chunk
                results = _process_chunk((all_row_tuples, None))
                # 分批写入（25000行一批）
                for b_start in range(0, len(results), BATCH_SIZE):
                    b_end = min(b_start + BATCH_SIZE, len(results))
                    batch_to_insert = []
                    for r in results[b_start:b_end]:
                        is_exc = r[14]
                        if is_exc:
                            exception_count += 1
                        else:
                            success_count += 1
                            total_fee += Decimal(str(r[12]))
                        batch_to_insert.append((
                            record_id,
                            r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7],
                            r[8], r[9], r[10], r[11], r[12], r[13],
                            is_exc, r[15], r[16]
                        ))
                    self._bulk_insert_details(session, batch_to_insert)
                    progress = 22 + min(b_end, row_count) / row_count * 68
                    report(progress,
                           f"计算中... {b_end}/{row_count} 行 ({progress:.0f}%)"
                           f"  成功 {success_count}，异常 {exception_count}")

            # ============ 阶段4：更新记录状态 (90%-100%) ============
            report(90, "计算完成，正在更新统计信息...")

            record.success_rows = success_count
            record.error_rows = exception_count
            record.total_fee = total_fee
            record.status = "success"
            record.completed_at = datetime.now()
            session.commit()

            report(100, f"✅ 全部完成！总计 {row_count} 行，运费 ¥{float(total_fee):.2f}")

            return {
                "record_id": record_id,
                "total_fee": float(total_fee),
                "success_count": success_count,
                "exception_count": exception_count,
                "total_rows": row_count
            }

        except Exception as e:
            session.rollback()
            record.status = "failed"
            record.error_message = str(e)
            session.commit()
            raise e
        finally:
            session.close()

    def _bulk_insert_details(self, session, values: List[Tuple]):
        """
        批量写入运费明细 - 使用原生SQLite连接 executemany，比ORM快10-20倍
        """
        if not values:
            return

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        params = []
        for v in values:
            params.append((
                v[0], v[1], v[2], v[3], v[4],
                v[5], v[6], v[7], v[8], v[9],
                v[10], v[11], v[12], v[13], v[14],
                v[15], v[16], v[17], now
            ))

        # 使用原生SQLite连接做最快的批量写入
        conn = session.connection().connection  # 拿到原生 sqlite3.Connection
        cursor = conn.cursor()
        try:
            cursor.executemany(
                f"""
                INSERT INTO {TABLE_FEE_DETAIL} (
                    record_id, row_index, tracking_no, station_code, station_name,
                    courier_code, courier_name, region_code, region_name, weight,
                    quantity, service_type, original_data, calculated_fee, rule_name,
                    is_exception, exception_type, remark, created_at
                ) VALUES (
                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?, ?, ?, ?, ?
                )
                """,
                params
            )
            conn.commit()
        finally:
            cursor.close()
