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
# 全局规则索引：初始化为空字典/列表，防止 None 调用 .get() 崩溃
# ============================================
# 客户级索引（客户名称/客户编码 - 最优先匹配）
_STATION_CODE_MAP = {}              # type: Dict[str, List[Tuple[...]]]
_STATION_NAME_MAP = {}              # type: Dict[str, List[Tuple[...]]]
# 网点级索引（网点名称/网点编码 - 次要匹配）
_OUTLET_CODE_MAP = {}               # type: Dict[str, List[Tuple[...]]]
_OUTLET_NAME_MAP = {}               # type: Dict[str, List[Tuple[...]]]
# 带区域限制的客户专属规则列表
_STATION_WITH_REGION_LIST = []      # type: List[Tuple[...]]
# 区域级规则
_REGION_MAP = {}                    # type: Dict[str, List[Tuple[...]]]
# 全局规则
_GLOBAL_RULES = []                  # type: List[Tuple[...]]
# 活动加价规则
_PROMOTION_RULES = []               # type: List[Dict]
_EMPTY_WEIGHT_FEE = 3.0
_RULES_LOADED = False


def _build_rule_indexes(force_reload: bool = False):
    """
    构建全局规则索引（只需要做1次）
    关键设计：规则的 stations 字段同时存入客户级索引和网点级索引，
    这样无论是客户名称"珀莱雅"还是网点名称"浙江杭州集包工厂"都能正确匹配。
    匹配优先级：客户名称 > 客户编码 > 网点名称 > 网点编码 > 区域 > 全局
    """
    global _STATION_CODE_MAP, _STATION_NAME_MAP, _OUTLET_CODE_MAP, _OUTLET_NAME_MAP
    global _STATION_WITH_REGION_LIST, _REGION_MAP, _GLOBAL_RULES, _PROMOTION_RULES
    global _EMPTY_WEIGHT_FEE, _RULES_LOADED

    if _RULES_LOADED and not force_reload:
        return

    try:
        from app.services.rule_service import RuleService
        rs = RuleService()
        raw_rules = rs.load_rules()
        _EMPTY_WEIGHT_FEE = rs._load_empty_weight_fee()
        _PROMOTION_RULES = rs.load_promotion_rules()
    except Exception:
        _RULES_LOADED = True
        return

    _STATION_CODE_MAP = {}     # 客户编码索引
    _STATION_NAME_MAP = {}     # 客户名称索引
    _OUTLET_CODE_MAP = {}      # 网点编码索引
    _OUTLET_NAME_MAP = {}      # 网点名称索引
    _STATION_WITH_REGION_LIST = []
    _REGION_MAP = {}
    _GLOBAL_RULES = []

    for r in raw_rules:
        region_kw = []
        if r.regions and r.regions.strip():
            region_kw = [k.strip() for k in r.regions.split(",") if k.strip()]

        # stations字段可能包含：客户名称/客户编码/网点名称/网点编码，全部尝试匹配
        station_values = []
        if r.stations and r.stations.strip():
            for s in r.stations.split(","):
                s = s.strip()
                if s:
                    station_values.append(s)

        rule_core = (
            float(r.min_weight), float(r.max_weight),
            float(r.first_fee), float(r.continued_fee), float(r.min_fee),
            r.name,
            r.continued_unit or "kg",
            r.weight_rounding or "actual"
        )

        if r.rule_type == "station":
            if not region_kw:
                # 无区域限制的客户/网点规则：
                # stations字段的值同时存入客户级索引和网点级索引
                for v in station_values:
                    # 作为客户编码/客户名称匹配
                    _STATION_CODE_MAP.setdefault(v, []).append((*rule_core, []))
                    _STATION_NAME_MAP.setdefault(v, []).append((*rule_core, []))
                    # 作为网点编码/网点名称匹配
                    _OUTLET_CODE_MAP.setdefault(v, []).append((*rule_core, []))
                    _OUTLET_NAME_MAP.setdefault(v, []).append((*rule_core, []))
            else:
                # 有区域限制的客户专属规则
                _STATION_WITH_REGION_LIST.append(
                    (station_values, station_values, region_kw, *rule_core)
                )

        elif r.rule_type == "region":
            for kw in region_kw:
                _REGION_MAP.setdefault(kw, []).append(rule_core)

        elif r.rule_type == "global":
            _GLOBAL_RULES.append(rule_core)

    # 按最小重量排序，确保小重量优先命中
    for v in _STATION_CODE_MAP.values():
        v.sort(key=lambda x: x[0])
    for v in _STATION_NAME_MAP.values():
        v.sort(key=lambda x: x[0])
    for v in _OUTLET_CODE_MAP.values():
        v.sort(key=lambda x: x[0])
    for v in _OUTLET_NAME_MAP.values():
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


def _parse_date(date_str: str):
    """解析日期，支持多种格式：2026-06-19 / 2026/6/19 / 2026.6.19 / 20260619 / 2026年6月19日"""
    if not date_str:
        return None
    s = str(date_str).strip()
    if not s:
        return None
    # 统一去掉中文字符和多余符号
    for ch in ["年", "月", "日", ".", "-", "/"]:
        s = s.replace(ch, "-")
    # 去掉连续的"-"
    while "--" in s:
        s = s.replace("--", "-")
    s = s.strip("-")
    parts = s.split("-")
    if len(parts) >= 3:
        try:
            y = int(parts[0])
            m = int(parts[1])
            d = int(parts[2])
            return datetime(y, m, d)
        except (ValueError, TypeError):
            pass
    # 兜底：纯8位数字如 20260619
    if s.isdigit() and len(s) == 8:
        try:
            return datetime(int(s[:4]), int(s[4:6]), int(s[6:8]))
        except (ValueError, TypeError):
            pass
    # 最后再尝试标准格式
    for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"]:
        try:
            return datetime.strptime(str(date_str).strip(), fmt)
        except Exception:
            continue
    return None


def _apply_promotion(base_fee: float, weight: float, region_str: str = "") -> Tuple[float, str]:
    """统一应用活动加价规则（支持按省份限定，省份留空=所有省份）"""
    if not _PROMOTION_RULES:
        return base_fee, ""

    try:
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        region_str = region_str or ""

        for pr in _PROMOTION_RULES:
            try:
                # 1. 检查日期
                start = _parse_date(pr.get("start_date", ""))
                end = _parse_date(pr.get("end_date", ""))
                if start is None or end is None:
                    continue
                if not (start <= today <= end):
                    continue

                # 2. 检查省份限定（regions字段为空=不限定，否则必须包含区域关键词）
                regions_val = str(pr.get("regions", "")).strip()
                if regions_val:
                    # 按逗号拆分省份关键词
                    region_kws = [k.strip() for k in regions_val.split(",") if k.strip()]
                    if region_kws and not any(k in region_str for k in region_kws):
                        # 有限定省份，但当前订单的区域不在限定范围内 → 跳过
                        continue

                # 3. 计算加价金额
                markup_type = str(pr.get("markup_type", "percent")).strip().lower()
                try:
                    markup_value = float(str(pr.get("markup_value", "0")).strip())
                except (ValueError, TypeError):
                    continue

                if markup_value <= 0:
                    continue

                promo_amount = 0.0
                if markup_type == "fixed":
                    promo_amount = markup_value
                elif markup_type == "weight":
                    promo_amount = weight * markup_value
                elif markup_type == "percent":
                    promo_amount = base_fee * (markup_value / 100.0)
                else:
                    continue

                promo_amount = round(promo_amount, 2)
                if promo_amount > 0:
                    promo_name = str(pr.get("name", "活动加价"))
                    region_note = f"[{regions_val}]" if regions_val else ""
                    return round(base_fee + promo_amount, 2), f"+ {promo_name}{region_note}(+¥{promo_amount})"
            except Exception:
                continue
    except Exception:
        pass

    return base_fee, ""


def match_rule_fast(weight: float, region: str, station_code: str = "", station_name: str = "",
                     customer_code: str = "", customer_name: str = "") -> Tuple[float, str, bool]:
    """
    规则匹配核心函数（模块级，方便多进程调用）
    匹配优先级（从高到低）：
      1. 客户名称（customer_name）- 如"珀莱雅"
      2. 客户编码（customer_code）
      3. 网点名称（station_name）- 如"浙江杭州集包工厂"
      4. 网点编码（station_code）
      5. 有区域限制的客户/网点专属规则
      6. 区域级规则
      7. 全局兜底规则
    返回: (fee, rule_name, is_exception)
    """
    if not _RULES_LOADED:
        _build_rule_indexes(force_reload=True)

    if weight is None or weight <= 0:
        base_fee = _EMPTY_WEIGHT_FEE
        rule_name = "无重量默认价"
        final_fee, promo_suffix = _apply_promotion(base_fee, weight, region or "")
        if promo_suffix:
            return final_fee, f"{rule_name} {promo_suffix}", False
        return final_fee, rule_name, False

    region_str = region or ""
    cust_name_str = customer_name or ""
    cust_code_str = customer_code or ""
    outlet_name_str = station_name or ""
    outlet_code_str = station_code or ""

    base_fee = 0.0
    rule_name = ""
    matched = False

    # ==================== 路径1：优先匹配 客户名称（最高优先级） ====================
    if not matched and cust_name_str:
        # 精确匹配客户名称
        cust_rules = _STATION_NAME_MAP.get(cust_name_str)
        if cust_rules:
            result = _calc_fee_from_rules(weight, cust_rules)
            if result is not None:
                base_fee, rule_name = result
                matched = True

    # ==================== 路径2：匹配 客户编码 ====================
    if not matched and cust_code_str:
        cust_rules = _STATION_CODE_MAP.get(cust_code_str)
        if cust_rules:
            result = _calc_fee_from_rules(weight, cust_rules)
            if result is not None:
                base_fee, rule_name = result
                matched = True

    # ==================== 路径3：匹配 网点名称 ====================
    if not matched and outlet_name_str:
        outlet_rules = _OUTLET_NAME_MAP.get(outlet_name_str)
        if outlet_rules:
            result = _calc_fee_from_rules(weight, outlet_rules)
            if result is not None:
                base_fee, rule_name = result
                matched = True

    # ==================== 路径4：匹配 网点编码 ====================
    if not matched and outlet_code_str:
        outlet_rules = _OUTLET_CODE_MAP.get(outlet_code_str)
        if outlet_rules:
            result = _calc_fee_from_rules(weight, outlet_rules)
            if result is not None:
                base_fee, rule_name = result
                matched = True

    # ==================== 路径5：带区域限制的客户/网点专属规则 ====================
    if not matched and _STATION_WITH_REGION_LIST:
        for item in _STATION_WITH_REGION_LIST:
            codes = item[0]
            names = item[1]
            region_kws = item[2]
            min_w = item[3]
            max_w = item[4]
            first_f = item[5]
            continued_f = item[6]
            min_f = item[7]
            r_name = item[8]
            continued_unit = item[9] if len(item) > 9 else "kg"
            weight_rounding = item[10] if len(item) > 10 else "actual"

            # 尝试匹配：客户名称 / 客户编码 / 网点名称 / 网点编码
            name_match = (cust_name_str and cust_name_str in names) or \
                         (cust_code_str and cust_code_str in codes) or \
                         (outlet_name_str and outlet_name_str in names) or \
                         (outlet_code_str and outlet_code_str in codes)
            if not name_match:
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
                base_fee = round(max(fee, min_f), 2)
                rule_name = r_name
                matched = True
                break

    # ==================== 路径6：区域级规则 ====================
    if not matched and region_str:
        for kw in _REGION_MAP.keys():
            if kw in region_str:
                result = _calc_fee_from_rules(weight, _REGION_MAP[kw])
                if result is not None:
                    base_fee, rule_name = result
                    matched = True
                    break

    # ==================== 路径7：全局兜底规则 ====================
    if not matched and _GLOBAL_RULES:
        result = _calc_fee_from_rules(weight, _GLOBAL_RULES)
        if result is not None:
            base_fee, rule_name = result
            matched = True

    # ==================== 没有任何规则命中，返回异常 ====================
    if not matched:
        return 0.0, "无匹配规则", True

    # 最后统一应用活动加价
    final_fee, promo_suffix = _apply_promotion(base_fee, weight, region_str)
    if promo_suffix:
        return final_fee, f"{rule_name} {promo_suffix}", False
    return final_fee, rule_name, False


# ============================================
# 多进程 worker：只负责计算（纯函数，无IO，无GUI）
# ============================================
def _process_chunk(args):
    """
    子进程工作函数：接收一块数据（行列表），计算后返回结果列表
    规则索引在子进程启动时通过 initializer 重建一次
    去重由主进程在分chunk前完成
    """
    try:
        chunk_rows, idx_list = args

        # 确保规则索引已构建（子进程第一次调用时需要）
        if not _RULES_LOADED:
            _build_rule_indexes(force_reload=True)

        results = []

        for row_vals in chunk_rows:
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
            excel_row_index = row_vals[14]

            try:
                weight = float(weight_str) if weight_str else 0.0
            except (ValueError, TypeError):
                weight = 0.0

            try:
                quantity = int(float(quantity_str)) if quantity_str else 1
            except (ValueError, TypeError):
                quantity = 1

            fee, rule_name, is_exc = match_rule_fast(weight, region_name,
                                                      station_code, station_name,
                                                      raw_customer_code, raw_customer)

            extra_data = None
            if raw_date or raw_customer_code or raw_customer:
                extra_data = json.dumps({
                    "business_date": raw_date,
                    "customer_code": raw_customer_code,
                    "customer_name": raw_customer
                }, ensure_ascii=False)

            results.append((
                excel_row_index, tracking_no, station_code, station_name,
                courier_code, courier_name, region_code, region_name,
                weight, quantity, service_type, extra_data,
                fee, rule_name, 1 if is_exc else 0,
                "invalid_data" if is_exc else None,
                remark or (f"无效重量:{weight_str}" if is_exc and weight <= 0 else "")
            ))

        return results
    except Exception:
        # 单个 chunk 处理失败时返回空列表，不影响其他 chunk
        return []


class CalculateService:
    """计算服务 - 支持300万行大数据量"""

    def __init__(self):
        self.parser = ExcelParser()
        self.rule_service = RuleService()
        # 每次新建 CalculateService 都强制重建索引，确保获取最新规则（包括活动加价规则）
        _build_rule_indexes(force_reload=True)
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

        # ============ 阶段0：验证规则索引 ============
        # 确保规则索引已正确加载（对于多进程也很重要：验证主进程能正确读取规则）
        _build_rule_indexes(force_reload=True)
        customer_rule_count = sum(len(v) for v in _STATION_NAME_MAP.values()) + \
                              sum(len(v) for v in _STATION_CODE_MAP.values())
        outlet_rule_count = sum(len(v) for v in _OUTLET_NAME_MAP.values()) + \
                            sum(len(v) for v in _OUTLET_CODE_MAP.values())
        report(1, f"规则校验完成（客户级规则 {customer_rule_count} 条，"
                f"网点级规则 {outlet_rule_count} 条，"
                f"区域规则 {sum(len(v) for v in _REGION_MAP.values())} 条，"
                f"全局规则 {len(_GLOBAL_RULES)} 条，"
                f"活动规则 {len(_PROMOTION_RULES)} 条）")

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
            # 关键修复：先删除该record_id可能存在的旧数据（防止程序崩溃重启后重复）
            try:
                conn_pre = session.connection().connection
                cur_pre = conn_pre.cursor()
                cur_pre.execute(f"DELETE FROM {TABLE_FEE_DETAIL} WHERE record_id = ?", (record_id,))
                conn_pre.commit()
                cur_pre.close()
            except Exception:
                pass

            # 决策：10万行以下走单进程（省启动开销），10万行以上自动多进程
            use_multiprocess = row_count >= MIN_ROWS_FOR_MULTIPROCESS

            # ============ 阶段3-a：统一构建行数据 + 全局去重（单号唯一）
            all_row_tuples = []
            seen_global = set()
            duplicate_count = 0
            empty_tracking_count = 0

            for pos in range(row_count):
                excel_row = pos + 2
                row_tup = []
                for col in raw_cols:
                    row_tup.append(col[pos] if col is not None else "")
                row_tup.append(excel_row)

                tracking_no = row_tup[0] or ""
                if not tracking_no:
                    # 空单号：保留但不去重（否则所有空单号只保留第一条）
                    empty_tracking_count += 1
                    all_row_tuples.append(row_tup)
                    continue
                if tracking_no in seen_global:
                    duplicate_count += 1
                    continue
                seen_global.add(tracking_no)
                all_row_tuples.append(row_tup)

            # 更新去重后的行数
            row_count_original = row_count
            row_count = len(all_row_tuples)

            report(19 if not use_multiprocess else 19,
                   f"数据构建完成。原始 {row_count_original:,} 行，"
                   f"去重后 {row_count:,} 行 "
                   f"（跳过重复 {duplicate_count:,} 行，空单号 {empty_tracking_count:,} 行）")

            # 释放raw_cols内存
            del raw_cols
            gc.collect()

            if use_multiprocess:
                import multiprocessing as mp
                cpu_count = mp.cpu_count()
                pool_size = max(1, min(cpu_count - 1, 8))

                report(20, f"大数据模式：启动 {pool_size} 个进程并行计算 "
                           f"（CPU {cpu_count}核）")

                num_chunks = pool_size * 3
                chunk_size = max(CHUNK_SIZE_FOR_MP // 4,
                                 (row_count + num_chunks - 1) // num_chunks)
                chunks = []
                for i in range(0, row_count, chunk_size):
                    end = min(i + chunk_size, row_count)
                    chunks.append((all_row_tuples[i:end], None))

                del all_row_tuples
                del seen_global
                gc.collect()

                report(22, f"已分 {len(chunks)} 块，启动并行计算...")

                processed = 0
                total_to_process = row_count
                success_count = 0
                exception_count = 0
                total_fee = Decimal("0")

                # 应用层去重集合：确保同一批次不会重复写入
                inserted_tracking = set()

                with mp.Pool(processes=pool_size) as pool:
                    for chunk_results in pool.imap_unordered(_process_chunk, chunks, chunksize=1):
                        for b_start in range(0, len(chunk_results), BATCH_SIZE):
                            b_end = min(b_start + BATCH_SIZE, len(chunk_results))
                            batch_to_insert = []
                            for r in chunk_results[b_start:b_end]:
                                is_exc = r[14]
                                t_no = r[1] or ""  # tracking_no
                                # 应用层去重：同一record_id下不重复单号
                                if t_no and t_no in inserted_tracking:
                                    continue
                                if t_no:
                                    inserted_tracking.add(t_no)
                                if is_exc:
                                    exception_count += 1
                                else:
                                    success_count += 1
                                    total_fee += Decimal(str(r[12]))
                                batch_to_insert.append((
                                    record_id,
                                    r[0], r[1], r[2], r[3], r[4],
                                    r[5], r[6], r[7], r[8], r[9],
                                    r[10], r[11], r[12], r[13],
                                    is_exc, r[15], r[16],
                                ))
                            self._bulk_insert_details(session, batch_to_insert)

                        processed += len(chunk_results)
                        progress = 22 + min(processed / total_to_process, 1.0) * 68
                        report(progress,
                               f"并行计算中... {processed:,}/{total_to_process:,} 行 "
                               f"({int(progress)}%)  成功 {success_count:,}，异常 {exception_count:,}")

                del inserted_tracking

            else:
                # ============ 单进程模式（< 10万行） ============
                report(22, f"单进程模式：开始计算 {row_count:,} 行...")

                results = _process_chunk((all_row_tuples, None))

                # 应用层去重 + 分批写入
                success_count = 0
                exception_count = 0
                total_fee = Decimal("0")
                inserted_tracking = set()

                for b_start in range(0, len(results), BATCH_SIZE):
                    b_end = min(b_start + BATCH_SIZE, len(results))
                    batch_to_insert = []
                    for r in results[b_start:b_end]:
                        is_exc = r[14]
                        t_no = r[1] or ""
                        if t_no and t_no in inserted_tracking:
                            continue
                        if t_no:
                            inserted_tracking.add(t_no)
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
                           f"计算中... {b_end:,}/{row_count:,} 行 ({int(progress)}%)"
                           f"  成功 {success_count:,}，异常 {exception_count:,}")

                del results, inserted_tracking, all_row_tuples, seen_global

            # ============ 阶段4：更新记录状态 (90%-100%) ============
            # 关键修复：从数据库读取实际插入的行数，替换可能不准确的内存计数
            actual_rows = success_count + exception_count
            actual_fee = total_fee
            try:
                verify_conn = session.connection().connection
                verify_cur = verify_conn.cursor()
                verify_cur.execute(
                    f"SELECT COUNT(*) FROM {TABLE_FEE_DETAIL} WHERE record_id = ?",
                    (record_id,)
                )
                actual_rows = verify_cur.fetchone()[0]
                verify_cur.execute(
                    f"SELECT COUNT(*) FROM {TABLE_FEE_DETAIL} WHERE record_id = ? AND is_exception = 1",
                    (record_id,)
                )
                actual_exc = verify_cur.fetchone()[0]
                verify_cur.execute(
                    f"SELECT COALESCE(SUM(calculated_fee), 0) FROM {TABLE_FEE_DETAIL} WHERE record_id = ?",
                    (record_id,)
                )
                actual_fee = Decimal(str(round(float(verify_cur.fetchone()[0]), 2)))
                verify_cur.close()
                report(92, f"数据库验证：实际 {actual_rows:,} 行（预期 {success_count + exception_count:,}），"
                           f"成功 {actual_rows - actual_exc:,}，异常 {actual_exc:,}")
                success_count = actual_rows - actual_exc
                exception_count = actual_exc
            except Exception:
                pass

            record.total_rows = actual_rows
            record.success_rows = success_count
            record.error_rows = exception_count
            record.total_fee = actual_fee
            record.status = "success"
            record.completed_at = datetime.now()
            session.commit()

            report(100, f"✅ 全部完成！总计 {actual_rows:,} 行，运费 ¥{float(actual_fee):.2f}")

            return {
                "record_id": record_id,
                "total_fee": float(actual_fee),
                "success_count": success_count,
                "exception_count": exception_count,
                "total_rows": actual_rows
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

        # 使用原生 SQLite 连接做最快的批量写入
        # INSERT OR IGNORE：如果 (record_id, tracking_no) 已存在则跳过，防止重复插入
        conn = session.connection().connection  # 拿到原生 sqlite3.Connection
        cursor = conn.cursor()
        try:
            cursor.executemany(
                f"""
                INSERT OR IGNORE INTO {TABLE_FEE_DETAIL} (
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
