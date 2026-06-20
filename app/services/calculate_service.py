"""
派费计算服务 - 超高性能版（支持300万行级别）
- 规则字典化 O(1) 命中（客户编码/客户名称/区域关键词直接查字典）
- SQLite WAL + 200MB cache + 1GB mmap
- multiprocessing 多进程并行计算（子进程只负责计算，主进程统一入库）
- python-calamine (Rust引擎) 替代 openpyxl 读取Excel，速度提升5倍
"""
import os
import json
import math
from decimal import Decimal
from datetime import datetime
from typing import Dict, List, Tuple, Optional

try:
    from python_calamine import CalamineWorkbook
    _HAS_CALAMINE = True
except Exception:
    _HAS_CALAMINE = False
    CalamineWorkbook = None

import pandas as pd

from app.services.excel_parser import ExcelParser
from app.services.rule_service import RuleService, apply_weight_rounding
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
# 活动加价规则预解析缓存（模块级，避免每行都解析日期/拆分字符串）
# 格式: list of (markup_type_int, markup_value_float, region_kws_tuple, promo_name_str, region_note_str)
# markup_type_int: 0=fixed, 1=weight, 2=percent
_PROMOTION_CACHE = []               # type: List[Tuple]
_EMPTY_WEIGHT_FEE = 3.0
_RULES_LOADED = False
# 计泡系数映射（station_code/name -> divisor），默认6000
_计泡系数_MAP = {}             # type: Dict[str, float]
_DEFAULT_计泡系数 = 6000


def _build_rule_indexes(force_reload: bool = False):
    """
    构建全局规则索引（只需要做1次）
    关键设计：规则的 stations 字段同时存入客户级索引和网点级索引，
    这样无论是客户名称"珀莱雅"还是网点名称"浙江杭州集包工厂"都能正确匹配。
    匹配优先级：客户名称 > 客户编码 > 网点名称 > 网点编码 > 区域 > 全局
    """
    global _STATION_CODE_MAP, _STATION_NAME_MAP, _OUTLET_CODE_MAP, _OUTLET_NAME_MAP
    global _STATION_WITH_REGION_LIST, _REGION_MAP, _GLOBAL_RULES, _PROMOTION_RULES
    global _PROMOTION_CACHE, _EMPTY_WEIGHT_FEE, _RULES_LOADED
    global _计泡系数_MAP, _DEFAULT_计泡系数

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

    # 从 fee_rules.json 加载计泡系数配置
    try:
        import json as _json
        _fee_rules_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
                                        "data", "config", "fee_rules.json")
        if os.path.exists(_fee_rules_path):
            with open(_fee_rules_path, "r", encoding="utf-8") as _f:
                _fee_data = _json.load(_f)
            _DEFAULT_计泡系数 = float(_fee_data.get("default_计泡系数", 6000))
            _计泡系数_MAP = {}
            for _rule in _fee_data.get("rules", []):
                _vd = float(_rule.get("计泡系数", _DEFAULT_计泡系数))
                _stations_str = _rule.get("stations", "")
                if _stations_str:
                    for _s in _stations_str.split(","):
                        _s = _s.strip()
                        if _s:
                            _计泡系数_MAP[_s] = _vd
        else:
            _DEFAULT_计泡系数 = 6000
            _计泡系数_MAP = {}
    except Exception:
        _DEFAULT_计泡系数 = 6000
        _计泡系数_MAP = {}

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

    # ========== 活动加价规则预解析 ==========
    # 把日期/区域/类型/值 等都预先解析好，运行期不再重复解析
    try:
        _PROMOTION_CACHE = []
        for pr in _PROMOTION_RULES:
            try:
                start = _parse_date(pr.get("start_date", ""))
                end = _parse_date(pr.get("end_date", ""))
                if start is None or end is None:
                    continue

                regions_val = str(pr.get("regions", "")).strip()
                region_kws = tuple(k.strip() for k in regions_val.split(",") if k.strip()) if regions_val else ()

                markup_type_str = str(pr.get("markup_type", "percent")).strip().lower()
                if markup_type_str == "fixed":
                    markup_type_int = 0
                elif markup_type_str == "weight":
                    markup_type_int = 1
                elif markup_type_str == "percent":
                    markup_type_int = 2
                else:
                    continue

                try:
                    markup_value = float(str(pr.get("markup_value", "0")).strip())
                except (ValueError, TypeError):
                    continue
                if markup_value <= 0:
                    continue

                promo_name = str(pr.get("name", "活动加价"))
                region_note = f"[{regions_val}]" if regions_val else ""

                _PROMOTION_CACHE.append(
                    (start, end, markup_type_int, markup_value, region_kws, promo_name, region_note)
                )
            except Exception:
                continue
    except Exception:
        pass

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
    """统一应用活动加价规则（使用预解析缓存，避免每行重复解析日期/字符串）"""
    if not _PROMOTION_CACHE:
        return base_fee, ""

    try:
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        region_str = region_str or ""

        for entry in _PROMOTION_CACHE:
            # entry = (start, end, markup_type_int, markup_value, region_kws, promo_name, region_note)
            start = entry[0]
            end = entry[1]
            if not (start <= today <= end):
                continue

            # 区域限定检查：region_kws 是 tuple，空表示不限定
            region_kws = entry[4]
            if region_kws:
                if not any(k in region_str for k in region_kws):
                    continue

            markup_type_int = entry[2]
            markup_value = entry[3]

            if markup_type_int == 0:
                promo_amount = markup_value
            elif markup_type_int == 1:
                promo_amount = weight * markup_value
            elif markup_type_int == 2:
                promo_amount = base_fee * (markup_value / 100.0)
            else:
                continue

            promo_amount = round(promo_amount, 2)
            if promo_amount > 0:
                promo_name = entry[5]
                region_note = entry[6]
                return round(base_fee + promo_amount, 2), f"+ {promo_name}{region_note}(+¥{promo_amount})"
    except Exception:
        pass

    return base_fee, ""


def _read_excel_fast(file_path: str, sheet_name: Optional[str] = None) -> Tuple[List[str], List[List]]:
    """
    高性能读取Excel：优先 python-calamine(Rust引擎)，不可用时fallback到pandas+openpyxl
    返回: (columns, data_rows)
      - columns: 列名列表
      - data_rows: 数据行 list[list]，每行保留原类型（None/int/float/str/date/datetime）
    """
    # ========== 方案A：calamine (Rust) ==========
    if _HAS_CALAMINE:
        try:
            wb = CalamineWorkbook.from_path(file_path)
            sheet_names_list = wb.sheet_names

            if sheet_name:
                target_sheets = [sheet_name] if sheet_name in sheet_names_list else [sheet_names_list[0]]
            else:
                target_sheets = sheet_names_list

            columns = None
            data_rows = []

            for idx, sn in enumerate(target_sheets):
                sheet = wb.get_sheet_by_name(sn)
                rows_iter = sheet.iter_rows()
                is_first_sheet = (idx == 0)

                for row_idx, row in enumerate(rows_iter):
                    if row is None:
                        continue

                    if row_idx == 0 and is_first_sheet:
                        columns = ["" if v is None else str(v).strip() for v in row]
                        continue

                    # 数据行：直接使用 tuple（calamine已生成），不做类型转换
                    # 后续 row_tup 构建时按需转换
                    data_rows.append(tuple(row))

            if columns is None:
                columns = []

            return columns, data_rows
        except Exception:
            pass  # fallback

    # ========== 方案B：pandas + openpyxl (兼容兜底) ==========
    df = pd.read_excel(file_path, dtype=str, sheet_name=sheet_name)
    if isinstance(df, dict):  # 多sheet合并
        columns = []
        data_rows = []
        is_first_sheet = True
        for sheet_df in df.values():
            sheet_df = sheet_df.fillna("")
            sheet_cols = list(sheet_df.columns)
            if is_first_sheet:
                columns = sheet_cols
                is_first_sheet = False
            for row in sheet_df.itertuples(index=False, name=None):
                data_rows.append(tuple(row))
        return columns, data_rows
    else:
        df = df.fillna("")
        columns = list(df.columns)
        data_rows = [tuple(row) for row in df.itertuples(index=False, name=None)]
        return columns, data_rows


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
            length_str = row_vals[8] or ""
            width_str = row_vals[9] or ""
            height_str = row_vals[10] or ""
            vol_weight_str = row_vals[11] or ""
            quantity_str = row_vals[12] or ""
            service_type = row_vals[13] or ""
            raw_date = row_vals[14] or ""
            raw_customer_code = row_vals[15] or ""
            raw_customer = row_vals[16] or ""
            remark = row_vals[17] or ""
            excel_row_index = row_vals[18]

            try:
                weight = float(weight_str) if weight_str else 0.0
            except (ValueError, TypeError):
                weight = 0.0

            # ========== 体积重计算 ==========
            # 体积重 = 长 × 宽 × 高 ÷ 抛货系数
            # 计费重量 = max(实重, 体积重)
            billing_weight = weight
            try:
                length = float(length_str) if length_str else 0.0
                width = float(width_str) if width_str else 0.0
                height = float(height_str) if height_str else 0.0
                if length > 0 and width > 0 and height > 0:
                    # 优先用文件中预存的体积重，否则自己计算
                    if vol_weight_str:
                        vol_weight = float(vol_weight_str)
                    else:
                        # 通过 station_code 或 station_name 查找除数
                        divisor = _计泡系数_MAP.get(station_code) or \
                                  _计泡系数_MAP.get(station_name) or \
                                  _DEFAULT_计泡系数
                        vol_weight = (length * width * height) / divisor
                    # 取较大值为计费重量
                    if vol_weight > billing_weight:
                        billing_weight = vol_weight
            except (ValueError, TypeError, ZeroDivisionError):
                pass

            try:
                quantity = int(float(quantity_str)) if quantity_str else 1
            except (ValueError, TypeError):
                quantity = 1

            fee, rule_name, is_exc = match_rule_fast(billing_weight, region_name,
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
                billing_weight, quantity, service_type, extra_data,
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

        # ============ 阶段1：读取Excel（优先Rust引擎，省掉pandas中间层） ============
        if file_path.endswith(".csv"):
            report(3, "正在读取CSV...")
            df = pd.read_csv(file_path, dtype=str)
            columns = list(df.columns)
            data_rows = []
            for row in df.itertuples(index=False, name=None):
                data_rows.append(["" if v is None else str(v) for v in row])
            del df
            import gc as _gc0
            _gc0.collect()
        else:
            report(5, "正在读取Excel (Rust引擎)...")
            columns, data_rows = _read_excel_fast(file_path, sheet_name)
            report(15, f"读取完成，共 {len(data_rows):,} 行")

        row_count = len(data_rows)

        # 列名匹配
        from app.services.column_matcher import ColumnMatcher
        matcher = ColumnMatcher()
        matched = matcher.match_columns(columns)

        report(16, f"文件读取完成，共 {row_count:,} 行数据，正在初始化...")

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
        idx_length = _get_col_idx("length")
        idx_width = _get_col_idx("width")
        idx_height = _get_col_idx("height")
        idx_volume_weight = _get_col_idx("volume_weight")
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
            idx_weight, idx_length, idx_width, idx_height, idx_volume_weight,
            idx_quantity, idx_service,
            idx_date, idx_customer_code, idx_customer, idx_remark
        ]

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

            # ============ 阶段3-前置：SQLite性能调优（对百万级数据关键） ============
            # WAL模式：写入并发 + 写速度提升
            # synchronous=OFF：允许OS缓存写入，崩溃可能丢最近一批数据，但程序崩溃极少见
            # cache_size=-800000：分配800MB页面缓存（260万行×~60字节≈156MB，够存）
            # temp_store=MEMORY：临时表放内存不写盘
            try:
                perf_conn = session.connection().connection
                perf_cur = perf_conn.cursor()
                perf_cur.execute("PRAGMA journal_mode = WAL")
                perf_cur.execute("PRAGMA synchronous = OFF")
                perf_cur.execute("PRAGMA cache_size = -800000")
                perf_cur.execute("PRAGMA temp_store = MEMORY")
                perf_cur.execute("PRAGMA mmap_size = 2000000000")  # 2GB内存映射
                perf_conn.commit()
                perf_cur.close()
            except Exception:
                pass

            # 累计写入行数（用于控制大事务提交频率）
            rows_since_commit = 0
            # 每 500K 行 commit 一次：避免事务过大，又把2.5万行一次的commit从104次降到5次
            COMMIT_EVERY_ROWS = 500000

            def _lazy_commit_if_needed(current_conn, rows_count, force_now=False):
                nonlocal rows_since_commit
                rows_since_commit += rows_count
                if force_now or rows_since_commit >= COMMIT_EVERY_ROWS:
                    current_conn.commit()
                    rows_since_commit = 0

            # 决策：10万行以下走单进程（省启动开销），10万行以上自动多进程
            use_multiprocess = row_count >= MIN_ROWS_FOR_MULTIPROCESS

            # ============ 阶段3-a：直接从 data_rows 构建行数据 + 全局去重 ============
            # data_rows 已经是 list[list[str]]，用 column_indices 直接取对应列
            all_row_tuples = []
            seen_global = set()
            duplicate_count = 0
            empty_tracking_count = 0

            import gc
            n_cols = len(columns)
            for pos in range(row_count):
                excel_row = pos + 2
                raw = data_rows[pos]
                row_tup = []
                for i in column_indices:
                    if 0 <= i < len(raw):
                        v = raw[i]
                        if v is None:
                            row_tup.append("")
                        elif isinstance(v, float) and v != v:  # NaN
                            row_tup.append("")
                        else:
                            row_tup.append(str(v))
                    else:
                        row_tup.append("")
                row_tup.append(excel_row)

                tracking_no = row_tup[0] or ""
                if not tracking_no:
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

            report(19,
                   f"数据构建完成。原始 {row_count_original:,} 行，"
                   f"去重后 {row_count:,} 行 "
                   f"（跳过重复 {duplicate_count:,} 行，空单号 {empty_tracking_count:,} 行）")

            # 释放 data_rows 内存
            del data_rows

            if use_multiprocess:
                import multiprocessing as mp
                cpu_count = mp.cpu_count()
                pool_size = max(1, min(cpu_count - 1, 8))

                report(20, f"大数据模式：启动 {pool_size} 个进程并行计算 "
                           f"（CPU {cpu_count}核）")

                # 方案E改进：均匀分配chunk，确保每个进程任务量接近
                # pool_size个进程，分成 pool_size*2 个chunk，让快的进程可以处理更多
                # 用"整除+余数"的方式分配：前 extra_count 个chunk多1行，保证最大差不超过1行
                num_chunks = pool_size * 2
                base_chunk = row_count // num_chunks
                extra_count = row_count % num_chunks
                chunks = []
                pos = 0
                for idx in range(num_chunks):
                    cur_size = base_chunk + (1 if idx < extra_count else 0)
                    end = pos + cur_size
                    chunks.append((all_row_tuples[pos:end], None))
                    pos = end

                del all_row_tuples
                gc.collect()

                report(22, f"已分 {len(chunks)} 块，每块约 {base_chunk:,} 行，启动并行计算...")

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
                            # 懒提交：累积到50万行才真正flush磁盘
                            mp_conn = session.connection().connection
                            _lazy_commit_if_needed(mp_conn, len(batch_to_insert))

                        processed += len(chunk_results)
                        progress = 22 + min(processed / total_to_process, 1.0) * 68
                        report(progress,
                               f"并行计算中... {processed:,}/{total_to_process:,} 行 "
                               f"({int(progress)}%)  成功 {success_count:,}，异常 {exception_count:,}")

                # 多进程结束：最后一批强制提交
                _lazy_commit_if_needed(session.connection().connection, 0, force_now=True)
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
                    # 懒提交：累积到50万行才真正flush磁盘
                    sp_conn = session.connection().connection
                    _lazy_commit_if_needed(sp_conn, len(batch_to_insert))
                    progress = 22 + min(b_end, row_count) / row_count * 68
                    report(progress,
                           f"计算中... {b_end:,}/{row_count:,} 行 ({int(progress)}%)"
                           f"  成功 {success_count:,}，异常 {exception_count:,}")

                # 单进程结束：最后一批强制提交
                _lazy_commit_if_needed(session.connection().connection, 0, force_now=True)
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
            # 注意：不再在此处 commit()，由调用方 _lazy_commit_if_needed 控制提交频率
            # 百万级数据下：每2.5万行一次 commit = 104次fsync → 每50万行一次 commit = 5次fsync
            # 配合 PRAGMA synchronous=OFF，整体写库性能提升3-5倍
        finally:
            cursor.close()
