"""
列名自动匹配服务
根据预设的列名别名，识别Excel中的列
适配申通快递派费Excel的常见列名
"""
from typing import Dict, List
from app.models.database import get_session
from app.models.column_mapping import ColumnMapping


# 申通快递运费结算Excel常见列名（与系统预设对应）
DEFAULT_MAPPINGS = {
    "tracking_no": ["快递单号", "单号", "运单号", "订单号", "tracking", "tracking_no", "运单编号", "条形码"],
    "station_code": ["网点编码", "网点代码", "网点号", "站点编码", "站点代码", "派件网点", "派送网点", "分部编码"],
    "station_name": ["网点名称", "网点", "站点", "分部", "派件网点名称", "派送网点名称"],
    "courier_code": ["快递员编码", "快递员代码", "快递员编号", "工号", "派件员编码", "派件员编号"],
    "courier_name": ["快递员姓名", "快递员", "派件员", "业务员", "收派员"],
    "region_code": ["区域编码", "目的地编码", "省份编码", "城市编码"],
    "region_name": ["目的省份", "目的省", "目的地省份", "收件地址", "区域", "目的地", "派送区域", "省份", "省市区", "目的地省市区"],
    "weight": ["重量", "重量(kg)", "重量(公斤)", "公斤", "实重", "计费重量", "实际重量", "结算重量", "称重重量"],
    "length": ["长", "长度", "长(cm)", "长(CM)", "尺寸长", "件长", "包装长"],
    "width": ["宽", "宽度", "宽(cm)", "宽(CM)", "尺寸宽", "件宽", "包装宽"],
    "height": ["高", "高度", "高(cm)", "高(CM)", "尺寸高", "件高", "包装高"],
    "volume_weight": ["体积重", "体积重量", "体积重量(kg)", "抛货体积重", "材积", "材积重", "材积重量", "体积(kg)"],
    "quantity": ["件数", "数量", "包裹数", "箱数", "票数", "总件数"],
    "service_type": ["服务类型", "时效类型", "快递类型", "件类型", "产品类型"],
    "sender": ["寄件人", "发件人", "发货人"],
    "receiver": ["收件人", "收货人"],
    "cod_amount": ["代收货款", "货款金额", "代收金额", "COD"],
    "business_date": ["业务日期", "日期", "结算日期", "派件日期", "运单日期", "业务时间", "date", "Date", "DATE"],
    "customer_code": ["客户编码", "客户代码", "客户编号", "客户ID", "customer_code", "customer_code"],
    "customer_name": ["客户名称", "客户", "收件客户", "发件客户", "客户名", "customer", "Customer"],
    "remark": ["备注", "说明", "remark"]
}


class ColumnMatcher:
    """列名匹配器"""

    def __init__(self):
        # 从数据库加载列名映射（如果没有则用默认）
        self.mappings = self._load_mappings()

    def _load_mappings(self) -> Dict:
        """从数据库加载映射，与默认合并（默认值始终生效，DB映射可扩展）"""
        mappings = DEFAULT_MAPPINGS.copy()
        try:
            session = get_session()
            db_mappings = session.query(ColumnMapping).all()
            if db_mappings:
                # DB映射与默认合并：同名字段合并别名列表，保留默认别名 + 扩展DB别名
                for m in db_mappings:
                    if m.standard_name in mappings:
                        existing = [a for a in (m.alias_names or []) if a not in mappings[m.standard_name]]
                        mappings[m.standard_name] = list(mappings[m.standard_name]) + existing
                    else:
                        mappings[m.standard_name] = list(m.alias_names or [])
            session.close()
        except Exception:
            pass
        return mappings

    def match_columns(self, columns: List[str]) -> Dict:
        """
        自动匹配列名
        :param columns: Excel中的列名列表
        :return: {"matched": {标准名: 实际列名}, "unmatched": [未匹配到的Excel列名]}
        """
        matched = {}
        unmatched = []

        for col in columns:
            col_clean = str(col).strip()
            match_result = self._match_single(col_clean)
            if match_result:
                matched[match_result] = col_clean
            else:
                unmatched.append(col_clean)

        return {
            "matched": matched,
            "unmatched": unmatched
        }

    def _match_single(self, column_name: str) -> str:
        """匹配单个列名"""
        col_clean = column_name.strip()
        col_lower = col_clean.lower()

        # 黑名单：含"城市"且不含"省"的列不能匹配为区域（区域=目的省份）
        is_city_column = ("城市" in col_clean and "省" not in col_clean) or "目的城市" in col_clean
        # 黑名单：含"签收"的列不是目的省份
        is_sign_column = "签收" in col_clean

        # 第一轮：精确匹配（优先级最高）
        for standard, aliases in self.mappings.items():
            for alias in aliases:
                if col_lower == alias.lower():
                    # 应用黑名单：region_name不允许是城市或签收相关
                    if standard == "region_name" and (is_city_column or is_sign_column):
                        continue
                    return standard

        # 第二轮：包含匹配 - alias的关键词必须出现在列名中，且最少3个字符
        for standard, aliases in self.mappings.items():
            for alias in aliases:
                alias_lower = alias.lower().strip()
                # 最少3个字符才有模糊匹配意义（避免"kg"匹配"体积重(kg)"等误匹配）
                if len(alias_lower) >= 3 and alias_lower in col_lower:
                    if standard == "region_name" and (is_city_column or is_sign_column):
                        continue
                    return standard

        return None

    def init_default_mappings(self):
        """初始化默认列名映射到数据库"""
        session = get_session()
        try:
            # 检查是否已初始化
            count = session.query(ColumnMapping).count()
            if count > 0:
                return False

            for standard, aliases in DEFAULT_MAPPINGS.items():
                mapping = ColumnMapping(
                    standard_name=standard,
                    alias_names=aliases,
                    is_required=(standard in ["region_name", "weight"]),
                    data_type="string"
                )
                session.add(mapping)
            session.commit()
            return True
        except Exception as e:
            session.rollback()
            raise e
        finally:
            session.close()
