"""
申通派费计算引擎
参考申通快递行业通用规则（可配置）
"""
from typing import Dict, Optional
from decimal import Decimal, ROUND_HALF_UP
import json
import os


class FeeCalculator:
    """派费计算器 - 核心计算逻辑"""

    def __init__(self, rules_file: Optional[str] = None):
        """
        初始化计算器
        :param rules_file: 规则配置文件路径
        """
        # 申通派费行业通用规则（默认配置，可在系统里修改）
        self.default_rules = {
            "version": "1.0",
            "company": "申通快递",
            "description": "申通快递行业通用派费规则（可修改）",
            "rules": [
                # 江浙沪 - 价格最低
                {
                    "name": "江浙沪派费",
                    "priority": 1,
                    "regions": ["上海", "江苏", "浙江", "杭州", "苏州", "南京", "宁波", "温州"],
                    "first_weight": 1.0,        # 首重1kg
                    "first_fee": 3.0,           # 首重3元
                    "continued_fee": 1.0,       # 续重1元/kg
                    "min_fee": 3.0
                },
                # 京津冀 + 广东 + 福建
                {
                    "name": "一线城市派费",
                    "priority": 2,
                    "regions": ["北京", "天津", "河北", "广州", "深圳", "东莞", "佛山", "福建", "厦门", "福州"],
                    "first_weight": 1.0,
                    "first_fee": 4.0,
                    "continued_fee": 1.5,
                    "min_fee": 4.0
                },
                # 其他省会城市
                {
                    "name": "省会城市派费",
                    "priority": 3,
                    "regions": ["成都", "重庆", "武汉", "长沙", "合肥", "南昌", "济南", "青岛", "郑州", "西安",
                                "太原", "石家庄", "沈阳", "大连", "长春", "哈尔滨", "昆明", "贵阳", "南宁", "海口"],
                    "first_weight": 1.0,
                    "first_fee": 5.0,
                    "continued_fee": 2.0,
                    "min_fee": 5.0
                },
                # 偏远地区
                {
                    "name": "偏远地区派费",
                    "priority": 99,
                    "regions": ["新疆", "西藏", "青海", "内蒙古", "甘肃", "宁夏", "海南"],
                    "first_weight": 1.0,
                    "first_fee": 10.0,
                    "continued_fee": 5.0,
                    "min_fee": 10.0
                }
            ],
            "default_rule": {
                # 未匹配区域的默认规则
                "name": "其他地区派费",
                "first_weight": 1.0,
                "first_fee": 6.0,
                "continued_fee": 2.5,
                "min_fee": 6.0
            }
        }
        self.rules = self.default_rules

    def calculate(self, weight: float, region: str = "") -> Dict:
        """
        计算派费
        :param weight: 重量（kg）
        :param region: 区域
        :return: {"fee": 派费金额, "rule_name": 命中的规则名, "is_exception": 是否异常}
        """
        # 异常：重量为空或0
        if weight is None or weight <= 0:
            return {
                "fee": 0.0,
                "rule_name": "无效重量",
                "is_exception": True,
                "remark": f"重量无效: {weight}"
            }

        # 找到匹配的规则
        rule = self._match_rule(region)
        if not rule:
            return {
                "fee": 0.0,
                "rule_name": "无匹配规则",
                "is_exception": True,
                "remark": f"区域[{region}]未匹配到规则"
            }

        # 计算派费
        first_weight = float(rule.get("first_weight", 1.0))
        first_fee = float(rule.get("first_fee", 5.0))
        continued_fee = float(rule.get("continued_fee", 2.0))
        min_fee = float(rule.get("min_fee", first_fee))

        weight = float(weight)

        if weight <= first_weight:
            # 未超首重
            fee = first_fee
        else:
            # 超过首重，按续重计费
            continued_weight = weight - first_weight
            fee = first_fee + continued_weight * continued_fee

        # 保底价
        fee = max(fee, min_fee)

        # 四舍五入到分
        fee = float(Decimal(str(fee)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

        return {
            "fee": fee,
            "rule_name": rule.get("name", "默认规则"),
            "is_exception": False,
            "remark": ""
        }

    def _match_rule(self, region: str):
        """根据区域匹配规则"""
        if not region:
            return self.rules.get("default_rule")

        region = str(region).strip()

        # 遍历所有规则，按优先级匹配
        sorted_rules = sorted(self.rules.get("rules", []), key=lambda r: r.get("priority", 999))
        for rule in sorted_rules:
            rule_regions = rule.get("regions", [])
            for r in rule_regions:
                if r in region or region in r:
                    return rule

        # 未匹配到，使用默认规则
        return self.rules.get("default_rule")

    def save_rules(self, file_path: str):
        """保存规则到文件"""
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(self.rules, f, ensure_ascii=False, indent=2)

    def load_rules(self, file_path: str):
        """从文件加载规则"""
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="utf-8") as f:
                self.rules = json.load(f)
            return True
        return False
