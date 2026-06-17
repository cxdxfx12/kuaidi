"""
多级结算引擎
网点结算、快递员结算、承包区结算、月结客户结算
"""
from decimal import Decimal
from collections import defaultdict
from typing import Dict, List
from app.models.fee_detail import FeeDetail


class SettlementEngine:
    """结算引擎"""

    def calculate_station_settlement(self, details: List[FeeDetail]) -> List[Dict]:
        """
        按网点汇总结算
        :return: [{"station_code": "X", "station_name": "X", "total_fee": 0, "commission": 0, "order_count": 0}, ...]
        """
        station_map = defaultdict(lambda: {
            "station_code": "",
            "station_name": "",
            "total_fee": Decimal("0"),
            "order_count": 0,
            "weight_sum": Decimal("0")
        })

        for d in details:
            if d.is_exception:
                continue
            # 以"网点编码+网点名称"做分组key，保证同名不同编码/同编码不同名称不被错误合并
            sc = d.station_code or ""
            sn = d.station_name or ""
            if not sc and not sn:
                sc = "未知"
                sn = "未知网点"
            elif not sc:
                sc = sn
            elif not sn:
                sn = sc
            key = (sc, sn)
            s = station_map[key]
            s["station_code"] = sc
            s["station_name"] = sn
            s["total_fee"] += Decimal(str(d.calculated_fee or 0))
            s["order_count"] += 1
            s["weight_sum"] += Decimal(str(d.weight or 0))

        # 计算分成（默认15%，可在系统配置）
        result = []
        for (code, name), info in station_map.items():
            commission_rate = Decimal("0.15")
            commission = info["total_fee"] * commission_rate
            result.append({
                "station_code": info["station_code"],
                "station_name": info["station_name"],
                "order_count": info["order_count"],
                "total_weight": float(info["weight_sum"]),
                "total_fee": float(info["total_fee"]),
                "commission_rate": float(commission_rate),
                "station_income": float(commission),
                "company_income": float(info["total_fee"] - commission)
            })

        # 按派费降序
        result.sort(key=lambda x: x["total_fee"], reverse=True)
        return result

    def calculate_courier_settlement(self, details: List[FeeDetail]) -> List[Dict]:
        """
        按快递员汇总结算
        """
        courier_map = defaultdict(lambda: {
            "courier_code": "",
            "courier_name": "",
            "station_code": "",
            "total_fee": Decimal("0"),
            "order_count": 0,
            "weight_sum": Decimal("0")
        })

        for d in details:
            if d.is_exception:
                continue
            cc = d.courier_code or ""
            cn = d.courier_name or ""
            if not cc and not cn:
                cc = "未知"
                cn = "未知快递员"
            elif not cc:
                cc = cn
            elif not cn:
                cn = cc
            key = (cc, cn)
            c = courier_map[key]
            c["courier_code"] = cc
            c["courier_name"] = cn
            c["station_code"] = d.station_code or ""
            c["total_fee"] += Decimal(str(d.calculated_fee or 0))
            c["order_count"] += 1
            c["weight_sum"] += Decimal(str(d.weight or 0))

        result = []
        for (code, name), info in courier_map.items():
            # 快递员提成比例默认80%
            commission_rate = Decimal("0.80")
            commission = info["total_fee"] * commission_rate
            result.append({
                "courier_code": info["courier_code"],
                "courier_name": info["courier_name"],
                "station_code": info["station_code"],
                "order_count": info["order_count"],
                "total_weight": float(info["weight_sum"]),
                "total_fee": float(info["total_fee"]),
                "commission_rate": float(commission_rate),
                "courier_income": float(commission),
                "station_share": float(info["total_fee"] - commission)
            })

        result.sort(key=lambda x: x["total_fee"], reverse=True)
        return result

    def calculate_contract_settlement(self, details: List[FeeDetail]) -> List[Dict]:
        """
        按承包区汇总结算
        承包区识别规则：网点编码前3位相同视为同一承包区
        """
        contract_map = defaultdict(lambda: {
            "contract_code": "",
            "contract_name": "",
            "station_codes": set(),
            "total_fee": Decimal("0"),
            "order_count": 0,
            "weight_sum": Decimal("0")
        })

        for d in details:
            if d.is_exception:
                continue
            # 承包区编码取网点编码前3位（如HZ001 -> HZ0）
            station_code = d.station_code or "未知"
            contract_code = station_code[:3] if len(station_code) >= 3 else station_code
            contract_name = f"{contract_code}承包区"

            key = contract_code
            c = contract_map[key]
            c["contract_code"] = contract_code
            c["contract_name"] = contract_name
            c["station_codes"].add(station_code)
            c["total_fee"] += Decimal(str(d.calculated_fee or 0))
            c["order_count"] += 1
            c["weight_sum"] += Decimal(str(d.weight or 0))

        result = []
        for code, info in contract_map.items():
            # 承包区分成比例默认10%
            commission_rate = Decimal("0.10")
            commission = info["total_fee"] * commission_rate
            result.append({
                "contract_code": info["contract_code"],
                "contract_name": info["contract_name"],
                "station_code": ",".join(sorted(info["station_codes"]))[:20],  # 显示涉及的网点
                "order_count": info["order_count"],
                "total_weight": float(info["weight_sum"]),
                "total_fee": float(info["total_fee"]),
                "commission_rate": float(commission_rate),
                "contract_income": float(commission)
            })

        result.sort(key=lambda x: x["total_fee"], reverse=True)
        return result

    def calculate_monthly_settlement(self, details: List[FeeDetail]) -> List[Dict]:
        """
        按月结客户汇总结算
        按"客户编码 + 客户名称"作为唯一键分组（避免仅靠编码造成客户名称乱码/缺失）
        """
        customer_map = defaultdict(lambda: {
            "customer_code": "",
            "customer_name": "",
            "total_fee": Decimal("0"),
            "cod_amount": Decimal("0"),
            "order_count": 0,
            "weight_sum": Decimal("0")
        })

        for d in details:
            if d.is_exception:
                continue
            # 从原始数据中提取客户信息（兼容中文/英文键名）
            original = d.original_data or {}
            customer_code = (
                original.get("客户编码") or original.get("customer_code") or
                original.get("客户代码") or original.get("月结客户") or ""
            )
            customer_name = (
                original.get("客户名称") or original.get("customer_name") or
                original.get("客户名") or ""
            )

            # 规范化：编码/名称至少有一个不为空，否则归类为散单
            cc = str(customer_code).strip()
            cn = str(customer_name).strip()
            if not cc and not cn:
                cc = "散单"
                cn = "散单客户"
            elif not cc:
                # 有名称但无编码：用名称作为分组，编码展示为空
                cc = cn  # 内部用名称做分组，保证独立
            elif not cn:
                cn = cc  # 有编码但无名称：用编码代替名称展示

            # 以"编码+名称"作为分组key，保证不同的组合不会合并在一起
            key = (cc, cn)
            c = customer_map[key]
            c["customer_code"] = cc if cc != cn or cc == "散单" else cc
            c["customer_name"] = cn
            c["total_fee"] += Decimal(str(d.calculated_fee or 0))
            c["order_count"] += 1
            c["weight_sum"] += Decimal(str(d.weight or 0))
            # 代收货款（如有）
            cod = original.get("代收货款", original.get("货款金额", original.get("cod_amount", 0)))
            try:
                c["cod_amount"] += Decimal(str(cod or 0))
            except:
                pass

        result = []
        for (code, name), info in customer_map.items():
            receivable = info["total_fee"] + info["cod_amount"]
            result.append({
                "customer_code": info["customer_code"],
                "customer_name": info["customer_name"],
                "order_count": info["order_count"],
                "total_weight": float(info["weight_sum"]),
                "total_fee": float(info["total_fee"]),
                "cod_amount": float(info["cod_amount"]),
                "status": "待结算",
                "receivable": float(receivable)
            })

        result.sort(key=lambda x: x["total_fee"], reverse=True)
        return result

    def calculate_summary(self, details: List[FeeDetail]) -> Dict:
        """
        总体汇总
        """
        total_fee = Decimal("0")
        success_count = 0
        exception_count = 0
        total_weight = Decimal("0")
        region_stats = defaultdict(lambda: {"fee": Decimal("0"), "count": 0})

        for d in details:
            if d.is_exception:
                exception_count += 1
            else:
                total_fee += Decimal(str(d.calculated_fee or 0))
                total_weight += Decimal(str(d.weight or 0))
                success_count += 1
                # 按区域统计
                region = d.region_name or "未知区域"
                region_stats[region]["fee"] += Decimal(str(d.calculated_fee or 0))
                region_stats[region]["count"] += 1

        return {
            "total_rows": len(details),
            "success_count": success_count,
            "exception_count": exception_count,
            "total_fee": float(total_fee),
            "total_weight": float(total_weight),
            "region_stats": [
                {"region": k, "fee": float(v["fee"]), "count": v["count"]}
                for k, v in sorted(region_stats.items(), key=lambda x: x[1]["fee"], reverse=True)
            ]
        }
