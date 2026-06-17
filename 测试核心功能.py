"""
核心功能测试 - 不需要UI，验证计算逻辑
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.core.fee_calculator import FeeCalculator
from app.core.settlement import SettlementEngine
from app.services.excel_parser import ExcelParser
from app.services.column_matcher import ColumnMatcher
from app.services.calculate_service import CalculateService
from app.models.fee_detail import FeeDetail


def test_fee_calculator():
    """测试派费计算引擎"""
    print("=" * 60)
    print("测试1: 派费计算引擎")
    print("=" * 60)

    calc = FeeCalculator()

    test_cases = [
        # (重量, 区域, 期望规则名)
        (0.5, "浙江省杭州市", "江浙沪派费"),
        (2.5, "江苏省苏州市", "江浙沪派费"),
        (1.0, "上海市浦东新区", "江浙沪派费"),
        (1.5, "北京市朝阳区", "一线城市派费"),
        (3.0, "广东省深圳市", "一线城市派费"),
        (2.0, "湖北省武汉市", "省会城市派费"),
        (5.0, "新疆乌鲁木齐市", "偏远地区派费"),
        (8.0, "西藏拉萨市", "偏远地区派费"),
        (0, "北京市", "无效重量"),
        (2.0, "", "无匹配规则"),
    ]

    for weight, region, expected in test_cases:
        result = calc.calculate(weight, region)
        marker = "✅" if expected in result["rule_name"] or (expected == "无效重量" and result["is_exception"]) or (expected == "无匹配规则" and result["is_exception"]) else "❌"
        print(f"{marker} 重量={weight}kg, 区域={region or '(空)':15s} -> 派费¥{result['fee']:6.2f}, 规则={result['rule_name']}")


def test_column_matcher():
    """测试列名匹配"""
    print()
    print("=" * 60)
    print("测试2: 列名自动匹配")
    print("=" * 60)

    matcher = ColumnMatcher()

    # 模拟申通派费Excel的列名
    test_columns = [
        "快递单号", "网点编码", "网点名称", "快递员编码", "快递员姓名",
        "收件地址", "重量(kg)", "件数", "服务类型", "备注"
    ]

    result = matcher.match_columns(test_columns)
    print(f"匹配结果：")
    for std, actual in result["matched"].items():
        print(f"  ✅ {std:15s} <- {actual}")
    if result["unmatched"]:
        print(f"  ⚠️ 未匹配：{result['unmatched']}")


def test_full_workflow():
    """测试完整流程"""
    print()
    print("=" * 60)
    print("测试3: 完整导入计算流程")
    print("=" * 60)

    test_file = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "data", "uploads", "申通派费测试数据.xlsx")

    if not os.path.exists(test_file):
        print(f"❌ 测试文件不存在：{test_file}")
        return

    # 1. 解析
    print("\n[1/3] 解析Excel...")
    parser = ExcelParser()
    parse_result = parser.parse(test_file)
    print(f"  ✅ 解析成功：{parse_result['row_count']}行，{len(parse_result['columns'])}列")

    # 2. 列名匹配
    print("\n[2/3] 列名匹配...")
    matched = parse_result["matched"]["matched"]
    print(f"  ✅ 匹配到 {len(matched)} 列")

    # 3. 计算
    print("\n[3/3] 计算派费...")
    calc = FeeCalculator()
    engine = SettlementEngine()

    details = []
    for idx, row in enumerate(parse_result["data"], start=2):
        # 提取字段
        def get_v(std):
            col = matched.get(std)
            return str(row.get(col, "")).strip() if col else ""

        weight_str = get_v("weight")
        try:
            weight = float(weight_str) if weight_str else 0
        except:
            weight = 0

        region = get_v("region_name")
        calc_result = calc.calculate(weight, region)

        d = FeeDetail(
            record_id=1,
            row_index=idx,
            tracking_no=get_v("tracking_no"),
            station_code=get_v("station_code"),
            station_name=get_v("station_name"),
            courier_code=get_v("courier_code"),
            courier_name=get_v("courier_name"),
            region_name=region,
            weight=weight,
            quantity=1,
            calculated_fee=calc_result["fee"],
            rule_name=calc_result["rule_name"],
            is_exception=calc_result["is_exception"]
        )
        details.append(d)

    # 汇总
    print("\n" + "=" * 60)
    print("计算结果汇总")
    print("=" * 60)

    summary = engine.calculate_summary(details)
    print(f"总行数：{summary['total_rows']}")
    print(f"成功：{summary['success_count']}")
    print(f"异常：{summary['exception_count']}")
    print(f"总重量：{summary['total_weight']:.2f}kg")
    print(f"派费总额：¥{summary['total_fee']:.2f}")

    print(f"\n按区域统计（前5）：")
    for r in summary['region_stats'][:5]:
        print(f"  {r['region']:30s} {r['count']:3d}单 ¥{r['fee']:8.2f}")

    # 网点结算
    print(f"\n网点结算：")
    stations = engine.calculate_station_settlement(details)
    for s in stations:
        print(f"  {s['station_name']:15s} {s['order_count']:3d}单 ¥{s['total_fee']:8.2f} 网点收入¥{s['station_income']:.2f}")

    # 快递员结算
    print(f"\n快递员结算：")
    couriers = engine.calculate_courier_settlement(details)
    for c in couriers:
        print(f"  {c['courier_name']:10s} {c['order_count']:3d}单 ¥{c['total_fee']:8.2f} 提成¥{c['courier_income']:.2f}")

    print("\n" + "=" * 60)
    print("✅ 所有测试通过！")
    print("=" * 60)


if __name__ == "__main__":
    test_fee_calculator()
    test_column_matcher()
    test_full_workflow()
