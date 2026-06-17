# 申通派费计算系统

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 运行程序

```bash
python main.py
```

### 3. 打包为exe（可选）

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name="申通派费计算系统" main.py
```

生成的exe在 `dist/` 目录下。

---

## 使用说明

### 第一步：准备Excel
确保你的Excel包含以下关键列（任一别名即可）：

| 字段 | 可识别的列名 | 必填 |
|------|-------------|------|
| 快递单号 | 快递单号/单号/运单号/订单号 | 否 |
| **重量** | 重量/kg/公斤/实重 | ✅ |
| **区域** | 收件地址/区域/目的地/省份 | ✅ |
| 件数 | 件数/数量/箱数 | 否 |
| 网点编码 | 网点编码/网点代码/站点编码 | 否 |
| 网点名称 | 网点名称/网点/站点 | 否 |
| 快递员 | 快递员/派件员/工号 | 否 |
| 备注 | 备注/说明 | 否 |

### 第二步：导入计算
1. 点击"选择Excel文件"
2. 系统自动识别列名（支持模糊匹配）
3. 点击"开始计算"
4. 计算完成后自动跳转到结果页

### 第三步：多级结算
- 切换到"多级结算"Tab
- 选择"网点结算"或"快递员结算"
- 一键导出结算单

---

## 项目结构

```
excelbest/
├── main.py                       # 启动入口
├── requirements.txt              # 依赖
├── app/
│   ├── core/                     # 核心业务
│   │   ├── fee_calculator.py    # 派费计算引擎
│   │   └── settlement.py        # 多级结算引擎
│   ├── services/                 # 服务层
│   │   ├── excel_parser.py      # Excel解析
│   │   ├── column_matcher.py    # 列名自动匹配
│   │   ├── calculate_service.py # 计算服务
│   │   └── export_service.py    # 导出服务
│   ├── models/                   # 数据模型
│   │   ├── database.py          # 数据库连接
│   │   ├── fee_record.py        # 计算记录
│   │   ├── fee_detail.py        # 派费明细
│   │   ├── station.py           # 网点
│   │   ├── courier.py           # 快递员
│   │   ├── commission_rule.py   # 分成规则
│   │   └── column_mapping.py    # 列名映射
│   └── ui/
│       └── main_window.py       # 主窗口
└── data/                        # 数据目录
    ├── app.db                   # SQLite数据库
    ├── uploads/                 # 上传文件
    └── exports/                 # 导出文件
```

---

## 派费规则自定义

编辑 `app/core/fee_calculator.py` 中的 `default_rules`：

```python
{
    "name": "江浙沪派费",
    "priority": 1,
    "regions": ["上海", "江苏", "浙江"],
    "first_weight": 1.0,        # 首重1kg
    "first_fee": 3.0,           # 首重3元
    "continued_fee": 1.0,       # 续重1元/kg
    "min_fee": 3.0              # 保底价3元
}
```

支持维度：
- `priority`: 优先级（数字越小越优先匹配）
- `regions`: 适用区域列表
- `first_weight`: 首重（kg）
- `first_fee`: 首重费用（元）
- `continued_fee`: 续重费用（元/kg）
- `min_fee`: 保底价（元）

---

## 常见问题

**Q1: Excel列名识别不到？**
- 在"列名映射"中手动添加别名（编辑 `app/services/column_matcher.py` 的 `DEFAULT_MAPPINGS`）
- 或在数据库的 `column_mappings` 表中添加

**Q2: 派费算错了？**
- 检查 `app/core/fee_calculator.py` 的规则配置
- 确认区域列是否正确识别

**Q3: 想重置数据库？**
- 删除 `data/app.db` 文件
- 重新运行程序
