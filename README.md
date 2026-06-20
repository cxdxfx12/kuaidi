# 大圣 · 快递物流派费结算系统

<div align="center">

**杭州喵喵至家网络有限公司 · 大圣智慧软件**

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![PyQt5](https://img.shields.io/badge/GUI-PyQt5-green)
![License](https://img.shields.io/badge/License-Internal-orange)
![Platform](https://img.shields.io/badge/Platform-Win%20%7C%20macOS-lightgrey)

</div>

---

## 项目简介

面向快递公司网点 / 区域 / 客户层级的派费结算系统，从 Excel 导入快递明细 → 按规则自动计算 → 生成多级结算报告 → 历史记录可随时查询和导出。

核心特性：

- **拖拽导入**：把 Excel 文件拖进窗口即可开始计算
- **四级规则匹配**：网点 → 门店 → 区域 → 全局（优先级从高到低）
- **差异化计泡**：顺丰 6000 / 圆通 8000 / EMS 5000 / 大件轻抛 12000，自动按快递公司匹配
- **活动加价**：双十一、高峰期按业务日期和省份自动加价
- **多进程计算**：10 万行以上自动开启多进程加速
- **快速导出**：一键导出 Excel 报告

---

## 目录结构

```
excelbest/
├── main.py                        # 程序入口
├── 大圣派费结算系统.spec            # PyInstaller 打包配置
├── build_ico.py                   # 图标生成脚本
├── requirements.txt               # Python 依赖
├── 结算软件开发.md                # 开发记录（含本次对话）
├── app/
│   ├── core/
│   │   └── fee_calculator.py      # 保底费 + 续重计算逻辑
│   ├── models/
│   │   ├── database.py            # 数据库初始化
│   │   ├── fee_detail.py          # 明细表模型
│   │   ├── fee_record.py          # 汇总记录模型
│   │   └── path_config.py         # 路径 & 应用数据目录
│   ├── services/
│   │   ├── calculate_service.py   # 计算服务（多进程、计泡匹配）
│   │   ├── column_matcher.py      # Excel 列名自动匹配
│   │   └── rule_service.py        # 规则加载 / 序列化
│   └── ui/
│       └── main_window.py         # PyQt5 主界面（含拖拽区域）
└── data/
    ├── config/
    │   └── fee_rules.json         # 规则配置 + 计泡系数_MAP
    └── icons/
        ├── monkey-icon.png        # 原始图标源文件
        └── dasheng.ico            # 多尺寸 ICO（16/24/32/48/64/128/256）
```

---

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 直接运行源码

```bash
python main.py
```

### 3. 打包为 Windows EXE

```bash
pyinstaller --noconfirm --clean "大圣派费结算系统.spec"
# 产物位置：dist/大圣派费结算系统.exe
```

### 4. 打包为 macOS DMG

参考仓库内 `build_mac.py` / `macos_remote_build.py` 脚本。

---

## 核心计算公式

### 计费重量

```
体积重 = 长 × 宽 × 高 ÷ 计泡系数
计费重量 = max(实重, 体积重)
```

### 费用

```
if 计费重量 ≤ 首重:
    费用 = max(首重费用, 保底费)
else:
    费用 = 首重费用 + (计费重量 - 首重) × 续重单价
```

活动期间额外加价按 `fee_rules.json` 中的活动规则叠加计算。

### 计泡系数按快递公司差异化

| 快递公司 | 计泡系数 |
|---------|---------|
| 顺丰 / SF / 京东 / 德邦 / 跨越速运 | **6000** |
| 圆通 / 中通 / 韵达 / 申通 / 极兔 / 天天 / 百世 | **8000** |
| EMS / 中国邮政 | **5000** |
| 优速 / 安能 / 天地华宇（大件轻抛） | **12000** |
| 未知/未匹配 | 用户设置的全局默认值 |

详见 [`data/config/fee_rules.json`](file:///e:/excelbest/data/config/fee_rules.json)。

---

## 规则配置说明

规则文件位置：`数据目录/fee_rules.json`（Windows 下通常是 `%APPDATA%/dasheng/config/fee_rules.json`，首次启动会从打包资源自动复制）

每个规则支持：

- `outlet`：门店名（精确匹配）
- `stations`：网点名（逗号分隔，支持中英文逗号）
- `regions`：省份（逗号分隔，支持中英文逗号）
- `first_weight_fee`：首重费用
- `continued_fee_per_kg`：续重单价
- `min_fee`：保底费
- `weight_round`：重量进位模式
- `promotions`：活动加价规则（按**业务日期**判断，包含 `start_date`/`end_date`/`provinces`/`percent`）

全局默认值可在 UI 中「规则配置 → 全局默认设置」页面调整。

---

## 性能

| 数据量 | 预估耗时 |
|--------|---------|
| 10,000 行 | 2–5 秒 |
| 50,000 行 | 5–12 秒 |
| 100,000 行 | 10–20 秒 |
| 500,000 行 | 40–80 秒 |

关键优化：

1. **python-calamine（Rust 引擎）** 读取 Excel，提速 3–5 倍
2. **ProcessPoolExecutor** 多进程并行计算（10 万行以上自动开启）
3. **SQLite WAL 模式 + 5000 行批量写入**
4. **计泡系数 / 规则索引预构建**，O(1) 查找

---

## 技术栈

| 模块 | 技术 |
|-----|-----|
| GUI | PyQt 5.15.9（Qt 5.15.2 LTS） |
| 数据处理 | pandas 1.5.3 + openpyxl 3.1.2 + python-calamine |
| 数据库 | SQLAlchemy 2.0.23 + SQLite |
| 打包 | PyInstaller 5.13.0（单文件 onefile 模式） |
| 开发工具 | Trae / Solo IDE |

---

## 开发记录

完整的开发历程、优化细节、问题排查记录见 [`结算软件开发.md`](file:///e:/excelbest/结算软件开发.md)。

---

## 常见问题

### Q1：修改了"全局默认设置 → 计泡系数"但计算结果没变？

请重新点击「导入并计算」让规则重载。另外注意：**如果快递公司名称能匹配到 `计泡系数_MAP`，将使用该映射值，不受全局默认影响**。全局默认值只在快递公司完全匹配不到时生效。

### Q2：活动加价规则设置了但没生效？

检查两点：
1. 活动期间是否包含快递单的**业务日期**（不是系统当前日期）
2. 限定省份名称是否正确匹配 — 中英文逗号都支持，但"浙江"和"浙江省"是不同的匹配目标

### Q3：Windows 桌面图标模糊？

清理 Windows 图标缓存后重新查看，或把 exe 换到不同目录测试。

### Q4：Excel 文件拖进去没反应？

确认文件扩展名为 `.xlsx` / `.xls` / `.csv`，并尝试点击右侧「选择 Excel 文件」按钮作为替代入口（两种方式等价）。

---

## License

内部项目 · 杭州喵喵至家网络有限公司
