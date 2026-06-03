# 岩土工程勘察报告自动生成工具 v4.0

从华宁勘察软件导出的 Excel 数据 + 正式报告模板 → 一键生成完整岩土工程勘察报告

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 准备项目配置 (复制 project_config.example.json 并修改)
cp project_config.example.json my_project.json

# 3. 运行 (推荐: 配置文件方式)
python generate_report.py --config my_project.json

# 或: 目录自动检测方式 (向后兼容 v3)
python generate_report.py --project "项目目录"

# 4. 输出
# → 项目目录/模板名_正式报告.docx
```

## v4.0 更新要点

- 新增 JSON 配置文件驱动，地层参数/承载力/桩基参数均可配置
- 修复日期硬编码 bug (原 `_update_dates` 固定替换 2025年11月→12月)
- 修复变量名遮蔽问题 (`ss` 局部变量遮蔽模块级函数)
- 引入 logging 框架替代 print，支持 `--verbose` 详细日志
- 表格索引可配置 (`table_indices`)，适配不同模板
- 依赖检查延迟导入，`--help` 在未安装依赖时仍可运行
- 全面添加类型注解，提升代码可读性
- 新增 `requirements.txt` 依赖清单

## 核心功能

- 自动读取华宁导出的 11 类 Excel 数据文件
- 智能识别钻孔类型（取土/标贯/一般/动探/波速）
- A3 双面板水样/盐样 xlsx 自动解析
- GB50021-2001 第12章腐蚀性自动评价
- 表格样式完整保留（只改文字不动格式）
- 配置驱动: 地层描述/承载力/桩基参数均从 JSON 读取

## 已验证项目

| # | 项目名称 | 钻孔数 | 地层数 | 模板来源 |
|---|---------|:-----:|:-----:|---------|
| 1 | 威海高区保税物流中心(B型) | 118 | 12层 | 自建正式报告 |
| 2 | 威海高区电子信息产业园二期 | 111 | 12层 | 复用案例1模板 |

## 命令行参数

| 参数 | 说明 |
|------|------|
| `--config`, `-c` | JSON 配置文件路径 (推荐) |
| `--project`, `-p` | 项目目录路径 (向后兼容) |
| `--template`, `-t` | 指定 .docx 模板 (覆盖配置) |
| `--output`, `-o` | 指定输出路径 (覆盖配置) |
| `--layers` | 地层名称映射 JSON 文件 |
| `--dry-run` | 仅加载数据，不生成报告 |
| `--verbose`, `-v` | 输出详细调试日志 |

## 配置文件说明

详见 `project_config.example.json`，主要配置项:

- `project_name` — 项目名称
- `base_dir` — 项目根目录 (绝对路径或相对于配置文件)
- `excel_dir` — Excel 数据子目录
- `template` — .docx 模板文件名
- `table_indices` — docx 中各表格的序号索引
- `layers` — 地层列表 (含描述模板、承载力参数、桩基参数)
- `replacements` — 全局文本替换规则
- `date_replacements` — 日期替换规则

## 文件说明

| 文件 | 说明 |
|------|------|
| `generate_report.py` | 主程序 (v4.0) |
| `corrosion_eval.py` | 腐蚀性评价独立模块 |
| `project_config.example.json` | 配置文件模板 |
| `requirements.txt` | Python 依赖清单 |
| `README.md` | 本文件 |

## 可移植性

工具可在任意 Windows/Linux 电脑上运行，只需：
1. Python 3.9+
2. `pip install -r requirements.txt`
3. 准备 .docx 模板和 Excel 数据

## 局限

- .doc 旧格式需先转为 .docx
- 地层描述模板适用于滨海/冲洪积地层
- 需华宁软件导出的标准格式 Excel
