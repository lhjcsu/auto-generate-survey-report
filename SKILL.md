---
name: auto-generate-survey-report
description: "从华宁勘察软件导出的Excel数据+正式报告模板，配置驱动一键生成岩土工程勘察报告.docx。含GB50021-2001腐蚀性自动评价。"
version: 4.0
author: GeoEngineer
category: geotechnical
platforms: ["windows", "linux"]
python: ">=3.9"
---

# 自动生成岩土工程勘察报告

从华宁勘察软件导出的 Excel/.XLS/.xlsx 数据自动填充到正式报告模板，配置驱动一键输出完整岩土工程勘察报告。

## 已验证案例

| 项目 | 钻孔 | 地层 | 模板 | 
|------|:--:|:--:|------|
| 威海高区保税物流中心(B型) | 118 | 12层 | 正式报告.docx |
| 威海高区电子信息产业园二期 | 111 | 12层 | 复用保税中心模板 |

## 工作流程

```
项目数据准备:
  ├── project_config.json    ← 项目配置文件 (推荐)
  ├── 报告模板.docx          ← 正式报告模板(完整章节+表格)
  └── Excel数据/             ← 华宁导出的数据文件
      ├── 勘探点一览表*.XLS
      ├── 场地地层厚度统计表*.xls
      ├── 物理力学性质指标统计表*.XLS
      ├── 标准贯入试验成果统计表*.XLS
      ├── 动力触探N63.5试验成果统计表*.XLS
      ├── 液化判别及液化指数计算成果表_*.XLS
      ├── 岩石试验指标分层统计表*.XLS
      ├── 水样.xlsx (可选)
      ├── 易溶盐土样.xlsx (可选)
      └── 建筑物特征一览表*.xlsx (可选)
              │
              ▼
    python generate_report.py --config project_config.json
              │
              ▼
    岩土工程勘察报告_正式报告.docx
```

## 使用方法

### 配置文件方式 (推荐)
```bash
python generate_report.py --config project_config.json
```

### 目录自动检测方式 (向后兼容 v3)
```bash
python generate_report.py --project "D:\项目\xxx项目"
```

### 指定模板
```bash
python generate_report.py --config project_config.json --template "D:\模板\报告模板.docx"
```

### 指定地层名称映射
```bash
python generate_report.py --project "D:\项目\xxx项目" --layers layers.json
```

### 仅加载数据查看（不生成报告）
```bash
python generate_report.py --config project_config.json --dry-run
```

### 详细日志
```bash
python generate_report.py --config project_config.json --verbose
```

## 配置文件格式 (project_config.json)

```json
{
  "project_name": "项目名称",
  "base_dir": "D:\\项目\\xxx项目",
  "excel_dir": "已有资料\\excel表格",
  "template": "报告模板.docx",
  "output_suffix": "_正式报告",

  "table_indices": {
    "buildings": 2,
    "workload": 3,
    "phys_spt_start": 4,
    "phys_spt_end": 15,
    "water_level": 16,
    "bearing_capacity": 18,
    "water_sample": 19,
    "water_corrosion": 20,
    "salt_sample": 21,
    "salt_corrosion": 22,
    "foundation_type": 23,
    "pile_params": 24
  },

  "layers": [
    {
      "id": "1",
      "name": "杂填土",
      "table_type": "cpt",
      "description": "黄褐色...场区普遍分布，厚度:{thick_min}~{thick_max}m,平均{thick_avg}m;...",
      "bearing": { "fak": "/", "es": "/", "e0": "/" },
      "pile": { "qsik1": "22", "qpk1": "/", "qsik2": "22", "qpk2": "/" }
    }
  ],

  "replacements": [
    ["原项目名", "新项目名"]
  ],

  "date_replacements": [
    ["2025年11月", "2025年12月"]
  ],

  "workload_paragraph": {
    "template": "本次勘察实际完成钻孔{total}个..."
  },

  "water_paragraph_template": "勘察期间测得钻孔内水位埋深约{depth_min}~{depth_max}m..."
}
```

### 地层描述模板变量

描述模板中可使用以下占位符，运行时自动替换为统计数据:

- `{thick_min}`, `{thick_max}`, `{thick_avg}` — 层厚
- `{depth_min}`, `{depth_max}`, `{depth_avg}` — 层底埋深
- `{elv_min}`, `{elv_max}`, `{elv_avg}` — 层底标高
- `{dist}` — 分布范围 (普遍/较普遍/局部，通用模板专用)

## 地层名称映射文件格式 (layers.json)

```json
{
  "1": "素填土",
  "2": "细砂",
  "3": "粉质粘土夹粉砂",
  "4": "淤泥质土",
  "5": "中细砂",
  "6": "淤泥质土",
  "7": "中粗砂",
  "8": "淤泥质土",
  "9": "粗砂",
  "10-1": "全风化片麻岩",
  "10-2": "强风化片麻岩",
  "10-3": "强风化片麻岩"
}
```

## 关键步骤

1. **配置加载**: 从 JSON 配置文件或项目目录自动检测获取项目参数
2. **数据加载**: xlrd/openpyxl 读取华宁导出 Excel, 自动识别双面板格式 (A3水样/盐样)
3. **腐蚀性评价**: 按 GB50021-2001 第12章, 自动评价水/土对混凝土和钢筋的腐蚀等级
4. **模板填充**: python-docx 读写 docx, 只替换文字不动表格格式
5. **输出**: 保存为 `{模板名}_正式报告.docx`

## 填充内容清单

- 建筑物特征表 (T2, 索引可配置)
- 工作量统计表+段落 (T3)
- 水位统计表+段落 (T16)
- 地层描述段落 (厚度/深度/标高数据注入, 描述模板从配置读取)
- 物理力学性质指标统计表 (T4-T15, 按层名匹配)
- 承载力建议值表 (T18, 参数从配置读取)
- 标贯/动探实测修正值
- 水样分析表+腐蚀性评价表 (T19-T20)
- 易溶盐分析表+腐蚀性评价表 (T21-T22)
- 液化判别段落
- 基础类型建议表 (T23)
- 桩基参数建议表 (T24, 参数从配置读取)
- 结论与建议段落
- 全局文本替换 (项目名/面积/日期, 规则从配置读取)

## 文件说明

| 文件 | 说明 |
|------|------|
| `generate_report.py` | 主程序 (v4.0, 配置驱动架构) |
| `corrosion_eval.py` | GB50021-2001 腐蚀性评价独立模块 |
| `project_config.example.json` | 完整配置文件模板 |
| `requirements.txt` | Python 依赖清单 |
| `README.md` | 用户文档 |
| `SKILL.md` | 本文件 (AI Agent 技能定义) |

## 依赖

```bash
pip install -r requirements.txt
# 或手动安装:
pip install python-docx xlrd openpyxl
```

## v4.0 改进要点

- JSON 配置文件驱动: 地层描述、承载力、桩基参数、表格索引均可配置
- 修复日期硬编码 bug (`_update_dates` 不再固定替换特定月份)
- 修复变量名遮蔽 (`ss` 局部变量遮蔽模块级函数)
- logging 框架替代 print, 支持 `--verbose` 调试
- 依赖延迟导入, `--help` 在未安装依赖时仍可运行
- 全面类型注解, 命名常量替代 magic number

## 局限与注意事项

- .doc 模板需先用 Word/WPS 转为 .docx
- A3双面板 Excel (水样/盐样) 的列偏移已硬编码, 不同版本可能需要调整
- 地层描述模板针对滨海/冲洪积地层, 其他地区需定制描述模板
- 新地区需通过配置文件调整承载力和桩基参数
