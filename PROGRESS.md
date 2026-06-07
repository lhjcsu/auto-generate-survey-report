# 勘察报告自动生成工具 - 按章节优化进度

## 按章节完成状态 (基于城投文化路项目模板)

| 章节 | 功能 | 状态 | commit | 配置节 |
|------|------|------|--------|--------|
| 一 工程概况 | 委托/场地/建筑/勘察等级/冻结深度 | ✅ | 5c898e0 | `project_overview` |
| 二 勘察目的 | 任务要求条件过滤 (基坑条目) | ✅ | b7ef083 | `project_overview.has_basement` |
| 三(一) 勘察方法 | N63.5动力触探/波速测试 条件过滤 | ✅ | 642539c | 自动检测 (n63_total/建筑高度>24m) |
| 三(二) 工作布置 | 布孔/孔深/水样 段落 (已有) | ✅ | - | `borehole_info` 自动注入 |
| 三(四) 完成情况 | 工作量统计表 + 完成段落 (已有) | ✅ | - | `borehole_info` 自动注入 |
| 三(五) 质量评述 | 勘察等级自动替换 (甲/乙/丙级) | ✅ | 642539c | `project_overview.survey_grade` |
| 三 工作进度 | 外业日期+钻机数量自动注入 | ✅ NEW | - | `project_overview.fieldwork_start/end/rig_count` |
| 三 其余段落 | 工艺描述/坐标系 → 模板保留, 手动改 | — | - | - |
| 四 气象/地质 | 标准冻结深度注入 | ✅ | 5c898e0 | `project_overview.frozen_depth` |
| 五 场地条件 | 地形/地下水/地震/不良地质 (11段) | ✅ | 5c898e0 | `site_conditions` |
| 五(二) 地层描述 | 华宁数据库动态生成 (ZHMS/DCSH) | ✅ NEW | - | `hn_db_dir` + `geo_age_names` |
| 六(一) | 岩土层逐层工程评价 | ✅ | 69a02ab | `analysis_evaluation.layer_eval` |
| 六(二) | 参数建议值表 (已有 T18) | ✅ | - | `bearing_values` |
| 六(三) | 场地稳定性及适宜性 (已有) | ✅ | ae3c1fb | `site_evaluation` |
| 六(四)1-2 | 水样/土样腐蚀性 (已有) | ✅ | - | `water_salt_tables` |
| 六(四)3 | 地下水力学作用/抗浮水位 | ✅ | 69a02ab | `analysis_evaluation.anti_float` |
| 六(五)1 | 基础选型分析 (已有 T23) | ✅ | - | `foundation_type` |
| 六(五)1+ | 基础选型补充说明 | ✅ | 69a02ab | `analysis_evaluation.foundation_text` |
| 六(五)2 | 桩基评价+参数表 (已有 T24) | ✅ | - | `pile_values` |
| 六(五)2+ | 桩基评价综合文字 | ✅ | 69a02ab | `analysis_evaluation.pile_eval` |
| 六(六) | 特殊性岩土评价 | ✅ | 69a02ab | `analysis_evaluation.special_soils` |
| 六(七) | 地基均匀性评价 (已有) | ✅ | 9c93fd9 | `foundation_evaluation` |
| 六(八) | 地基稳定性评价 | ✅ | 69a02ab | `analysis_evaluation.stability` |
| 六(九) | 基坑开挖 (has_basement控制) | ✅ | 69a02ab | `analysis_evaluation.excavation` |
| 六(十) | 工程风险 (has_basement控制) | ✅ | 69a02ab | `analysis_evaluation.risk` |
| 六(十一) | 建筑物变形分析 | ✅ | 69a02ab | `analysis_evaluation.deformation` |
| 七(一) | 结论 (自动占位符+配置段落) | ✅ | 79a61d4 | `conclusion_suggestions.conclusion` |
| 七(二) | 建议 (12条配置段落) | ✅ | 79a61d4 | `conclusion_suggestions.suggestions` |
| 技术标准 | 配置驱动+条件过滤 | ✅ | d1ab318 | `technical_standards` |

## 条件过滤逻辑汇总

| 条件 | 检测方式 | 影响范围 |
|------|---------|---------|
| `has_basement` | 建筑物数据含"地下" 或 `project_overview.has_basement` | 第二章: 删除基坑任务条目并重编号 |
| `has_basement` | 同上 | 第六章: 清空基坑开挖(九)+工程风险(十)段落 |
| N63.5 | `borehole_info.n63_total > 0` | 第三章: 删除动力触探方法段落 |
| 高层建筑 | 任一建筑物高度 > 24m (正则提取数字) | 第三章: 删除波速测试方法段落 |
| 液化 | 标贯数据自动判别 | 六(三)/七(一): 注入液化结论文字 |
| 岩层/岩样/水样/桩基 | 数据自动检测 | 技术标准: 条件包含对应规范 |

## 配置结构总览

```json
{
  "project_overview": {
    "client, survey_stage, commission_text, site_location, building_desc": "...",
    "importance_level, site_complexity, foundation_grade, survey_grade": "...",
    "frozen_depth": "0.50m",
    "has_basement": true
  },
  "hn_db_dir": "D:\\...\\sj",
  "hn_project_code": "17",
  "geo_age_names": {
    "N": "新元古界荣成序列威海单元（NhηγRw）",
    "Q4|ml": "第四系人工堆积层（Q4ml）"
  },
  "site_conditions": {
    "terrain_text, topography_text, environment_text, surface_water_text": "...",
    "seismic_params_text, site_class_text, seismic_stability_text": "...",
    "seismic_zone_text, soft_soil_text, adverse_text, buried_text": "..."
  },
  "site_evaluation": {
    "paragraphs": ["..."],
    "stability_grade, suitability_grade, liquefaction_text, ...": "..."
  },
  "foundation_evaluation": {
    "paragraphs": ["..."],
    "uniformity_text, weak_layer_text, deformation_text, ...": "..."
  },
  "analysis_evaluation": {
    "layer_eval": { "1": "...", "2": "...", "5-1": "..." },
    "anti_float, foundation_text, pile_eval, special_soils": ["段落..."],
    "stability, excavation, risk, deformation": ["段落..."]
  },
  "conclusion_suggestions": {
    "conclusion": ["1、...", "2、...", "..."],
    "suggestions": ["1、...", "2、...", "..."]
  },
  "technical_standards": {
    "always": [{"code": "...", "name": "..."}],
    "conditional": [{"code": "...", "name": "...", "condition": "..."}]
  }
}
```

## fill_all() 执行顺序 (20步)

1. `_global_replace()` — 全局文本替换
2. `_fill_project_overview()` — 第一章: 工程概况
3. `_fill_survey_purpose()` — 第二章: 任务要求条件过滤
4. `_fill_buildings_table()` — 建筑物特征表
5. `_fill_workload()` — 第三章: 工作量表 + 条件过滤(N63.5/波速) + 勘察等级
6. `_fill_water_level()` — 水位表
7. `_fill_layer_descriptions()` — 地层描述 (华宁数据库模式/回退模板模式)
8. `_fill_phys_spt_tables()` — 物理力学表
9. `_fill_bearing_capacity()` — 承载力表
10. `_fill_water_salt_tables()` — 水样/盐样表
11. `_fill_corrosion_eval()` — 腐蚀性评价
12. `_fill_liquefaction()` — 液化判别
13. `_fill_foundation_tables()` — 基础建议表+桩基参数
14. `_fill_conclusion()` — 第七章: 结论与建议
15. `_fill_site_conditions()` — 第五章: 场地条件
16. `_fill_site_evaluation()` — 场地稳定性评价
17. `_fill_foundation_evaluation()` — 地基评价
18. `_fill_analysis_evaluation()` — 第六章: 分析评价各子节
19. `_fill_standards()` — 技术标准列表
20. `_apply_date_replacements()` — 日期替换

## 提交历史 (12次提交)

| commit | 描述 |
|--------|------|
| 395507c | v4.0 配置驱动重构 |
| d1ab318 | 技术标准自动填充 |
| ae3c1fb | 场地稳定性及适宜性评价 |
| 9c93fd9 | 地基评价模块 |
| 5c898e0 | 第一/四/五章配置驱动填充 |
| b7ef083 | 第二章任务要求条件过滤 |
| 69a02ab | 第六章分析评价9个子节 |
| 79a61d4 | 第七章结论与建议 |
| a2814fd | 更新进度文档 |
| 642539c | 第三章: N63.5/波速条件过滤 + 勘察等级注入 |
| NEW | 华宁数据库读取器 + 第五章(二)地层描述动态生成 |

## 华宁数据库读取器 (HuaNingDBReader) — NEW

### 概述
直接读取华宁HNCAD勘察软件原始数据库文件，动态生成第五章(二)岩土结构及工程特性段落。
解决了之前依赖模板段落匹配导致不同项目地层不对应的问题。

### 读取的文件

| 文件 | 内容 | 编码 |
|------|------|------|
| `ZHMS.{code}` | 地层描述 (层号,岩土名:描述) | GB2312 |
| `DCSH.{code}` | 地层序列+地质年代 (层号,岩性代码,,,成因,年代) | GB2312 |
| `DCSJ.{code}` | 逐孔地层 (孔号,层号,层底深度) 最后层无深度=未穿透 | GB2312 |
| `BG.{code}` | 标贯数据 (孔号,起始深度,终止深度,N值) | GB2312 |
| `TY.{code}` | 取样数据 (孔号,深度,类型: 0=扰动/1=原状/2=岩样) | GB2312 |

### 动态生成流程

1. 从 ZHMS 读取每层岩土名称和描述
2. 从 DCSH 读取地质年代分组 (Q4ml → Q4al+pl → Q4dl+el → 基岩N)
3. 从 DCSJ 检测未穿透层 (最后一层无深度) 和最大揭露厚度
4. 从 BG 统计每层标贯次数 (用DCSJ层底深度定位所属层)
5. 从 TY 统计每层取样数量 (扰动样/原状样)
6. 用 XML addnext() 在(二)节标题后动态插入:
   - 加粗的地质年代标题 (如 "第四系人工堆积层（Q4ml）")
   - 每层: 地层描述段落 + 试验信息段落 + 表标题段落
   - 收尾段落

### 配置项

- `hn_db_dir`: 华宁数据库目录路径
- `hn_project_code`: 项目代码 (或自动检测)
- `geo_age_names`: 地质年代名称覆盖 (用于基岩等复杂名称)

### 回退机制
未配置 `hn_db_dir` 时自动回退到模板匹配模式 (`_fill_layer_descriptions_fallback`)

## 待优化项目

- 五(二) 试验信息段落: 标贯/取样数量与华宁简报核对 (层边界定位精度)
- 五(二) 岩石质量段落: 需从配置或岩体力学数据生成 (坚硬程度/完整程度/质量等级)
- 五(二) 物理力学表: 动态表格插入 (当前依赖模板固定表格)
- 第六章表7-1 (结论参数表) 自动填充 — 目前需工程师手动填写
- 地基基础评价总纲 (4.5.6) — 天然地基/桩基/地基处理方案建议细化

## 文件清单

- `generate_report.py` — 主程序 (v4.0+)
- `corrosion_eval.py` — 腐蚀性评价模块
- `project_config.example.json` — 配置文件模板
- `requirements.txt` — 依赖
- `README.md` — 文档

## 新对话继续方式

在新对话中发送: "读取 E:\Claudecode代码\auto-generate-survey-report\PROGRESS.md 继续开发"
