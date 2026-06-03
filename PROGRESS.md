# 勘察报告自动生成工具 - 按章节优化进度

## 按章节完成状态 (基于城投文化路项目模板)

| 章节 | 功能 | 状态 | commit | 配置节 |
|------|------|------|--------|--------|
| 一 工程概况 | 委托/场地/建筑/勘察等级/冻结深度 | ✅ | 5c898e0 | `project_overview` |
| 二 勘察目的 | 任务要求条件过滤 (基坑条目) | ✅ | b7ef083 | `project_overview.has_basement` |
| 三 工作概述 | 工作量表+钻孔类型 (已有) | ✅ | - | `borehole_type_names` |
| 四 气象/地质 | 标准冻结深度注入 | ✅ | 5c898e0 | `project_overview.frozen_depth` |
| 五 场地条件 | 地形/地下水/地震/不良地质 (11段) | ✅ | 5c898e0 | `site_conditions` |
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

## 配置结构总览

```json
{
  "project_overview": {
    "client, survey_stage, commission_text, site_location, building_desc": "...",
    "importance_level, site_complexity, foundation_grade, survey_grade": "...",
    "frozen_depth": "0.50m",
    "has_basement": true
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
    "anti_float": ["段落1", "段落2", "..."],
    "foundation_text": ["..."],
    "pile_eval": ["..."],
    "special_soils": ["..."],
    "stability": ["..."],
    "excavation": ["..."],
    "risk": ["..."],
    "deformation": ["..."]
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

## fill_all() 执行顺序 (18步)

1. `_global_replace()` — 全局文本替换
2. `_fill_project_overview()` — 第一章: 工程概况
3. `_fill_survey_purpose()` — 第二章: 任务要求条件过滤
4. `_fill_buildings_table()` — 建筑物特征表
5. `_fill_workload()` — 工作量表
6. `_fill_water_level()` — 水位表
7. `_fill_layer_descriptions()` — 地层描述
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

## 提交历史 (8次提交)

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

## 待优化项目

- 第六章表7-1 (结论参数表) 自动填充 — 目前需工程师手动填写
- 地基基础评价总纲 (4.5.6) — 天然地基/桩基/地基处理方案建议细化
- 坐标系统和高程系统说明 (4.2.6§8) — 配置项

## 文件清单

- `generate_report.py` — 主程序 (v4.0+)
- `corrosion_eval.py` — 腐蚀性评价模块
- `project_config.example.json` — 配置文件模板
- `requirements.txt` — 依赖
- `README.md` — 文档

## 新对话继续方式

在新对话中发送: "读取 E:\Claudecode代码\auto-generate-survey-report\PROGRESS.md 继续开发"
