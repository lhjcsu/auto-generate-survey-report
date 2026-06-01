---
name: auto-generate-survey-report
description: "从华宁勘察软件导出的Excel数据自动生成岩土工程勘察报告。读取.XLS/.xlsx数据，填充正式报告.docx模板，输出完整报告。"
version: 1.0
author: Hermes Agent
category: geotechnical
---

# 自动生成工程勘察报告

从华宁勘察软件导出的 Excel 数据自动填充到正式报告模板，一键生成完整的岩土工程勘察报告（.docx）。

## 依赖安装

```bash
pip install xlrd openpyxl python-docx pywin32 -i https://pypi.tuna.tsinghua.edu.cn/simple
```

## 工作流程

华宁导出 Excel (.XLS/.xlsx) → Python 读取 → 加载正式报告 .docx 模板 → python-docx 填充表格+更新段落 → 输出完整报告 .docx

## 关键步骤

1. 先用 win32com 将 `.doc` 转为 `.docx`
2. 用 python-docx 读取 docx，遍历段落和表格
3. 遍历 doc.tables，根据表头内容识别对应数据表，用真实数据替换
4. 遍历 doc.paragraphs，更新数据关键段落（孔数、深度、工作量等）

## 典型华宁导出Excel文件

勘探点一览表19.XLS、标准贯入试验成果统计表19.XLS、动力触探N63.5试验成果统计表19.XLS.XLS、分层土工试验成果表19.XLS、物理力学性质指标统计表19.XLS、勘察工作量明细表19.xls、场地各孔厚度层底深度标高及层顶深度标高统计明细表1.xlsx、试验室土工试验成果报告表19.XLS、岩石试验指标分层统计表19.XLS、重型圆锥动力触探N63.5分层统计及承载力计算表(按点)19.XLS、标准贯入试验液化判别及液化指数计算成果表_19.XLS

## 项目案例

威海高区保税物流中心（B型）项目: D:\项目\威海高区保税物流中心（B型）项目\，脚本 generate_report.py（35KB），模板 25019威海高区保税物流中心.docx，118个钻孔/12层/25表/352段落
