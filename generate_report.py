#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
岩土工程勘察报告 — 自动生成工具 v4.0
============================================================
从华宁勘察软件导出的 Excel 数据 + 正式报告 .docx 模板，
一键生成完整的岩土工程勘察报告。

用法:
    # 方式一: 使用 JSON 配置文件 (推荐)
    python generate_report.py --config project_config.json

    # 方式二: 指定项目目录 (向后兼容 v3)
    python generate_report.py --project <项目目录>

    # 可选参数
    --template  指定 .docx 模板路径
    --output    指定输出文件路径
    --layers    指定地层名称映射 JSON
    --dry-run   仅加载数据，不生成报告
    --verbose   输出详细日志

依赖:
    pip install python-docx xlrd openpyxl
"""
from __future__ import annotations

import argparse
import copy
import glob
import json
import logging
import os
import re
import sys
from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple

# 第三方依赖: 延迟导入以支持 --help 在未安装依赖时仍可运行
try:
    import openpyxl
    import xlrd
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    # 仅在真正需要使用时才报错
    openpyxl = None  # type: ignore[assignment]
    xlrd = None  # type: ignore[assignment]
    Document = None  # type: ignore[assignment,misc]

from corrosion_eval import evaluate_corrosion


def _check_dependencies() -> None:
    """检查第三方依赖是否已安装"""
    missing = []
    if openpyxl is None:
        missing.append("openpyxl")
    if xlrd is None:
        missing.append("xlrd")
    if Document is None:
        missing.append("python-docx")
    if missing:
        print(f"错误: 缺少依赖包: {', '.join(missing)}")
        print(f"请运行: pip install {' '.join(missing)}")
        sys.exit(1)

__version__ = "4.0"

# ============================================================
# 日志配置
# ============================================================

logger = logging.getLogger("survey_report")


def setup_logging(verbose: bool = False) -> None:
    """配置日志输出格式和级别"""
    level = logging.DEBUG if verbose else logging.INFO
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter("%(message)s"))
    logger.setLevel(level)
    logger.addHandler(handler)


# ============================================================
# 默认表格索引 (可通过配置文件覆盖)
# ============================================================

DEFAULT_TABLE_INDICES: Dict[str, int] = {
    "buildings": 2,          # T2  建筑物特征表
    "workload": 3,           # T3  工作量统计表
    "phys_spt_start": 4,     # T4~T15 物理力学 & 原位测试
    "phys_spt_end": 15,
    "water_level": 16,       # T16 水位统计表
    "bearing_capacity": 18,  # T18 承载力建议值
    "water_sample": 19,      # T19 水样分析表
    "water_corrosion": 20,   # T20 水腐蚀性评价表
    "salt_sample": 21,       # T21 易溶盐分析表
    "salt_corrosion": 22,    # T22 土腐蚀性评价表
    "foundation_type": 23,   # T23 基础类型建议表
    "pile_params": 24,       # T24 桩基参数建议表
}

# 物理力学指标: 行号 → 指标键名
PHYS_ROW_MAP: Dict[int, str] = {
    1: "W", 2: "gamma", 3: "e0",
    4: "WL", 5: "WP", 6: "IP", 7: "IL",
}

# 物理力学统计列: 统计类型 → 列号
STAT_COL_MAP: Dict[str, int] = {
    "min": 1, "max": 2, "avg": 3,
    "n": 4, "std": 5, "cv": 6, "std_val": 7,
}

# 统计关键词 → 内部键名
STAT_KEYWORD_MAP: Dict[str, str] = {
    "最小值": "min", "最大值": "max", "数据个数": "n",
    "平均值": "avg", "标准差": "std", "变异系数": "cv", "标准值": "std_val",
}


# ============================================================
# 工具函数
# ============================================================

def safe_float(v: Any, default: Optional[float] = None) -> Optional[float]:
    """安全转换为 float，失败返回 default"""
    if v is None:
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s in ("", "-", "—", "/"):
        return default
    try:
        return float(s)
    except (ValueError, TypeError):
        return default


def safe_str(v: Any) -> str:
    """安全转换为去首尾空格的字符串"""
    if v is None:
        return ""
    return str(v).strip()


def fmt_val(v: Any, fmt: str = ".2f") -> str:
    """格式化浮点数，None 返回空串"""
    if v is None:
        return ""
    try:
        return f"{float(v):{fmt}}"
    except (ValueError, TypeError):
        return ""


def fmt_val_int(v: Any) -> str:
    """格式化整数或保留一位小数"""
    if v is None:
        return ""
    try:
        f = float(v)
        return str(int(f)) if f == int(f) else f"{f:.1f}"
    except (ValueError, TypeError):
        return ""


def find_file(directory: str, pattern: str) -> Optional[str]:
    """在目录及其子目录中查找匹配文件 (忽略大小写)"""
    if not os.path.isdir(directory):
        return None
    pattern_lower = pattern.lower()
    # 先搜当前目录
    for f in os.listdir(directory):
        if pattern_lower in f.lower():
            return os.path.join(directory, f)
    # 递归搜索
    for root, _dirs, files in os.walk(directory):
        for f in files:
            if pattern_lower in f.lower():
                return os.path.join(root, f)
    return None


def set_cell(table: Any, row: int, col: int, val: Any) -> None:
    """设置 docx 表格单元格文字，保留原有格式"""
    if row >= len(table.rows):
        return
    if col >= len(table.rows[row].cells):
        return
    cell = table.rows[row].cells[col]
    text = str(val) if val is not None else ""
    para = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    if para.runs:
        para.runs[0].text = text
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.add_run(text)


def set_para_text(para: Any, text: str) -> None:
    """设置段落文字，保留首个 run 的格式"""
    if para.runs:
        para.runs[0].text = text
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.add_run(text)


def replace_in_para(para: Any, old: str, new: str) -> None:
    """在段落的所有 run 中执行文本替换"""
    for run in para.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)


def layer_sort_key(layer_id: str) -> Tuple[int, int]:
    """地层排序键: 按数字排序，支持 '10-1' 格式"""
    parts = str(layer_id).replace("层", "").replace("第", "").split("-")
    major = int(parts[0]) if parts[0].lstrip("-").isdigit() else 999
    minor = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
    return (major, minor)


# ============================================================
# 项目配置管理
# ============================================================

class ProjectConfig:
    """项目配置 — 支持 JSON 配置文件和目录自动检测两种方式"""

    def __init__(
        self,
        config_path: Optional[str] = None,
        project_dir: Optional[str] = None,
        template_override: Optional[str] = None,
        output_override: Optional[str] = None,
        layers_override: Optional[str] = None,
    ):
        self.raw: Dict[str, Any] = {}
        self.project_name: str = ""
        self.base_dir: str = ""
        self.excel_dir: str = ""
        self.template_path: str = ""
        self.output_path: str = ""
        self.layers_file: Optional[str] = layers_override

        # 表格索引 (可被配置文件覆盖)
        self.table_indices: Dict[str, int] = dict(DEFAULT_TABLE_INDICES)

        # 地层配置
        self.layers: List[Dict[str, Any]] = []

        # 全局替换规则
        self.replacements: List[Tuple[str, str]] = []

        # 含水层/液化配置
        self.water_aquifer_layers: List[str] = []
        self.liquefaction_layers: Dict[str, Dict] = {}

        # 段落模板
        self.workload_template: str = ""
        self.water_paragraph_template: str = ""

        if config_path:
            self._load_from_config(config_path)
        elif project_dir:
            self._load_from_directory(project_dir)
        else:
            raise ValueError("必须提供 config_path 或 project_dir")

        # CLI 覆盖
        if template_override:
            self.template_path = template_override
        if output_override:
            self.output_path = output_override

    def _load_from_config(self, config_path: str) -> None:
        """从 JSON 配置文件加载"""
        if not os.path.isfile(config_path):
            raise FileNotFoundError(f"配置文件不存在: {config_path}")

        with open(config_path, "r", encoding="utf-8") as f:
            self.raw = json.load(f)

        config_dir = os.path.dirname(os.path.abspath(config_path))
        self.project_name = self.raw.get("project_name", "未知项目")

        # base_dir: 优先使用配置中的绝对路径，否则相对于配置文件
        base = self.raw.get("base_dir", "")
        if base and os.path.isabs(base):
            self.base_dir = base
        elif base:
            self.base_dir = os.path.join(config_dir, base)
        else:
            self.base_dir = config_dir

        # excel_dir
        excel_sub = self.raw.get("excel_dir", "")
        if excel_sub:
            self.excel_dir = os.path.join(self.base_dir, excel_sub)
        else:
            self.excel_dir = self._find_excel_dir()

        # 模板
        tpl = self.raw.get("template", "")
        if tpl and os.path.isabs(tpl):
            self.template_path = tpl
        elif tpl:
            self.template_path = os.path.join(self.base_dir, tpl)
        else:
            self.template_path = self._find_template()

        # 输出
        suffix = self.raw.get("output_suffix", "_正式报告")
        tpl_name = os.path.splitext(os.path.basename(self.template_path))[0]
        self.output_path = os.path.join(self.base_dir, f"{tpl_name}{suffix}.docx")

        # 表格索引覆盖
        if "table_indices" in self.raw:
            self.table_indices.update(self.raw["table_indices"])

        # 地层配置
        self.layers = self.raw.get("layers", [])

        # 全局替换
        self.replacements = [tuple(r) for r in self.raw.get("replacements", [])]

        # 含水层/液化
        self.water_aquifer_layers = self.raw.get("water_aquifer_layers", [])
        self.liquefaction_layers = self.raw.get("liquefaction_layers", {})

        # 段落模板
        wp = self.raw.get("workload_paragraph", {})
        self.workload_template = wp.get("template", "")
        self.water_paragraph_template = self.raw.get("water_paragraph_template", "")

        logger.info(f"  项目: {self.project_name}")
        logger.info(f"  目录: {self.base_dir}")
        logger.info(f"  Excel: {self.excel_dir}")
        logger.info(f"  模板: {self.template_path}")
        logger.info(f"  输出: {self.output_path}")

    def _load_from_directory(self, project_dir: str) -> None:
        """从项目目录自动检测 (向后兼容 v3)"""
        if not os.path.isdir(project_dir):
            raise FileNotFoundError(f"项目目录不存在: {project_dir}")

        self.base_dir = project_dir
        self.project_name = os.path.basename(project_dir)
        self.excel_dir = self._find_excel_dir()
        self.template_path = self._find_template()

        tpl_name = os.path.splitext(os.path.basename(self.template_path))[0]
        self.output_path = os.path.join(self.base_dir, f"{tpl_name}_正式报告.docx")

    def _find_excel_dir(self) -> str:
        """自动查找 Excel 数据目录"""
        candidates = [
            os.path.join(self.base_dir, "已有资料", "excel表格"),
            os.path.join(self.base_dir, "excel表格"),
            os.path.join(self.base_dir, "已有资料"),
            self.base_dir,
        ]
        for d in candidates:
            if not os.path.isdir(d):
                continue
            xls_files = (
                glob.glob(os.path.join(d, "*.XLS"))
                + glob.glob(os.path.join(d, "*.xls"))
                + glob.glob(os.path.join(d, "*.xlsx"))
            )
            if len(xls_files) >= 3:
                return d
        # 递归查找
        for root, _dirs, files in os.walk(self.base_dir):
            xls_count = sum(1 for f in files if f.lower().endswith((".xls", ".xlsx")))
            if xls_count >= 3:
                return root
        return self.base_dir

    def _find_template(self) -> str:
        """查找 .docx 模板文件"""
        for f in os.listdir(self.base_dir):
            fp = os.path.join(self.base_dir, f)
            if (
                f.endswith(".docx")
                and "正式报告" not in f
                and "自动生成" not in f
                and "temp" not in f.lower()
            ):
                return fp
        # 递归查找
        for root, _dirs, files in os.walk(self.base_dir):
            for f in files:
                if (
                    f.endswith(".docx")
                    and "正式报告" not in f
                    and "自动生成" not in f
                ):
                    return os.path.join(root, f)
        raise FileNotFoundError(f"未找到 .docx 模板文件: {self.base_dir}")

    def get_layer_config(self, layer_id: str) -> Optional[Dict[str, Any]]:
        """获取指定地层的配置"""
        for layer in self.layers:
            if str(layer.get("id", "")) == str(layer_id):
                return layer
        return None

    def get_layer_description(self, layer_id: str) -> Optional[str]:
        """获取地层描述模板"""
        cfg = self.get_layer_config(layer_id)
        return cfg.get("description") if cfg else None

    def get_bearing_values(self, layer_id: str) -> Optional[Dict[str, str]]:
        """获取承载力参数"""
        cfg = self.get_layer_config(layer_id)
        return cfg.get("bearing") if cfg else None

    def get_pile_values(self, layer_id: str) -> Optional[Dict[str, str]]:
        """获取桩基参数"""
        cfg = self.get_layer_config(layer_id)
        return cfg.get("pile") if cfg else None

    def get_workload_config(self) -> Dict[str, Any]:
        """获取工作量段落配置"""
        return self.raw.get("workload_paragraph", {})

    def get_date_replacements(self) -> List[Tuple[str, str]]:
        """获取日期替换规则"""
        return [tuple(r) for r in self.raw.get("date_replacements", [])]

    def get_technical_standards(self) -> Dict[str, Any]:
        """获取技术标准配置"""
        return self.raw.get("technical_standards", {})

    def get_site_evaluation(self) -> Dict[str, Any]:
        """获取场地稳定性及适宜性评价配置"""
        return self.raw.get("site_evaluation", {})

    def get_foundation_evaluation(self) -> Dict[str, Any]:
        """获取地基评价配置 (4.5.7 §2 均匀性 / §5 软弱下卧层 / §6 变形参数)"""
        return self.raw.get("foundation_evaluation", {})

    def get_project_overview(self) -> Dict[str, Any]:
        """获取拟建工程概况配置 (第一章)"""
        return self.raw.get("project_overview", {})

    def get_site_conditions(self) -> Dict[str, Any]:
        """获取场地条件配置 (第五章: 地形地貌/地下水/地震/不良地质)"""
        return self.raw.get("site_conditions", {})

    def get_analysis_evaluation(self) -> Dict[str, Any]:
        """获取岩土工程分析评价配置 (第六章各子节段落)"""
        return self.raw.get("analysis_evaluation", {})

    def get_conclusion_suggestions(self) -> Dict[str, Any]:
        """获取结论与建议配置 (第七章: conclusion/suggestions 段落数组)"""
        return self.raw.get("conclusion_suggestions", {})

    def get_hn_db_dir(self) -> str:
        """获取华宁数据库目录 (绝对路径或相对于 base_dir)"""
        d = self.raw.get("hn_db_dir", "")
        if d and os.path.isabs(d):
            return d
        elif d:
            return os.path.join(self.base_dir, d)
        return ""

    def get_geo_age_names(self) -> Dict[str, str]:
        """获取地质年代名称覆盖配置 {group_key: display_name}"""
        return self.raw.get("geo_age_names", {})

    def get_hn_project_code(self) -> str:
        """获取华宁项目代码"""
        return self.raw.get("hn_project_code", "")


# ============================================================
# 数据加载器
# ============================================================

class SurveyDataLoader:
    """从华宁导出 Excel 中加载所有勘察数据"""

    def __init__(self, excel_dir: str):
        self.excel_dir = excel_dir

    # ---- 通用读取 ----

    @staticmethod
    def _read_sheet(path: str) -> List[List[Any]]:
        """读取 xls 或 xlsx 文件的首个工作表"""
        if path.lower().endswith(".xlsx"):
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            ws = wb.worksheets[0]
            rows = [list(row) for row in ws.iter_rows(values_only=True)]
            wb.close()
            return rows
        else:
            wb = xlrd.open_workbook(path)
            ws = wb.sheet_by_index(0)
            return [
                [ws.cell(r, c).value for c in range(ws.ncols)]
                for r in range(ws.nrows)
            ]

    # ---- 勘探点一览表 ----

    def load_boreholes(self) -> List[Dict[str, Any]]:
        """勘探点一览表 → 钻孔列表"""
        path = find_file(self.excel_dir, "勘探点一览表")
        if not path:
            raise FileNotFoundError("未找到 勘探点一览表*.XLS")

        rows = self._read_sheet(path)
        boreholes: List[Dict[str, Any]] = []

        for r in range(6, len(rows)):
            seq = safe_str(rows[r][0]) if len(rows[r]) > 0 else ""
            if not seq:
                continue
            if seq in ("合计", "总 计") or seq.startswith("="):
                break
            if len(rows[r]) < 14:
                continue

            boreholes.append({
                "seq": seq,
                "id": safe_str(rows[r][1]),
                "type": safe_str(rows[r][2]),
                "elevation": safe_float(rows[r][3]),
                "depth": safe_float(rows[r][4]),
                "wt_depth": safe_float(rows[r][7]),
                "wt_elv": safe_float(rows[r][8]),
                "undisturbed": int(safe_float(rows[r][9], 0)),
                "disturbed": int(safe_float(rows[r][10], 0)),
                "rock": int(safe_float(rows[r][11], 0)),
                "spt": int(safe_float(rows[r][12], 0)),
                "n63": int(safe_float(rows[r][13], 0)),
            })

        logger.info(f"  勘探点: {len(boreholes)} 个")
        return boreholes

    @staticmethod
    def classify_boreholes(
        boreholes: List[Dict[str, Any]],
    ) -> Dict[str, Any]:
        """智能分类钻孔类型并汇总统计"""
        types: Dict[str, int] = defaultdict(int)
        for bh in boreholes:
            types[bh["type"]] += 1

        result: Dict[str, Any] = {
            "total": len(boreholes),
            "total_depth": sum(bh["depth"] or 0 for bh in boreholes),
        }
        result["qutu"] = types.get("取土孔", 0)
        result["biaoguan"] = types.get("标贯孔", 0)
        result["yiban"] = types.get("一般性钻孔", 0)
        result["zhongtan"] = types.get("重探孔", types.get("重探", 0))
        result["bosk"] = types.get("波速孔", 0)
        result["ctrl"] = result["qutu"] + result["zhongtan"] + result["bosk"]
        result["general"] = result["total"] - result["ctrl"]
        result["undisturbed"] = sum(bh["undisturbed"] for bh in boreholes)
        result["disturbed"] = sum(bh["disturbed"] for bh in boreholes)
        result["rock"] = sum(bh["rock"] for bh in boreholes)
        result["spt_total"] = sum(bh["spt"] for bh in boreholes)
        result["n63_total"] = sum(bh["n63"] for bh in boreholes)
        result["all_types"] = dict(types)

        # 水位统计
        wtd = [bh["wt_depth"] for bh in boreholes if bh["wt_depth"]]
        wte = [bh["wt_elv"] for bh in boreholes if bh["wt_elv"]]
        if wtd:
            result["wt_depth_min"] = min(wtd)
            result["wt_depth_max"] = max(wtd)
            result["wt_elv_min"] = min(wte)
            result["wt_elv_max"] = max(wte)

        # 高程统计
        elvs = [bh["elevation"] for bh in boreholes if bh["elevation"]]
        if elvs:
            result["elv_min"] = min(elvs)
            result["elv_max"] = max(elvs)

        return result

    # ---- 地层厚度统计 ----

    def load_layer_stats(self) -> Dict[str, Dict[str, Any]]:
        """场地地层厚度统计表 → 地层统计数据"""
        path = find_file(self.excel_dir, "地层厚度统计") or find_file(
            self.excel_dir, "厚度层底深度"
        )
        if not path:
            logger.warning("  [!] 未找到地层厚度统计表，跳过")
            return {}

        rows = self._read_sheet(path)
        layers: Dict[str, Dict[str, Any]] = {}

        for r in range(4, len(rows)):
            lid = safe_str(rows[r][0])
            if not lid:
                continue
            if len(rows[r]) < 17:
                continue
            layers[lid] = {
                "thick_min": safe_float(rows[r][1]),
                "thick_max": safe_float(rows[r][2]),
                "thick_avg": safe_float(rows[r][3]),
                "depth_min": safe_float(rows[r][4]),
                "depth_max": safe_float(rows[r][5]),
                "depth_avg": safe_float(rows[r][6]),
                "elv_min": safe_float(rows[r][7]),
                "elv_max": safe_float(rows[r][8]),
                "elv_avg": safe_float(rows[r][9]),
                "n": safe_float(rows[r][10]),
                "top_d_min": safe_float(rows[r][11]),
                "top_d_max": safe_float(rows[r][12]),
                "top_d_avg": safe_float(rows[r][13]),
                "top_e_min": safe_float(rows[r][14]),
                "top_e_max": safe_float(rows[r][15]),
                "top_e_avg": safe_float(rows[r][16]),
            }

        logger.info(f"  地层统计: {len(layers)} 层")
        return layers

    # ---- 物理力学性质指标统计 ----

    def load_physical_stats(self) -> Dict[str, Dict[str, Any]]:
        """物理力学性质指标统计表"""
        path = find_file(self.excel_dir, "物理力学")
        if not path:
            logger.warning("  [!] 未找到物理力学统计表，跳过")
            return {}

        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(0)
        phys: Dict[str, Dict[str, Any]] = {}
        current: Optional[str] = None

        for r in range(8, ws.nrows):
            lid = safe_str(ws.cell(r, 0).value)
            lname = safe_str(ws.cell(r, 1).value)
            stat = safe_str(ws.cell(r, 2).value).replace(" ", "")

            if lid:
                current = lid
            if not current:
                continue
            if current not in phys:
                phys[current] = {"name": lname, "stats": {}}

            if stat in STAT_KEYWORD_MAP:
                sk = STAT_KEYWORD_MAP[stat]
                phys[current]["stats"][sk] = {
                    "W": safe_float(ws.cell(r, 6).value),
                    "gamma": safe_float(ws.cell(r, 8).value),
                    "e0": safe_float(ws.cell(r, 10).value),
                    "WL": safe_float(ws.cell(r, 12).value),
                    "WP": safe_float(ws.cell(r, 13).value),
                    "IP": safe_float(ws.cell(r, 14).value),
                    "IL": safe_float(ws.cell(r, 15).value),
                }

        return phys

    # ---- 标准贯入试验 ----

    def load_spt_stats(self) -> Dict[str, Dict[str, Any]]:
        """标准贯入试验成果统计表"""
        path = find_file(self.excel_dir, "标准贯入")
        if not path:
            logger.warning("  [!] 未找到标贯统计表，跳过")
            return {}

        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(0)
        spt: Dict[str, Dict[str, Any]] = {}
        current: Optional[str] = None

        for r in range(6, ws.nrows):
            lid = safe_str(ws.cell(r, 0).value)
            stat = safe_str(ws.cell(r, 1).value).replace(" ", "")

            if lid and (
                lid.replace("-", "").replace(".", "").isdigit() or "-" in lid
            ):
                current = lid
            if not current:
                continue
            if stat in STAT_KEYWORD_MAP:
                spt.setdefault(current, {})[STAT_KEYWORD_MAP[stat]] = {
                    "raw": safe_float(ws.cell(r, 7).value),
                    "corr": safe_float(ws.cell(r, 8).value),
                }

        return spt

    # ---- 动力触探 ----

    def load_cpt_stats(self) -> Dict[str, Dict[str, Any]]:
        """动力触探 N63.5 试验成果统计表"""
        path = find_file(self.excel_dir, "动力触探")
        if not path:
            logger.warning("  [!] 未找到动探统计表，跳过")
            return {}

        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(0)
        cpt: Dict[str, Dict[str, Any]] = {}
        current: Optional[str] = None

        for r in range(ws.nrows):
            lid = safe_str(ws.cell(r, 0).value)
            stat = safe_str(ws.cell(r, 1).value).replace(" ", "")

            if lid and (lid.isdigit() or "-" in lid):
                current = lid
            if not current:
                continue
            if stat in STAT_KEYWORD_MAP:
                cpt.setdefault(current, {})[STAT_KEYWORD_MAP[stat]] = {
                    "raw": safe_float(ws.cell(r, 5).value),
                    "corr": safe_float(ws.cell(r, 6).value),
                }

        return cpt

    # ---- 液化判别 ----

    def load_liquefaction(self) -> Tuple[List[List[Any]], int, int]:
        """液化判别及液化指数计算成果表"""
        path = find_file(self.excel_dir, "液化")
        if not path:
            return [], 0, 0

        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(0)
        data: List[List[Any]] = []

        for r in range(6, ws.nrows):
            if not safe_str(ws.cell(r, 0).value):
                continue
            data.append([ws.cell(r, c).value for c in range(min(20, ws.ncols))])

        liq_count = sum(
            1 for row in data if len(row) > 10 and "液" in safe_str(row[10])
        )
        return data, liq_count, len(data) - liq_count

    # ---- 岩石试验 ----

    def load_rock_stats(self) -> Dict[str, Dict[str, Any]]:
        """岩石试验指标分层统计表"""
        path = find_file(self.excel_dir, "岩石试验")
        if not path:
            return {}

        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(0)
        rock: Dict[str, Dict[str, Any]] = {}
        current: Optional[str] = None

        for r in range(7, ws.nrows):
            lid = safe_str(ws.cell(r, 0).value)
            stat = safe_str(ws.cell(r, 1).value).replace(" ", "")

            if lid:
                current = lid
            if not current:
                continue
            if stat in STAT_KEYWORD_MAP:
                rock.setdefault(current, {})[STAT_KEYWORD_MAP[stat]] = safe_float(
                    ws.cell(r, 6).value
                )

        return rock

    # ---- 建筑物特征 ----

    def load_buildings(self) -> List[Dict[str, str]]:
        """建筑物特征一览表"""
        path = find_file(self.excel_dir, "建筑物特征")
        if not path:
            return []

        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.worksheets[0]
            buildings: List[Dict[str, str]] = []

            for row in ws.iter_rows(values_only=True):
                row_list = list(row)
                name = safe_str(row_list[0])
                if not name or name in ("合计", "说明", "备注"):
                    continue
                if "名称" in name or "序号" in name:
                    continue
                buildings.append({
                    "name": name,
                    "floors": safe_str(row_list[2]) if len(row_list) > 2 else "",
                    "height": safe_str(row_list[3]) if len(row_list) > 3 else "",
                    "size": safe_str(row_list[4]) if len(row_list) > 4 else "",
                    "span": safe_str(row_list[5]) if len(row_list) > 5 else "",
                    "indoor_elv": safe_str(row_list[6]) if len(row_list) > 6 else "",
                    "structure": safe_str(row_list[7]) if len(row_list) > 7 else "",
                })

            wb.close()
            logger.info(f"  建筑物: {len(buildings)} 栋")
            return buildings
        except Exception as e:
            logger.error(f"  [!] 建筑物特征读取异常: {e}")
            return []

    # ---- 水样 (支持 A3 双面板) ----

    def load_water_samples(self) -> List[Dict[str, Any]]:
        """水样 xlsx 解析 (支持 A3 双面板格式)"""
        path = find_file(self.excel_dir, "水样")
        if not path:
            return []

        def parse_panel(
            label_col: int, name_col: int, val_col: int, extra_col: int,
            rows: List[List[Any]],
        ) -> Dict[str, Any]:
            result: Dict[str, Any] = {}
            # 从 label 列提取野编号
            for r in rows[:7]:
                txt = safe_str(r[label_col]) if len(r) > label_col else ""
                if "野编号" in txt:
                    m = re.search(r"(\d+)", txt)
                    if m:
                        result["field_id"] = m.group(1)
            # 解析离子含量
            for r in rows:
                if len(r) <= max(name_col, val_col, extra_col):
                    continue
                name = safe_str(r[name_col])
                val = safe_float(r[val_col]) if len(r) > val_col else None
                extra_lbl = (
                    safe_str(r[extra_col - 1]) if len(r) > extra_col - 1 else ""
                )
                extra_val = safe_float(r[extra_col]) if len(r) > extra_col else None

                if name == "PH值":
                    result["pH"] = val
                elif name == "Mg2+":
                    result["Mg"] = val
                elif name == "NH4+":
                    result["NH4"] = val or 0
                elif name == "Cl-":
                    result["Cl"] = val
                elif name == "SO42-":
                    result["SO4"] = val
                elif name == "HCO3-":
                    result["HCO3"] = val
                elif name == "CO32-":
                    result["CO3"] = val or 0
                elif name == "OH-":
                    result["OH"] = val or 0

                if safe_str(r[label_col]) == "PH值":
                    result["pH"] = safe_float(r[val_col])
                if "总矿化度" in extra_lbl:
                    result["TDS"] = extra_val
                if "侵蚀" in extra_lbl:
                    result["CO2_agg"] = extra_val or 0
            return result

        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.worksheets[0]
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            wb.close()

            samples: List[Dict[str, Any]] = []
            # 左面板: label=0, name=1, val=2, extra=6
            # 右面板: label=9, name=10, val=11, extra=15
            panel_configs = [(0, 1, 2, 6), (9, 10, 11, 15)]
            for cfg in panel_configs:
                w = parse_panel(*cfg, rows)
                if w.get("SO4"):
                    samples.append(w)

            # 补充野编号
            if len(rows) > 4:
                for cfg_idx, label_col in [(0, 0), (1, 9)]:
                    if cfg_idx < len(samples):
                        txt = (
                            safe_str(rows[4][label_col])
                            if len(rows[4]) > label_col
                            else ""
                        )
                        m = re.search(r"(\d+)", txt)
                        if m:
                            samples[cfg_idx]["field_id"] = m.group(1)

            logger.info(f"  水样: {len(samples)} 件")
            return samples
        except Exception as e:
            logger.error(f"  [!] 水样读取异常: {e}")
            return []

    # ---- 易溶盐 (支持 A3 双面板) ----

    def load_salt_samples(self) -> List[Dict[str, Any]]:
        """易溶盐土样 xlsx 解析 (支持 A3 双面板格式)"""
        path = find_file(self.excel_dir, "易溶盐")
        if not path:
            return []

        def parse_panel(
            name_a: int, val_a: int, name_b: int, val_b: int,
            rows: List[List[Any]],
        ) -> Dict[str, Any]:
            result: Dict[str, Any] = {}
            for r in rows:
                if len(r) <= val_b:
                    continue
                for nc, vc in [(name_a, val_a), (name_b, val_b)]:
                    if len(r) <= nc:
                        continue
                    ion = safe_str(r[nc])
                    v = safe_float(r[vc]) if len(r) > vc else None
                    if ion == "K+":
                        result["K"] = v
                    elif ion == "Na+":
                        result["Na"] = v
                    elif ion == "Ca2+":
                        result["Ca"] = v
                    elif ion == "Mg2+":
                        result["Mg"] = v
                    elif ion == "Cl-":
                        result["Cl"] = v
                    elif ion == "SO42-":
                        result["SO4"] = v
                    elif ion == "HCO3-":
                        result["HCO3"] = v
                    elif ion == "CO32-":
                        result["CO3"] = v or 0
                    elif ion == "OH-":
                        result["OH"] = v or 0

                # 检查总含盐量和 pH (左侧)
                txt0 = safe_str(r[name_a - 1]) if name_a > 0 and len(r) > name_a - 1 else ""
                if "总含盐量" in txt0:
                    result["TDS"] = safe_float(r[name_a]) if len(r) > name_a else None
                if "PH值" in txt0:
                    result["pH"] = safe_float(r[name_a]) if len(r) > name_a else None
                # 检查右侧
                txt_r = safe_str(r[name_a + 5]) if len(r) > name_a + 5 else ""
                if "总含盐量" in txt_r:
                    result["TDS"] = (
                        safe_float(r[name_a + 6]) if len(r) > name_a + 6 else None
                    )
                if "PH值" in txt_r:
                    result["pH"] = (
                        safe_float(r[name_a + 6]) if len(r) > name_a + 6 else None
                    )
            return result

        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.worksheets[0]
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            wb.close()

            samples: List[Dict[str, Any]] = []
            for cfg in [(0, 1, 3, 4), (7, 8, 10, 11)]:
                s = parse_panel(*cfg, rows)
                if s.get("SO4"):
                    samples.append(s)

            logger.info(f"  易溶盐: {len(samples)} 件")
            return samples
        except Exception as e:
            logger.error(f"  [!] 易溶盐读取异常: {e}")
            return []


# ============================================================
# 华宁勘察数据库读取器
# ============================================================

class HuaNingDBReader:
    """读取华宁 HNCAD 勘察软件原始数据库文件

    华宁数据库文件命名: FileType.ProjectCode (如 ZHMS.17, DCSH.17)
    编码: GB2312
    """

    # 成因代号 → 中文名
    ORIGIN_NAMES = {
        "ml": "人工堆积层",
        "al": "冲积层",
        "pl": "洪积层",
        "al+pl": "冲洪积层",
        "dl": "坡积层",
        "el": "残积层",
        "dl+el": "坡残积层",
        "eol": "风积层",
        "l": "湖积层",
        "m": "海积层",
        "mc": "海陆交互相沉积层",
        "fgl": "冰水沉积层",
    }

    # 年代代号 → 中文前缀
    AGE_PREFIXES = {
        "Q4": "第四系全新统",
        "Q3": "第四系上更新统",
        "Q2": "第四系中更新统",
        "Q1": "第四系下更新统",
        "N": "新近系",
        "Nh": "新元古界",
        "Pt": "元古界",
        "Ar": "太古界",
    }

    # 岩石代码 → 岩石名称 (5位岩性码 = 风化程度2位 + 岩石类型3位)
    ROCK_TYPE_MAP = {
        "150": "花岗片麻岩",
        "151": "花岗岩",
        "100": "砂岩",
        "101": "泥岩",
        "102": "页岩",
        "103": "石灰岩",
        "110": "片麻岩",
        "120": "大理岩",
        "130": "板岩",
        "140": "石英岩",
    }

    # 取样类型代号
    SAMPLE_TYPE_NAMES = {
        0: "扰动样",
        1: "原状样",
        2: "标贯试样",
        3: "水样",
        4: "岩样",
    }

    def __init__(self, db_dir: str, project_code: str = ""):
        self.db_dir = db_dir
        self.project_code = project_code
        if not project_code and db_dir:
            self.project_code = self._detect_project_code()

    # ---- 内部工具方法 ----

    def _detect_project_code(self) -> str:
        """自动检测项目代码"""
        for f in os.listdir(self.db_dir):
            if "." in f and f.upper().startswith("ZHMS"):
                return f.split(".")[-1]
        for f in os.listdir(self.db_dir):
            if "." in f and not f.startswith("_") and not f.startswith("."):
                parts = f.split(".")
                if len(parts) == 2 and parts[1].isdigit():
                    return parts[1]
        return ""

    def _find_file(self, file_type: str) -> Optional[str]:
        """查找数据库文件 (大小写不敏感)"""
        target_upper = f"{file_type.upper()}.{self.project_code}"
        for f in os.listdir(self.db_dir):
            if f.upper() == target_upper:
                return os.path.join(self.db_dir, f)
        # 不区分扩展名
        for f in os.listdir(self.db_dir):
            base = f.split(".")[0].upper()
            if base == file_type.upper() and os.path.isfile(os.path.join(self.db_dir, f)):
                return os.path.join(self.db_dir, f)
        return None

    def _read_lines(self, file_type: str) -> List[List[str]]:
        """读取逗号分隔的数据库文件，自动检测编码"""
        path = self._find_file(file_type)
        if not path:
            return []
        content = None
        for enc in ("gb2312", "gbk", "utf-8"):
            try:
                with open(path, "r", encoding=enc) as f:
                    content = f.read()
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        if content is None:
            logger.warning(f"    无法读取 {file_type}: 编码不支持")
            return []
        lines: List[List[str]] = []
        for line in content.strip().split("\n"):
            line = line.strip()
            if line:
                lines.append(line.split(","))
        return lines

    # ---- ZHMS: 地层描述 ----

    def read_layer_descriptions(self) -> Dict[str, Dict[str, str]]:
        """读取 ZHMS 文件 → {layer_id: {name, description}}

        格式: 层号,岩土名:描述内容
        """
        lines = self._read_lines("ZHMS")
        result: Dict[str, Dict[str, str]] = {}
        for fields in lines:
            if len(fields) < 2:
                continue
            layer_id = fields[0].strip()
            rest = ",".join(fields[1:])  # 合并可能因逗号分割的内容
            if layer_id.upper() == "END":
                break
            colon_idx = rest.find(":")
            if colon_idx > 0:
                name = rest[:colon_idx].strip()
                desc = rest[colon_idx + 1:].strip()
            else:
                name = rest.strip()
                desc = ""
            if layer_id and name:
                result[layer_id] = {"name": name, "description": desc}
        return result

    # ---- DCSH: 地层序列与地质年代 ----

    def read_layer_sequence(self) -> List[Dict[str, str]]:
        """读取 DCSH 文件 → 地层序列 (含地质年代和成因)

        格式: 层号,岩性代码,,,成因,年代,
        """
        lines = self._read_lines("DCSH")
        result: List[Dict[str, str]] = []
        for fields in lines:
            if len(fields) < 2:
                continue
            layer_id = fields[0].strip()
            if layer_id.upper() == "END":
                break
            rock_code = fields[1].strip() if len(fields) > 1 else ""
            origin = fields[4].strip() if len(fields) > 4 else ""
            age = fields[5].strip() if len(fields) > 5 else ""
            if layer_id:
                result.append({
                    "layer_id": layer_id,
                    "rock_code": rock_code,
                    "origin": origin,
                    "age": age,
                })
        return result

    # ---- DCSJ: 逐孔地层数据 ----

    def read_borehole_layers(self) -> Tuple[
        Dict[str, List[Dict[str, Any]]],
        set,
        Dict[str, float],
    ]:
        """读取 DCSJ 文件 → (borehole_layers, unpentrated_ids, max_exposed)

        格式: 孔号,层号,层底深度,
        最后一层无深度 → 未穿透
        """
        lines = self._read_lines("DCSJ")
        bh_layers: Dict[str, List[Dict[str, Any]]] = {}
        max_exposed: Dict[str, float] = {}
        current_bh: Optional[str] = None

        for fields in lines:
            if not fields:
                continue
            bh_id = fields[0].strip() if fields[0].strip() else ""
            if bh_id and bh_id.upper() != "END":
                current_bh = bh_id
            elif bh_id.upper() == "END":
                break
            if not current_bh or len(fields) < 2:
                continue

            layer_id = fields[1].strip()
            depth_str = fields[2].strip() if len(fields) > 2 else ""

            if current_bh not in bh_layers:
                bh_layers[current_bh] = []

            depth: Optional[float] = None
            if depth_str:
                try:
                    depth = float(depth_str)
                except ValueError:
                    depth = None

            bh_layers[current_bh].append({
                "layer_id": layer_id,
                "depth": depth,
            })

            # 跟踪每层的最大揭露深度
            if depth is not None:
                prev = max_exposed.get(layer_id, 0)
                if depth > prev:
                    max_exposed[layer_id] = depth

        # 检测未穿透层: 某钻孔最后一层无深度
        unpentrated: set = set()
        for bh_id, layers in bh_layers.items():
            if layers and layers[-1]["depth"] is None:
                unpentrated.add(layers[-1]["layer_id"])

        return bh_layers, unpentrated, max_exposed

    # ---- BG: 标贯数据 ----

    def read_spt_data(self) -> List[Dict[str, Any]]:
        """读取 BG 文件 → SPT 数据列表

        格式: 孔号,起始深度,终止深度,N值
        (孔号为空表示续上一行孔号)
        """
        lines = self._read_lines("BG")
        result: List[Dict[str, Any]] = []
        current_bh: Optional[str] = None

        for fields in lines:
            if not fields:
                continue
            bh_id = fields[0].strip() if fields[0].strip() else ""
            if bh_id and bh_id.upper() != "END":
                current_bh = bh_id
            elif bh_id.upper() == "END":
                break
            if not current_bh or len(fields) < 4:
                continue

            try:
                depth_start = float(fields[1])
                depth_end = float(fields[2])
                n_value = float(fields[3])
            except (ValueError, IndexError):
                continue

            result.append({
                "borehole_id": current_bh,
                "depth_start": depth_start,
                "depth_end": depth_end,
                "n_value": n_value,
            })

        return result

    # ---- TY: 取样数据 ----

    def read_sample_data(self) -> List[Dict[str, Any]]:
        """读取 TY 文件 → 取样数据列表

        格式: 孔号,深度,类型代码
        类型: 0=扰动样, 1=原状样, 2=标贯试样
        """
        lines = self._read_lines("TY")
        result: List[Dict[str, Any]] = []
        current_bh: Optional[str] = None

        for fields in lines:
            if not fields:
                continue
            bh_id = fields[0].strip() if fields[0].strip() else ""
            if bh_id and bh_id.upper() != "END":
                current_bh = bh_id
            elif bh_id.upper() == "END":
                break
            if not current_bh or len(fields) < 3:
                continue

            try:
                depth = float(fields[1])
                sample_type = int(fields[2])
            except (ValueError, IndexError):
                continue

            result.append({
                "borehole_id": current_bh,
                "depth": depth,
                "sample_type": sample_type,
            })

        return result

    # ---- 数据汇总 ----

    def _assign_to_layer(
        self,
        depth: float,
        layer_boundaries: List[Tuple[str, Optional[float]]],
    ) -> Optional[str]:
        """根据深度确定所属地层

        layer_boundaries: [(layer_id, bottom_depth), ...]
        """
        prev_depth = 0.0
        for lid, bdepth in layer_boundaries:
            if bdepth is None:
                # 未穿透层, depth 落在该层
                return lid
            if depth <= bdepth:
                return lid
            prev_depth = bdepth
        return None

    def compute_spt_counts(
        self,
        spt_records: List[Dict[str, Any]],
        bh_layers: Dict[str, List[Dict[str, Any]]],
    ) -> Dict[str, int]:
        """统计每层标贯次数"""
        counts: Dict[str, int] = defaultdict(int)
        for rec in spt_records:
            bh_id = rec["borehole_id"]
            midpoint = (rec["depth_start"] + rec["depth_end"]) / 2.0
            boundaries = [
                (l["layer_id"], l["depth"])
                for l in bh_layers.get(bh_id, [])
            ]
            lid = self._assign_to_layer(midpoint, boundaries)
            if lid:
                counts[lid] += 1
        return dict(counts)

    def compute_sample_counts(
        self,
        sample_records: List[Dict[str, Any]],
        bh_layers: Dict[str, List[Dict[str, Any]]],
    ) -> Dict[str, Dict[int, int]]:
        """统计每层各类取样数量

        返回: {layer_id: {sample_type_code: count}}
        """
        counts: Dict[str, Dict[int, int]] = defaultdict(lambda: defaultdict(int))
        for rec in sample_records:
            bh_id = rec["borehole_id"]
            depth = rec["depth"]
            stype = rec["sample_type"]
            boundaries = [
                (l["layer_id"], l["depth"])
                for l in bh_layers.get(bh_id, [])
            ]
            lid = self._assign_to_layer(depth, boundaries)
            if lid:
                counts[lid][stype] += 1
        return {lid: dict(types) for lid, types in counts.items()}

    def build_age_groups(
        self,
        layer_sequence: List[Dict[str, str]],
    ) -> List[Tuple[str, List[str]]]:
        """构建地质年代分组 (保持 DCSH 顺序)

        返回: [(group_key, [layer_ids]), ...]
        group_key 格式: "age|origin" 或 "age" (无成因时)
        """
        groups: List[Tuple[str, List[str]]] = []
        current_key: Optional[str] = None

        for rec in layer_sequence:
            age = rec.get("age", "")
            origin = rec.get("origin", "")
            layer_id = rec["layer_id"]

            # 基岩层 origin 为纯数字 (如 "1") 时视为无成因
            if origin and not origin.isdigit():
                key = f"{age}|{origin}"
            else:
                key = age

            if key != current_key:
                groups.append((key, []))
                current_key = key

            groups[-1][1].append(layer_id)

        return groups

    def build_age_display_name(
        self,
        group_key: str,
        layer_names: Dict[str, str],
        config_overrides: Optional[Dict[str, str]] = None,
    ) -> str:
        """构建地质年代显示名 (如 '第四系人工堆积层（Q4ml）')"""
        # 先检查配置覆盖
        if config_overrides and group_key in config_overrides:
            return config_overrides[group_key]

        parts = group_key.split("|")
        age = parts[0]
        origin = parts[1] if len(parts) > 1 else ""

        # 基岩 (无成因)
        if not origin:
            rock_name = ""
            for lid in layer_names:
                name = layer_names.get(lid, "")
                for prefix in ("强风化", "中风化", "微风化", "碎块状强风化",
                               "全风化", "弱风化"):
                    if name.startswith(prefix):
                        rock_name = name[len(prefix):]
                        break
                if rock_name:
                    break
            # 尝试从岩石代码映射
            if not rock_name:
                rock_name = self.ROCK_TYPE_MAP.get("150", "花岗片麻岩")
            return f"{self.AGE_PREFIXES.get(age, age)}{rock_name}（{age}）"

        # 第四系沉积层
        age_prefix = self.AGE_PREFIXES.get(age, age)
        origin_name = self.ORIGIN_NAMES.get(origin, origin + "层")
        return f"{age_prefix}{origin_name}（{age}{origin}）"

    # ---- 主读取入口 ----

    def read(self) -> Dict[str, Any]:
        """读取华宁数据库全部数据

        返回:
        {
            "available": True/False,
            "descriptions": {layer_id: {name, description}},
            "layer_sequence": [{layer_id, rock_code, origin, age}],
            "age_groups": [(group_key, [layer_ids])],
            "borehole_layers": {bh_id: [{layer_id, depth}]},
            "unpentrated": set(),
            "max_exposed": {layer_id: max_depth},
            "spt_records": [...],
            "sample_records": [...],
            "spt_counts": {layer_id: count},
            "sample_counts": {layer_id: {type: count}},
        }
        """
        descriptions = self.read_layer_descriptions()
        if not descriptions:
            logger.info("    华宁数据库: ZHMS 文件不存在或为空")
            return {"available": False}

        layer_sequence = self.read_layer_sequence()
        bh_layers, unpentrated, max_exposed = self.read_borehole_layers()
        spt_records = self.read_spt_data()
        sample_records = self.read_sample_data()

        age_groups = self.build_age_groups(layer_sequence)
        spt_counts = self.compute_spt_counts(spt_records, bh_layers)
        sample_counts = self.compute_sample_counts(sample_records, bh_layers)

        logger.info(
            f"    华宁数据库: {len(descriptions)} 层描述, "
            f"{len(layer_sequence)} 层序列, "
            f"{len(bh_layers)} 孔, "
            f"{len(spt_records)} 条标贯, "
            f"{len(sample_records)} 件取样"
        )

        return {
            "available": True,
            "descriptions": descriptions,
            "layer_sequence": layer_sequence,
            "age_groups": age_groups,
            "borehole_layers": bh_layers,
            "unpentrated": unpentrated,
            "max_exposed": max_exposed,
            "spt_records": spt_records,
            "sample_records": sample_records,
            "spt_counts": spt_counts,
            "sample_counts": sample_counts,
        }


# ============================================================
# 报告填充器
# ============================================================

class ReportFiller:
    """将勘察数据填充到 .docx 报告模板"""

    def __init__(
        self,
        template_path: str,
        output_path: str,
        data: Dict[str, Any],
        layer_names: Dict[str, str],
        layer_ids: List[str],
        config: ProjectConfig,
    ):
        self.template = template_path
        self.output = output_path
        self.data = data
        self.layer_names = layer_names
        self.layer_ids = layer_ids
        self.config = config
        self.ti = config.table_indices  # 表格索引快捷引用
        self.doc = Document(template_path)

    def save(self) -> None:
        """保存生成的报告"""
        self.doc.save(self.output)

    def fill_all(self) -> None:
        """执行全部填充步骤"""
        logger.info("")
        logger.info("=" * 60)
        logger.info("  开始填充报告（只改文字，不动格式）")
        logger.info("=" * 60)

        self._global_replace()
        self._fill_project_overview()       # 1. 第一章: 拟建工程概况
        self._fill_survey_purpose()         # 2. 第二章: 勘察目的条件过滤
        self._fill_buildings_table()        # 3. 建筑物特征表
        self._fill_workload()
        self._fill_water_level()
        self._fill_layer_descriptions()
        self._fill_phys_spt_tables()
        self._fill_bearing_capacity()
        self._fill_water_salt_tables()
        self._fill_corrosion_eval()
        self._fill_liquefaction()
        self._fill_foundation_tables()
        self._fill_conclusion()
        self._fill_site_conditions()
        self._fill_site_evaluation()
        self._fill_foundation_evaluation()
        self._fill_analysis_evaluation()    # 16. 第六章: 分析评价各子节
        self._fill_standards()
        self._apply_date_replacements()

        logger.info(f"\n  保存: {self.output}")

    # ---- 全局文本替换 ----

    def _global_replace(self) -> None:
        """按配置执行全局文本替换"""
        replacements = self.config.replacements
        if not replacements:
            return

        logger.info("  全局文本替换...")
        for old, new in replacements:
            for p in self.doc.paragraphs:
                replace_in_para(p, old, new)
            for t in self.doc.tables:
                for row in t.rows:
                    for cell in row.cells:
                        for cp in cell.paragraphs:
                            replace_in_para(cp, old, new)

    # ---- 第一章: 拟建工程概况 ----

    def _fill_project_overview(self) -> None:
        """填充拟建工程概况章节（第一章）

        配置字段:
            commission_text:    委托段落全文（支持 {client}/{project_name}/{survey_stage} 占位符）
            site_location:      场地位置描述段落
            building_desc:      建筑物描述段落（含面积、层数等）
            importance_level:   重要性等级（一级/二级/三级）
            site_complexity:    场地复杂程度等级（一级/二级/三级）
            foundation_grade:   地基等级（一级/二级/三级）
            survey_grade:       岩土工程勘察等级（甲级/乙级/丙级）
            survey_grade_text:  勘察等级完整描述（覆盖自动拼装）
        """
        overview = self.config.get_project_overview()
        if not overview:
            return

        # 准备占位符
        fmt_vars: Dict[str, str] = {
            "client": overview.get("client", ""),
            "project_name": overview.get("project_name", self.config.project_name),
            "survey_stage": overview.get("survey_stage", "详细勘察"),
        }

        # 自动拼装勘察等级文字
        grade_parts = []
        for key, label in [
            ("importance_level", "重要性等级"),
            ("site_complexity", "场地复杂程度等级"),
            ("foundation_grade", "地基等级"),
        ]:
            val = overview.get(key, "")
            if val:
                grade_parts.append(f"{label}{val}")
        survey_grade = overview.get("survey_grade", "")
        if survey_grade:
            grade_parts.append(f"综合确定岩土工程勘察等级为{survey_grade}")
        auto_grade_text = "，".join(grade_parts) + "。" if grade_parts else ""
        grade_text = overview.get("survey_grade_text", auto_grade_text)

        # 在模板中定位并替换
        commission_text = overview.get("commission_text", "")
        site_location = overview.get("site_location", "")
        building_desc = overview.get("building_desc", "")

        replaced_count = 0
        for p in self.doc.paragraphs:
            txt = p.text.strip()
            if not txt:
                continue

            # 1. 委托段落: 匹配 "受XX的委托" 或 "承担了" 开头
            if commission_text and ("的委托" in txt or "承担了" in txt):
                try:
                    full_text = commission_text.format(**fmt_vars)
                except KeyError:
                    full_text = commission_text
                set_para_text(p, full_text)
                replaced_count += 1
                continue

            # 2. 场地位置: 匹配 "勘察场地位于" 开头
            if site_location and txt.startswith("勘察场地位于"):
                set_para_text(p, site_location)
                replaced_count += 1
                continue

            # 3. 建筑物描述: 匹配 "本次勘察建筑物" 或 "总建筑面积"
            if building_desc and (
                "本次勘察建筑物" in txt or "总建筑面积" in txt
            ):
                set_para_text(p, building_desc)
                replaced_count += 1
                continue

            # 4. 勘察等级: 匹配 "岩土工程勘察等级"
            if grade_text and "岩土工程勘察等级" in txt:
                set_para_text(p, grade_text)
                replaced_count += 1
                continue

            # 5. 标准冻结深度 (第四章): 匹配 "标准冻结深度"
            frozen_depth = overview.get("frozen_depth", "")
            if frozen_depth and "标准冻结深度" in txt:
                set_para_text(p, f"场地土标准冻结深度{frozen_depth}。")
                replaced_count += 1
                continue

        if replaced_count:
            logger.info(f"  工程概况: 替换 {replaced_count} 段")
        else:
            logger.debug("  工程概况: 未找到匹配段落")

    # ---- 勘察目的 / 任务要求条件过滤 (第二章) ----

    def _fill_survey_purpose(self) -> None:
        """根据条件过滤勘察目的/任务要求条目（第二章）

        当无地下室 (has_basement=False) 时，删除含"基坑开挖"的条目并重新编号。
        可通过配置 project_overview.has_basement 显式控制。
        """
        overview = self.config.get_project_overview()
        if not overview:
            return

        # has_basement 判断: 优先取配置显式值, 否则从建筑物数据自动检测
        conditions = self._evaluate_standard_conditions()
        has_basement = overview.get("has_basement", conditions.get("has_basement", False))

        if has_basement:
            return  # 有地下室/基坑, 全部保留

        # 条件关键词: 包含此关键词的条目在条件不满足时删除
        cond_keyword = "基坑开挖"

        # 收集编号条目: [(para_index, item_number)]
        numbered_items: List[Tuple[int, int]] = []
        removed_indices: List[int] = []
        num_idx = 1
        for i, p in enumerate(self.doc.paragraphs):
            txt = p.text.strip()
            if not txt:
                continue
            prefix = f"{num_idx}."
            if txt.startswith(prefix):
                if cond_keyword in txt:
                    removed_indices.append(i)
                    # 仍加入列表以便后续重编号 (最终会被清空)
                    numbered_items.append((i, num_idx))
                else:
                    numbered_items.append((i, num_idx))
                num_idx += 1

        if not removed_indices:
            return

        # 清空需要删除的段落
        for idx in removed_indices:
            set_para_text(self.doc.paragraphs[idx], "")

        # 重新编号: 剩余条目按顺序 1,2,3...
        remaining = [(idx, n) for idx, n in numbered_items if idx not in removed_indices]
        for seq, (idx, old_num) in enumerate(remaining, 1):
            p = self.doc.paragraphs[idx]
            old_prefix = f"{old_num}."
            new_prefix = f"{seq}."
            new_text = new_prefix + p.text.strip()[len(old_prefix):]
            set_para_text(p, new_text)

        logger.info(f"  勘察目的: 删除基坑相关条目 {len(removed_indices)} 条, 重编号 {len(remaining)} 条")

    # ---- 建筑物特征表 ----

    def _fill_buildings_table(self) -> None:
        buildings = self.data.get("buildings", [])
        if not buildings:
            return

        logger.info(f"  建筑物特征表 ({len(buildings)} 栋)...")
        idx = self.ti["buildings"]
        t = self.doc.tables[idx] if len(self.doc.tables) > idx else None
        if not t:
            return

        building_keys = ["name", "floors", "height", "size", "span", "indoor_elv", "structure"]
        for bi, b in enumerate(buildings):
            if bi + 2 >= len(t.rows):
                break
            for ci, key in enumerate(building_keys):
                set_cell(t, bi + 2, ci, b.get(key, ""))

    # ---- 工作量表 + 段落 ----

    def _fill_workload(self) -> None:
        logger.info("  工作量表 + 段落...")
        bh_info = self.data["borehole_info"]

        # 工作量统计表
        idx = self.ti["workload"]
        t = self.doc.tables[idx] if len(self.doc.tables) > idx else None
        if t:
            set_cell(t, 1, 3, str(bh_info["total"]))
            set_cell(t, 2, 3, f"{bh_info['total_depth']:.0f}/{bh_info['total']}")
            set_cell(t, 3, 3, str(bh_info["total"]))
            set_cell(t, 4, 3, str(bh_info["undisturbed"]))
            set_cell(t, 5, 3, str(bh_info["disturbed"]))
            set_cell(t, 6, 3, str(bh_info["rock"]))
            set_cell(t, 7, 3, str(bh_info["spt_total"]))
            set_cell(t, 8, 3, str(bh_info["n63_total"]))
            if len(t.rows) > 9:
                set_cell(t, 9, 3, str(len(self.data.get("water_samples", []))))
            if len(t.rows) > 10:
                set_cell(t, 10, 3, str(len(self.data.get("salt_samples", []))))
            if len(t.rows) > 11:
                set_cell(t, 11, 3, str(bh_info.get("bosk", 0)))

        # 工作量段落
        for p in self.doc.paragraphs:
            txt = p.text.strip()
            if txt.startswith("布孔") and "钻孔" in txt:
                set_para_text(p, (
                    f"布孔：沿拟建物周边及角点并结合网度布设钻孔{bh_info['total']}个，"
                    f"间距16.0～20.00m；其中控制性钻孔{bh_info['ctrl']}个、"
                    f"一般性钻孔{bh_info['general']}个。"
                ))
            elif "钻孔类型" in txt and "取土孔" in txt:
                set_para_text(p, (
                    f"钻孔类型：取土孔{bh_info['qutu']}个、标贯孔{bh_info['biaoguan']}个，"
                    f"鉴别孔{bh_info['yiban']}个，波速孔{bh_info.get('bosk', 0)}个。"
                ))
            elif "钻孔孔深" in txt and "强风化" in txt:
                set_para_text(p, (
                    "钻孔孔深：控制性钻孔进入强风化岩面以下7-9m、"
                    "一般性钻孔进入强风化岩面以下不小于6m。"
                ))
            elif "水样及易溶盐" in txt:
                n_water = len(self.data.get("water_samples", []))
                n_salt = len(self.data.get("salt_samples", []))
                set_para_text(p, f"水样及易溶盐土样：地下水水样{n_water}件、易溶盐土样{n_salt}件。")
            elif "本次勘察实际完成" in txt:
                set_para_text(p, (
                    f"本次勘察实际完成钻孔{bh_info['total']}个，"
                    f"其中控制孔{bh_info['ctrl']}个，一般孔{bh_info['general']}个。"
                    f"其中取土孔{bh_info['qutu']}个、标贯孔{bh_info['biaoguan']}个，"
                    f"鉴别孔{bh_info['yiban']}个，波速孔{bh_info.get('bosk', 0)}个，"
                    f"勘察工作量统计表见表3-1。"
                ))

        # --- 勘察方法条件过滤 (第三章(一)) ---
        # N63.5 动力触探: 无数据时删除
        has_n63 = bh_info.get("n63_total", 0) > 0

        # 波速测试: 无高层建筑 (>24m) 时删除
        buildings = self.data.get("buildings", [])
        has_highrise = False
        for b in buildings:
            h_str = str(b.get("height", ""))
            # 提取数字部分 (兼容 "36m", "36.5", "18F" 等格式)
            h_match = re.match(r"(\d+\.?\d*)", h_str)
            h_val = float(h_match.group(1)) if h_match else 0
            if h_val > 24:
                has_highrise = True
                break

        if not has_n63 or not has_highrise:
            cleared = 0
            for p in self.doc.paragraphs:
                txt = p.text.strip()
                if not txt:
                    continue
                if not has_n63 and ("N63.5" in txt or "动力触探" in txt):
                    set_para_text(p, "")
                    cleared += 1
                elif not has_highrise and "波速测试" in txt:
                    set_para_text(p, "")
                    cleared += 1
            if cleared:
                logger.info(f"    勘察方法: 条件删除 {cleared} 段 (N63.5={has_n63}, 高层={has_highrise})")

        # --- 质量评述: 勘察等级注入 (第三章(五)) ---
        overview = self.config.get_project_overview()
        survey_grade = (overview or {}).get("survey_grade", "")
        if survey_grade:
            for p in self.doc.paragraphs:
                txt = p.text.strip()
                if "勘察等级为" in txt:
                    for grade in ("甲级", "乙级", "丙级"):
                        if grade in txt:
                            replace_in_para(p, grade, survey_grade)
                            break

    # ---- 水位表 + 段落 ----

    def _fill_water_level(self) -> None:
        logger.info("  水位表 + 段落...")
        bh_info = self.data["borehole_info"]
        if "wt_depth_min" not in bh_info:
            return

        idx = self.ti["water_level"]
        t = self.doc.tables[idx] if len(self.doc.tables) > idx else None

        if t:
            wtd = [
                bh["wt_depth"]
                for bh in self.data.get("boreholes", [])
                if bh.get("wt_depth")
            ]
            wte = [
                bh["wt_elv"]
                for bh in self.data.get("boreholes", [])
                if bh.get("wt_elv")
            ]
            if wtd:
                set_cell(t, 1, 0, str(len(wtd)))
                set_cell(t, 1, 1, f"{min(wtd):.2f}")
                set_cell(t, 1, 2, f"{max(wtd):.2f}")
                set_cell(t, 1, 3, f"{sum(wtd) / len(wtd):.2f}")
                set_cell(t, 1, 4, f"{min(wte):.2f}")
                set_cell(t, 1, 5, f"{max(wte):.2f}")
                set_cell(t, 1, 6, f"{sum(wte) / len(wte):.2f}")

        for p in self.doc.paragraphs:
            if "勘察期间测得钻孔内水位埋深" in p.text and "wt_depth_min" in bh_info:
                set_para_text(p, (
                    f"勘察期间测得钻孔内水位埋深"
                    f"{bh_info['wt_depth_min']:.2f}~{bh_info['wt_depth_max']:.2f}m，"
                    f"水位标高"
                    f"{bh_info['wt_elv_min']:.2f}~{bh_info['wt_elv_max']:.2f}m，"
                    f"详见表5-13。"
                ))
                break

    # ---- 地层描述段落 (第五章(二) 岩土结构及工程特性) ----

    def _build_layer_display_name(
        self, lid: str, name: str, geo_code: str,
    ) -> str:
        """构建地层段落的显示名称

        格式: {lid}{name}（{geo_code}）
        例: 1杂填土（Q4ml）, 5-1强风化花岗片麻岩（NhηγRw）
        """
        if geo_code:
            return f"{lid}{name}（{geo_code}）"
        return f"{lid}层{name}"

    def _format_layer_stats_text(
        self,
        lid: str,
        layer_stats: Dict[str, Any],
        hn_data: Dict[str, Any],
        bh_layers: Dict[str, List[Dict[str, Any]]],
    ) -> str:
        """构建地层段落的统计文本 (厚度/标高/埋深 或 未穿透)"""
        unpentrated = hn_data.get("unpentrated", set())
        max_exposed = hn_data.get("max_exposed", {})
        ldata = layer_stats.get(lid, {})

        if lid in unpentrated:
            # 未穿透: 使用最大揭露厚度
            max_d = max_exposed.get(lid, 0)
            # 计算分布
            bh_count = sum(
                1 for layers in bh_layers.values()
                if any(l["layer_id"] == lid for l in layers)
            )
            total_bh = len(bh_layers) or 1
            ratio = bh_count / total_bh
            dist = "普遍" if ratio > 0.5 else "较普遍" if ratio > 0.2 else "局部"
            return (
                f"该层在场区{dist}分布，未穿透，"
                f"最大揭露厚度{max_d:.2f}m。"
            )

        # 正常穿透层: 厚度/标高/埋深统计
        if not ldata:
            return ""

        # 计算分布
        bh_count = sum(
            1 for layers in bh_layers.values()
            if any(l["layer_id"] == lid for l in layers)
        )
        total_bh = len(bh_layers) or 1
        ratio = bh_count / total_bh
        dist = "普遍" if ratio > 0.5 else "较普遍" if ratio > 0.2 else "局部"

        return (
            f"场区{dist}分布，"
            f"厚度:{fmt_val(ldata.get('thick_min'))}~"
            f"{fmt_val(ldata.get('thick_max'))}m,"
            f"平均{fmt_val(ldata.get('thick_avg'))}m;"
            f"层底标高:{fmt_val(ldata.get('elv_min'))}~"
            f"{fmt_val(ldata.get('elv_max'))}m,"
            f"平均{fmt_val(ldata.get('elv_avg'))}m;"
            f"层底埋深:{fmt_val(ldata.get('depth_min'))}~"
            f"{fmt_val(ldata.get('depth_max'))}m,"
            f"平均{fmt_val(ldata.get('depth_avg'))}m。"
        )

    def _format_test_info_text(
        self,
        lid: str,
        table_seq: int,
        hn_data: Dict[str, Any],
    ) -> str:
        """构建试验信息段落文字"""
        spt_count = hn_data.get("spt_counts", {}).get(lid, 0)
        sample_types = hn_data.get("sample_counts", {}).get(lid, {})

        type_0 = sample_types.get(0, 0)  # 扰动样
        type_1 = sample_types.get(1, 0)  # 原状样

        parts: List[str] = []
        if type_1 > 0:
            parts.append(f"取原状样{type_1}件")
        if type_0 > 0:
            parts.append(f"取扰动样{type_0}件")
        if spt_count > 0:
            parts.append(f"进行标准贯入试验{spt_count}次")

        if not parts:
            return ""

        table_ref = f"，有关工程特性指标见表5-{table_seq}" if table_seq > 0 else ""
        return "该层" + "，".join(parts) + table_ref + "。"

    def _make_bold_paragraph(self, text: str, template_para: Any) -> Any:
        """创建一个加粗段落 (复制模板段落的格式)"""
        new_p = OxmlElement("w:p")

        # 复制段落属性 (pPr) 并设置加粗
        if template_para._element.find(qn("w:pPr")) is not None:
            pPr = copy.deepcopy(template_para._element.find(qn("w:pPr")))
            new_p.append(pPr)

        # 创建 run
        new_r = OxmlElement("w:r")

        # 复制 run 属性 (rPr) 并强制加粗
        if template_para.runs:
            src_rPr = template_para.runs[0]._element.find(qn("w:rPr"))
            if src_rPr is not None:
                rPr = copy.deepcopy(src_rPr)
                # 确保有 w:b 加粗标记
                if rPr.find(qn("w:b")) is None:
                    b_elem = OxmlElement("w:b")
                    rPr.insert(0, b_elem)
                new_r.append(rPr)
            else:
                rPr = OxmlElement("w:rPr")
                b_elem = OxmlElement("w:b")
                rPr.append(b_elem)
                new_r.append(rPr)
        else:
            rPr = OxmlElement("w:rPr")
            b_elem = OxmlElement("w:b")
            rPr.append(b_elem)
            new_r.append(rPr)

        # 设置文字
        new_t = OxmlElement("w:t")
        new_t.text = text
        new_t.set(qn("xml:space"), "preserve")
        new_r.append(new_t)
        new_p.append(new_r)

        return new_p

    def _make_normal_paragraph(self, text: str, template_para: Any) -> Any:
        """创建一个普通段落 (复制模板段落的格式, 不加粗)"""
        new_p = OxmlElement("w:p")

        # 复制段落属性 (pPr)
        if template_para._element.find(qn("w:pPr")) is not None:
            pPr = copy.deepcopy(template_para._element.find(qn("w:pPr")))
            # 移除 pPr 中的加粗 (pPr/rPr/w:b)
            pPr_rPr = pPr.find(qn("w:rPr"))
            if pPr_rPr is not None:
                b_elem = pPr_rPr.find(qn("w:b"))
                if b_elem is not None:
                    pPr_rPr.remove(b_elem)
            new_p.append(pPr)

        # 创建 run
        new_r = OxmlElement("w:r")

        # 复制 run 属性 (rPr), 但确保不加粗
        if template_para.runs:
            src_rPr = template_para.runs[0]._element.find(qn("w:rPr"))
            if src_rPr is not None:
                rPr = copy.deepcopy(src_rPr)
                # 确保没有 w:b 加粗标记
                b_elem = rPr.find(qn("w:b"))
                if b_elem is not None:
                    rPr.remove(b_elem)
                new_r.append(rPr)

        # 设置文字
        new_t = OxmlElement("w:t")
        new_t.text = text
        new_t.set(qn("xml:space"), "preserve")
        new_r.append(new_t)
        new_p.append(new_r)

        return new_p

    def _fill_layer_descriptions(self) -> None:
        """填充第五章(二) 岩土结构及工程特性

        优先使用华宁数据库数据 (ZHMS/DCSH/BG/TY/DCSJ) 动态生成;
        若无数据库则回退到模板匹配模式。
        """
        hn_data = self.data.get("hn_data", {})
        layer_stats = self.data.get("layers", {})

        if not hn_data or not hn_data.get("available"):
            logger.info("  地层描述段落... (模板模式)")
            self._fill_layer_descriptions_fallback(layer_stats)
            return

        logger.info("  地层描述段落... (华宁数据库模式)")

        descriptions = hn_data.get("descriptions", {})
        age_groups = hn_data.get("age_groups", [])
        bh_layers = hn_data.get("borehole_layers", {})
        unpentrated = hn_data.get("unpentrated", set())

        # 从 DCSH 层序列构建层号→地质代号映射
        geo_codes: Dict[str, str] = {}
        for rec in hn_data.get("layer_sequence", []):
            lid = rec["layer_id"]
            age = rec.get("age", "")
            origin = rec.get("origin", "")
            # 基岩层 origin 为纯数字时不计入代号
            if origin and not origin.isdigit():
                geo_codes[lid] = f"{age}{origin}"
            elif age:
                geo_codes[lid] = age

        # 获取地质年代名称配置覆盖
        geo_age_overrides = self.config.get_geo_age_names()

        # 定位 (二) 和 (三) 节
        section_start = None
        section_end = None
        intro_para = None
        closing_para = None
        intro_idx = None
        closing_idx = None

        for i, p in enumerate(self.doc.paragraphs):
            txt = p.text.strip()
            if not txt:
                continue
            if re.match(r"[（(]二[）)]", txt):
                section_start = i
            elif section_start is not None and re.match(r"[（(]三[）)]", txt):
                section_end = i
                break

        if section_start is None:
            logger.warning("    未找到(二)节标记, 回退到模板模式")
            self._fill_layer_descriptions_fallback(layer_stats)
            return

        # 查找 intro 和 closing
        if section_start + 1 < len(self.doc.paragraphs):
            intro_para = self.doc.paragraphs[section_start + 1]
            intro_idx = section_start + 1

        if section_end is not None:
            closing_idx = section_end - 1
            if closing_idx > (intro_idx or section_start):
                closing_para = self.doc.paragraphs[closing_idx]

        # 确定插入锚点和需要清除的范围
        body = self.doc.element.body
        anchor = intro_para._element if intro_para else \
            self.doc.paragraphs[section_start]._element

        # 清除 intro 与 (三) 之间的所有内容 (段落+表格)
        if intro_para and section_end is not None:
            start_elem = intro_para._element
            end_elem = self.doc.paragraphs[section_end]._element
            elems_to_remove = []
            clearing = False
            for child in body:
                if child is start_elem:
                    clearing = True
                    continue
                if child is end_elem:
                    break
                if clearing:
                    elems_to_remove.append(child)
            for elem in elems_to_remove:
                body.remove(elem)

        # 保留 closing 段落 (如果有)
        # 否则在末尾添加
        if closing_para is not None and closing_para._element.getparent() is not None:
            # closing 段落还在, 移到 section_end 之前
            pass
        else:
            closing_para = None

        # 生成新段落内容
        hn_reader = HuaNingDBReader("", "")  # 仅用于静态方法
        table_seq = 0
        inserted_elements: List[Any] = []

        for group_key, layer_ids_in_group in age_groups:
            # 1) 地质年代标题 (加粗)
            # 获取层名映射 (合并 HN 描述和现有 layer_names)
            merged_names = dict(self.layer_names)
            for lid, info in descriptions.items():
                if lid not in merged_names:
                    merged_names[lid] = info["name"]

            age_display = hn_reader.build_age_display_name(
                group_key, merged_names, geo_age_overrides,
            )
            bold_p = self._make_bold_paragraph(age_display, intro_para)
            inserted_elements.append(bold_p)

            # 2) 每层段落
            for lid in layer_ids_in_group:
                desc_info = descriptions.get(lid, {})
                name = desc_info.get("name", self.layer_names.get(lid, f"第{lid}层"))
                desc_text = desc_info.get("description", "")
                geo_code = geo_codes.get(lid, "")

                # 使用配置文件的描述覆盖 (如果有)
                config_desc = self.config.get_layer_description(lid)
                if config_desc:
                    # 配置描述是完整模板, 可含 {thick_*} 占位符
                    ldata = layer_stats.get(lid, {})
                    try:
                        desc_text = config_desc.format(
                            thick_min=fmt_val(ldata.get("thick_min")),
                            thick_max=fmt_val(ldata.get("thick_max")),
                            thick_avg=fmt_val(ldata.get("thick_avg")),
                            depth_min=fmt_val(ldata.get("depth_min")),
                            depth_max=fmt_val(ldata.get("depth_max")),
                            depth_avg=fmt_val(ldata.get("depth_avg")),
                            elv_min=fmt_val(ldata.get("elv_min")),
                            elv_max=fmt_val(ldata.get("elv_max")),
                            elv_avg=fmt_val(ldata.get("elv_avg")),
                        )
                    except (KeyError, IndexError):
                        pass

                # 构建完整段落
                display_name = self._build_layer_display_name(lid, name, geo_code)
                stats_text = self._format_layer_stats_text(
                    lid, layer_stats, hn_data, bh_layers,
                )

                para_text = f"{display_name}：{desc_text}"
                if desc_text and not desc_text.endswith(("。", "，")):
                    para_text += "。"
                if stats_text:
                    if not para_text.endswith("。"):
                        para_text += "。"
                    para_text += stats_text

                layer_p = self._make_normal_paragraph(para_text, intro_para)
                inserted_elements.append(layer_p)

                # 3) 试验信息段落
                table_seq += 1
                test_text = self._format_test_info_text(lid, table_seq, hn_data)
                if test_text:
                    test_p = self._make_normal_paragraph(test_text, intro_para)
                    inserted_elements.append(test_p)

                # 4) 表标题段落
                table_title = f"第{lid}层{name}工程特性指标统计表         表5-{table_seq}"
                title_p = self._make_normal_paragraph(table_title, intro_para)
                inserted_elements.append(title_p)

        # 5) 收尾段落
        closing_text = "以上各层的埋藏与分布见工程地质剖面图。"
        closing_p = self._make_normal_paragraph(closing_text, intro_para)
        inserted_elements.append(closing_p)

        # 在 anchor 之后批量插入
        for elem in reversed(inserted_elements):
            anchor.addnext(elem)

        # 如果有原有的 closing 段落 (被保留), 移到 inserted 之后
        # (已在上方处理)

        logger.info(
            f"    华宁模式: {len(age_groups)} 个年代组, "
            f"{sum(len(ids) for _, ids in age_groups)} 层, "
            f"{table_seq} 张表"
        )

    def _fill_layer_descriptions_fallback(self, layer_stats: Dict) -> None:
        """回退模式: 基于模板段落匹配填充地层描述 (原始逻辑)"""
        default_desc_templates: Dict[str, str] = {
            "1": (
                "黄褐色、灰褐色，松散、局部中密，强度不均，主要成分为风化岩渣土、"
                "碎石及建筑垃圾，块石含量约20~30%，径多10~30cm。回填时间约8年，"
                "尚未完成自重固结。场区普遍分布，"
                "厚度:{thick_min}~{thick_max}m,平均{thick_avg}m;"
                "层底标高:{elv_min}~{elv_max}m,平均{elv_avg}m;"
                "层底埋深:{depth_min}~{depth_max}m,平均{depth_avg}m。"
            ),
            "2": (
                "灰黄色、灰色，稍密~中密，饱和，主要成分为长石、石英，"
                "颗粒级配一般，含少量云母碎片。场区普遍分布，"
                "厚度:{thick_min}~{thick_max}m,平均{thick_avg}m;"
                "层底标高:{elv_min}~{elv_max}m,平均{elv_avg}m;"
                "层底埋深:{depth_min}~{depth_max}m,平均{depth_avg}m。"
            ),
        }

        generic_desc = (
            "场区{dist}分布，"
            "厚度:{thick_min}~{thick_max}m,平均{thick_avg}m;"
            "层底标高:{elv_min}~{elv_max}m,平均{elv_avg}m;"
            "层底埋深:{depth_min}~{depth_max}m,平均{depth_avg}m。"
        )

        filled: set = set()
        for p in self.doc.paragraphs:
            txt = p.text.strip()
            for lid in self.layer_ids:
                if lid in filled:
                    continue
                name = self.layer_names.get(lid, "")
                ldata = layer_stats.get(lid, {})
                if not ldata:
                    continue
                if not (
                    txt.startswith(f"{lid}层")
                    or txt.startswith(f"{lid}、")
                    or (txt and txt[0].isdigit() and name in txt[:15])
                ):
                    continue

                desc_tpl = (
                    self.config.get_layer_description(lid)
                    or default_desc_templates.get(lid)
                )

                fmt_data = {
                    "thick_min": fmt_val(ldata.get("thick_min")),
                    "thick_max": fmt_val(ldata.get("thick_max")),
                    "thick_avg": fmt_val(ldata.get("thick_avg")),
                    "depth_min": fmt_val(ldata.get("depth_min")),
                    "depth_max": fmt_val(ldata.get("depth_max")),
                    "depth_avg": fmt_val(ldata.get("depth_avg")),
                    "elv_min": fmt_val(ldata.get("elv_min")),
                    "elv_max": fmt_val(ldata.get("elv_max")),
                    "elv_avg": fmt_val(ldata.get("elv_avg")),
                }

                try:
                    if desc_tpl:
                        text = desc_tpl.format(**fmt_data)
                    else:
                        n_num = ldata.get("n", 0) or 0
                        dist = (
                            "普遍" if n_num > 50
                            else "较普遍" if n_num > 20
                            else "局部"
                        )
                        text = generic_desc.format(dist=dist, **fmt_data)
                except (KeyError, IndexError):
                    text = f"{lid}层{name}"

                set_para_text(p, text)
                filled.add(lid)
                break

        logger.info(f"    回退模式已更新: {sorted(filled)}")

    # ---- 物理力学 & 原位测试表 ----

    def _fill_phys_spt_tables(self) -> None:
        logger.info("  物理力学 & 原位测试表...")
        phys = self.data.get("phys", {})
        spt = self.data.get("spt", {})
        cpt = self.data.get("cpt", {})

        start = self.ti["phys_spt_start"]
        end = self.ti["phys_spt_end"]

        for ti in range(start, end + 1):
            if ti >= len(self.doc.tables):
                break
            t = self.doc.tables[ti]

            # 通过表头匹配地层
            hdr = "".join(
                t.rows[0].cells[c].text for c in range(min(6, len(t.rows[0].cells)))
            )
            matched = None
            for lid in self.layer_ids:
                if self.layer_names.get(lid, "") in hdr:
                    matched = lid
                    break
            if not matched:
                continue

            pd = phys.get(matched, {}).get("stats", {})
            sd = spt.get(matched, {})
            cd = cpt.get(matched, {})

            # 填充物理力学指标
            if pd:
                for ri, indicator in PHYS_ROW_MAP.items():
                    if ri >= len(t.rows):
                        continue
                    for sk, ci in STAT_COL_MAP.items():
                        v = pd.get(sk, {}).get(indicator)
                        if v is not None:
                            set_cell(t, ri, ci, fmt_val_int(v) if sk == "n" else fmt_val(v))

            # 填充标贯/动探数据
            nr = len(t.rows)
            ts = sd if sd else cd
            if ts:
                for sk, ci in STAT_COL_MAP.items():
                    raw = ts.get(sk, {}).get("raw")
                    if raw is not None and nr > 2:
                        set_cell(t, nr - 2, ci, fmt_val_int(raw))
                    corr = ts.get(sk, {}).get("corr")
                    if corr is not None and nr > 1:
                        set_cell(t, nr - 1, ci, fmt_val_int(corr))

            logger.debug(f"    T{ti}: {matched}层 {self.layer_names.get(matched, '')}")

    # ---- 承载力建议值表 ----

    def _fill_bearing_capacity(self) -> None:
        logger.info("  承载力建议值表...")
        idx = self.ti["bearing_capacity"]
        t = self.doc.tables[idx] if len(self.doc.tables) > idx else None
        if not t:
            return

        for li, lid in enumerate(self.layer_ids):
            if li + 2 >= len(t.rows):
                break
            name = self.layer_names.get(lid, "")

            # 优先从配置读取，否则留空让工程师手动填写
            bearing = self.config.get_bearing_values(lid)
            if bearing:
                fak = bearing.get("fak", "")
                es12 = bearing.get("es", "")
                es = bearing.get("e0", "")  # 注意: 列含义需匹配模板
                e0 = ""
            else:
                fak, es12, es, e0 = "", "", "", ""
                logger.debug(f"    [!] {lid}层{name} 无承载力配置")

            set_cell(t, li + 2, 0, lid)
            set_cell(t, li + 2, 1, name)
            set_cell(t, li + 2, 2, fak)
            set_cell(t, li + 2, 3, es12)
            set_cell(t, li + 2, 4, es)
            set_cell(t, li + 2, 5, e0)

    # ---- 水样 / 盐样表 ----

    def _fill_water_salt_tables(self) -> None:
        logger.info("  水样 / 盐样表...")
        water_samples = self.data.get("water_samples", [])
        salt_samples = self.data.get("salt_samples", [])

        # T19: 水样分析表
        idx_w = self.ti["water_sample"]
        t19 = self.doc.tables[idx_w] if len(self.doc.tables) > idx_w else None
        if t19 and water_samples:
            water_cols = [
                (2, "SO4"), (3, "Mg"), (4, "NH4"), (5, "OH"),
                (6, "TDS"), (7, "pH"), (8, "CO2_agg"), (9, "HCO3"), (10, "Cl"),
            ]
            for wi, w in enumerate(water_samples[:3]):
                if wi + 1 >= len(t19.rows):
                    break
                set_cell(t19, wi + 1, 0, str(w.get("field_id", "")))
                set_cell(t19, wi + 1, 1, "水")
                for ci, key in water_cols:
                    if ci < len(t19.rows[wi + 1].cells):
                        set_cell(t19, wi + 1, ci, fmt_val(w.get(key)))

        # T21: 易溶盐分析表
        idx_s = self.ti["salt_sample"]
        t21 = self.doc.tables[idx_s] if len(self.doc.tables) > idx_s else None
        if t21 and salt_samples:
            salt_cols = [(3, "SO4"), (4, "Mg"), (5, "TDS"), (6, "Cl"), (7, "pH")]
            for si, s in enumerate(salt_samples[:2]):
                if si + 1 >= len(t21.rows):
                    break
                set_cell(t21, si + 1, 0, str(s.get("location", "")))
                set_cell(t21, si + 1, 1, "土")
                for ci, key in salt_cols:
                    if ci < len(t21.rows[si + 1].cells):
                        set_cell(t21, si + 1, ci, fmt_val(s.get(key)))

    # ---- 腐蚀性评价 ----

    def _fill_corrosion_eval(self) -> None:
        logger.info("  腐蚀性评价...")
        water_samples = self.data.get("water_samples", [])
        salt_samples = self.data.get("salt_samples", [])
        corr = evaluate_corrosion(water_samples, salt_samples)
        if not corr:
            return

        corrosion_levels = ["微", "弱", "中", "强"]

        # T20: 水腐蚀性评价
        idx_wc = self.ti["water_corrosion"]
        t20 = self.doc.tables[idx_wc] if len(self.doc.tables) > idx_wc else None
        if t20 and corr.get("water"):
            wc = corr["water"]
            for wi in range(len(water_samples[:3])):
                if wi + 1 >= len(t20.rows):
                    break
                set_cell(t20, wi + 1, 2, wc["II_conc"].get("SO4", "微"))
                set_cell(t20, wi + 1, 3, wc["II_conc"].get("SO4", "微"))
                set_cell(t20, wi + 1, 4, wc["II_conc"].get("Mg", "微"))
                set_cell(t20, wi + 1, 5, wc["II_conc"].get("NH4", "微"))
                set_cell(t20, wi + 1, 7, wc["perm"])
                if len(t20.rows[wi + 1].cells) > 8:
                    set_cell(t20, wi + 1, 8, wc["steel_wet"])
                if len(t20.rows[wi + 1].cells) > 9:
                    set_cell(t20, wi + 1, 9, wc["steel_dry"])

        # T22: 土腐蚀性评价
        idx_sc = self.ti["salt_corrosion"]
        t22 = self.doc.tables[idx_sc] if len(self.doc.tables) > idx_sc else None
        if t22 and corr.get("soil"):
            sc = corr["soil"]
            for si in range(len(salt_samples[:2])):
                if si + 1 >= len(t22.rows):
                    break
                set_cell(t22, si + 1, 2, sc["II_conc"].get("SO4", "微"))
                set_cell(t22, si + 1, 3, sc["II_conc"].get("Mg", "微"))
                set_cell(t22, si + 1, 5, sc["perm"])
                if len(t22.rows[si + 1].cells) > 6:
                    set_cell(t22, si + 1, 6, sc["steel"])

        # 腐蚀性评价段落
        if corr.get("water") and corr.get("soil"):
            wc = corr["water"]
            sc = corr["soil"]
            for p in self.doc.paragraphs:
                txt = p.text.strip()
                if "腐蚀性综合评价" in txt:
                    worst_level = max(
                        wc["II_conc"].values(),
                        key=lambda x: corrosion_levels.index(x),
                    )
                    set_para_text(p, (
                        f"腐蚀性综合评价：地下水对混凝土结构具{worst_level}腐蚀性；"
                        f"地下水在长期浸水条件下对钢筋混凝土结构中的钢筋具"
                        f"{wc['steel_wet']}腐蚀性(Cl⁻)，"
                        f"在干湿交替条件下对钢筋混凝土结构中的钢筋具"
                        f"{wc['steel_dry']}腐蚀性(Cl⁻)。"
                    ))
                elif "场地土对混凝土结构具" in txt and "对钢筋混凝土" in txt:
                    set_para_text(p, (
                        f"腐蚀性评价：场地土对混凝土结构具"
                        f"{sc['II_conc'].get('SO4', '微')}腐蚀性，"
                        f"对钢筋混凝土结构中的钢筋具{sc['steel']}腐蚀性(Cl⁻)。"
                    ))

    # ---- 液化判别段落 ----

    def _fill_liquefaction(self) -> None:
        liq_data, liq_liq, liq_non = self.data.get("liquefaction", ([], 0, 0))
        logger.info(f"  液化判别 ({len(liq_data)} 点, 液化 {liq_liq})...")

        for p in self.doc.paragraphs:
            txt = p.text.strip()
            if "综合确定" in txt and "液化" in txt:
                if liq_liq > 0:
                    set_para_text(p, "综合确定场地饱和砂土层存在液化，场地液化等级为轻微。")
                else:
                    set_para_text(p, "综合确定场地饱和砂土层均不液化。")
                break
            if "进行液化判别" in txt and ("个点" in txt or "个，" in txt):
                if liq_data:
                    text = (
                        f"对饱和砂土层进行液化判别共{len(liq_data)}个点，"
                        f"其中液化{liq_liq}个点，不液化{liq_non}个点，"
                    )
                    text += "液化等级为轻微；" if liq_liq > 0 else "均不液化。"
                    set_para_text(p, text)
                break

    # ---- 基础建议表 + 桩基参数 ----

    def _fill_foundation_tables(self) -> None:
        logger.info("  基础建议表 + 桩基参数...")
        buildings = self.data.get("buildings", [])

        # T23: 基础类型选择
        idx_f = self.ti["foundation_type"]
        t23 = self.doc.tables[idx_f] if len(self.doc.tables) > idx_f else None
        if t23 and buildings:
            for bi, b in enumerate(buildings):
                if bi + 1 >= len(t23.rows):
                    break
                h = safe_float(b.get("height", 0))
                set_cell(t23, bi + 1, 0, b["name"])
                set_cell(t23, bi + 1, 1, b.get("height", ""))
                set_cell(t23, bi + 1, 2, "桩基础")

                # 持力层选择 (可通过配置自定义)
                if h and h > 20:
                    pl = "10-2强风化片麻岩"
                elif h and h > 10:
                    pl = "7层中粗砂或9层粗砂"
                else:
                    pl = "7层中粗砂"
                set_cell(t23, bi + 1, 3, pl)
                set_cell(t23, bi + 1, 4, f"{bi + 1}～{bi + 2}剖面")

        # T24: 桩基参数
        idx_p = self.ti["pile_params"]
        t24 = self.doc.tables[idx_p] if len(self.doc.tables) > idx_p else None
        if not t24:
            return

        for li, lid in enumerate(self.layer_ids):
            if li + 2 >= len(t24.rows):
                break
            name = self.layer_names.get(lid, "")
            pile = self.config.get_pile_values(lid)
            if pile:
                q1 = pile.get("qsik1", "")
                q2 = pile.get("qpk1", "")
                q3 = pile.get("qsik2", "")
                q4 = pile.get("qpk2", "")
            else:
                q1 = q2 = q3 = q4 = ""
                logger.debug(f"    [!] {lid}层{name} 无桩基参数配置")

            set_cell(t24, li + 2, 0, lid)
            set_cell(t24, li + 2, 1, name)
            set_cell(t24, li + 2, 2, q1)
            set_cell(t24, li + 2, 3, q2)
            set_cell(t24, li + 2, 4, q3)
            set_cell(t24, li + 2, 5, q4)

    # ---- 第七章: 结论与建议 ----

    def _fill_conclusion(self) -> None:
        """填充第七章结论与建议 (配置驱动 + 自动占位符注入)

        配置结构 conclusion_suggestions:
            conclusion:   (一)结论 — 段落数组, 支持占位符
            suggestions:  (二)建议 — 段落数组

        自动占位符:
            {layer_names}        地层名称列表
            {elv_min}            最低孔口高程
            {elv_max}            最高孔口高程
            {elv_range}          高程范围
            {frozen_depth}       标准冻结深度
            {liquefaction_text}  液化判别结论
            {stability_grade}    场地稳定性等级
            {suitability_grade}  适宜性等级
            {corrosion_water}    地下水腐蚀性结论
            {corrosion_soil}     场地土腐蚀性结论
        """
        cs = self.config.get_conclusion_suggestions()
        if not cs:
            return

        # --- 自动生成占位符数据 ---
        bh_info = self.data.get("borehole_info", {})
        overview = self.config.get_project_overview() or {}
        site_eval = self.config.get_site_evaluation() or {}

        # 地层名称
        layer_names = "、".join(
            self.layer_names[lid]
            for lid in self.layer_ids
            if lid in self.layer_names
        )

        # 高程统计
        elv_min = bh_info.get("elv_min")
        elv_max = bh_info.get("elv_max")
        elv_range = ""
        if elv_min is not None and elv_max is not None:
            elv_range = f"{elv_min:.2f}～{elv_max:.2f}m"

        # 液化判别
        liq_data, liq_liq, liq_non = self.data.get(
            "liquefaction", ([], 0, 0)
        )
        if liq_liq > 0:
            liq_text = "存在液化土层"
        elif liq_data:
            liq_text = "不液化"
        else:
            liq_text = "未揭露液化土层"

        # 腐蚀性评价 (从数据自动提取)
        water_samples = self.data.get("water_samples", [])
        salt_samples = self.data.get("salt_samples", [])
        corr = evaluate_corrosion(water_samples, salt_samples)
        corr_water = ""
        corr_soil = ""
        if corr.get("water"):
            wc = corr["water"]
            worst = max(
                wc["II_conc"].values(),
                key=lambda x: ["微", "弱", "中", "强"].index(x),
            )
            corr_water = f"地下水对混凝土结构具{worst}腐蚀性"
        if corr.get("salt"):
            sc = corr["salt"]
            worst_s = max(
                sc["II_conc"].values(),
                key=lambda x: ["微", "弱", "中", "强"].index(x),
            )
            corr_soil = f"场地土对混凝土结构具{worst_s}腐蚀性"

        fmt_vars: Dict[str, str] = {
            "layer_names": layer_names,
            "elv_min": fmt_val(elv_min),
            "elv_max": fmt_val(elv_max),
            "elv_range": elv_range,
            "frozen_depth": overview.get("frozen_depth", ""),
            "liquefaction_text": liq_text,
            "stability_grade": site_eval.get("stability_grade", "基本稳定"),
            "suitability_grade": site_eval.get("suitability_grade", "较适宜"),
            "corrosion_water": corr_water,
            "corrosion_soil": corr_soil,
        }

        # --- 处理 (一)结论 子节 ---
        conclusion_paras = cs.get("conclusion", [])
        if conclusion_paras:
            self._fill_conclusion_subsection(
                "(一)结论", conclusion_paras, fmt_vars,
            )

        # --- 处理 (二)建议 子节 ---
        suggestions_paras = cs.get("suggestions", [])
        if suggestions_paras:
            self._fill_conclusion_subsection(
                "(二)建议", suggestions_paras, fmt_vars,
            )

    def _fill_conclusion_subsection(
        self,
        heading: str,
        config_paras: List[str],
        fmt_vars: Dict[str, str],
    ) -> None:
        """填充结论/建议子节的段落"""
        # 定位子节标题
        start_idx: Optional[int] = None
        for i, p in enumerate(self.doc.paragraphs):
            if heading in p.text:
                start_idx = i + 1
                break
        if start_idx is None:
            return

        # 找到子节结束位置 (下一个同级标题或文档末尾)
        end_idx = len(self.doc.paragraphs)
        for j in range(start_idx, len(self.doc.paragraphs)):
            txt = self.doc.paragraphs[j].text.strip()
            if txt.startswith("(") and ")" in txt[:6] and heading not in txt:
                end_idx = j
                break
            if txt.startswith("七") or txt.startswith("八"):
                end_idx = j
                break

        # 格式化并替换
        replaced = 0
        for k, tpl in enumerate(config_paras):
            try:
                text = tpl.format(**fmt_vars)
            except KeyError:
                text = tpl
            idx = start_idx + k
            if idx < end_idx:
                set_para_text(self.doc.paragraphs[idx], text)
                replaced += 1

        # 多余模板段落清空
        for j in range(start_idx + len(config_paras), end_idx):
            if self.doc.paragraphs[j].text.strip():
                set_para_text(self.doc.paragraphs[j], "")

        if replaced:
            logger.info(f"  结论{heading}: 替换 {replaced} 段")

    # ---- 第五章: 场地条件 (地形地貌/地下水/地震/不良地质) ----

    def _fill_site_conditions(self) -> None:
        """填充第五章各子节的文字段落

        配置结构 site_conditions:
            terrain_text:       地貌描述 (触发: "地貌单元" 或 "地貌")
            topography_text:    地形地势 (触发: "地形地势" 或 "钻孔孔口高程")
            environment_text:   周边环境 (触发: "场区周边环境" 或 "建筑红线")
            surface_water_text: 地表水 (触发: "地表水" 或 "地表水体")
            seismic_params_text: 地震参数 (触发: "地震设计基本参数" 或 "设计基本地震动")
            site_class_text:    场地类别 (触发: "等效剪切波速" 或 "场地类别判定")
            seismic_stability_text: 地震稳定性 (触发: "地震稳定性")
            seismic_zone_text:  抗震地段 (触发: "抗震地段" 或 "抗震一般地段")
            soft_soil_text:     软土震陷 (触发: "软土震陷" 或 "震陷判别")
            adverse_text:       不良地质作用 (触发: "不良地质作用" 且含 "未发现" 或 "崩塌")
            buried_text:        不利埋藏物 (触发: "不利埋藏物" 或 "埋藏的河道")
        """
        sc = self.config.get_site_conditions()
        if not sc:
            return

        # 高程统计 (自动注入)
        bh_info = self.data.get("borehole_info", {})
        elv_min = fmt_val(bh_info.get("elv_min"))
        elv_max = fmt_val(bh_info.get("elv_max"))

        # 液化自动判断
        liq_data, liq_liq, liq_non = self.data.get("liquefaction", ([], 0, 0))
        if liq_liq > 0:
            auto_liq_short = "存在液化"
        elif liq_data:
            auto_liq_short = "不液化"
        else:
            auto_liq_short = "未揭露液化土层"

        # 占位符变量
        fmt_vars: Dict[str, str] = {
            "elv_min": elv_min,
            "elv_max": elv_max,
            "liquefaction_result": auto_liq_short,
        }

        # 触发词 → 配置键映射
        triggers: List[Tuple[List[str], str]] = [
            (["地貌单元", "所处地貌"], "terrain_text"),
            (["地形地势", "钻孔孔口高程"], "topography_text"),
            (["建筑红线", "场区周边环境", "周边环境"], "environment_text"),
            (["地表水体", "地表水"], "surface_water_text"),
            (["地震设计基本参数", "设计基本地震动"], "seismic_params_text"),
            (["等效剪切波速", "场地类别判定"], "site_class_text"),
            (["地震稳定性"], "seismic_stability_text"),
            (["抗震地段"], "seismic_zone_text"),
            (["软土震陷", "震陷判别"], "soft_soil_text"),
            (["不良地质作用"], "adverse_text"),
            (["不利埋藏物", "埋藏的河道"], "buried_text"),
        ]

        replaced_count = 0
        for p in self.doc.paragraphs:
            txt = p.text.strip()
            if not txt:
                continue

            for keywords, config_key in triggers:
                config_text = sc.get(config_key, "")
                if not config_text:
                    continue
                if any(kw in txt for kw in keywords):
                    try:
                        filled = config_text.format(**fmt_vars)
                    except KeyError:
                        filled = config_text
                    set_para_text(p, filled)
                    replaced_count += 1
                    break

        if replaced_count:
            logger.info(f"  场地条件: 替换 {replaced_count} 段")
        else:
            logger.debug("  场地条件: 未找到匹配段落")

    # ---- 场地稳定性及适宜性评价 (CJJ57-2012) ----

    def _fill_site_evaluation(self) -> None:
        """根据配置和实际数据生成场地稳定性及适宜性评价段落"""
        eval_config = self.config.get_site_evaluation()
        if not eval_config:
            return

        paragraphs = eval_config.get("paragraphs", [])
        if not paragraphs:
            return

        # 从数据自动生成液化文本
        liq_data, liq_liq, liq_non = self.data.get("liquefaction", ([], 0, 0))
        if liq_liq > 0:
            auto_liq = f"场地饱和砂土层存在液化（液化等级轻微）"
        elif liq_data:
            auto_liq = "场地饱和砂土层均不液化"
        else:
            auto_liq = "场地未揭露液化土层"

        # 从腐蚀性评价自动获取文本
        water_samples = self.data.get("water_samples", [])
        salt_samples = self.data.get("salt_samples", [])
        corr = evaluate_corrosion(water_samples, salt_samples)
        if corr.get("water"):
            wc = corr["water"]
            worst_level = max(
                wc["II_conc"].values(),
                key=lambda x: ["微", "弱", "中", "强"].index(x),
            )
            auto_corr = f"地下水对混凝土结构具{worst_level}腐蚀性"
        else:
            auto_corr = ""

        # 准备占位符变量: 优先使用配置中的值，其次自动生成
        placeholders = {
            "liquefaction_text": eval_config.get("liquefaction_text", auto_liq),
            "corrosion_text": eval_config.get("corrosion_text", auto_corr),
            "adverse_geology": eval_config.get("adverse_geology", "不良地质作用不发育"),
            "buried_objects": eval_config.get(
                "buried_objects",
                "未发现埋藏的河道、沟浜、墓穴、防空洞、孤石等对工程不利的埋藏物",
            ),
            "seismic_section": eval_config.get("seismic_section", "对建筑抗震一般地段"),
            "stability_grade": eval_config.get("stability_grade", "基本稳定"),
            "suitability_grade": eval_config.get("suitability_grade", "较适宜"),
            "suitability_text": eval_config.get("suitability_text", ""),
        }

        # 格式化各段落
        formatted: List[str] = []
        for tpl in paragraphs:
            try:
                text = tpl.format(**placeholders)
            except KeyError:
                text = tpl  # 未定义的占位符保留原文
            formatted.append(text)

        # 在模板中定位并替换: 找到包含触发关键词的首段，替换后续段
        trigger_keywords = [
            "场地稳定性", "稳定性评价", "适宜性评价",
            "场地稳定性及适宜性", "稳定性和适宜性",
        ]

        replaced = False
        for i, p in enumerate(self.doc.paragraphs):
            txt = p.text.strip()
            if any(kw in txt for kw in trigger_keywords):
                # 首段替换为格式化后的第一段
                set_para_text(p, formatted[0])
                # 后续配置段落: 尝试替换模板中的后续段落
                for j, ftext in enumerate(formatted[1:], start=1):
                    target_idx = i + j
                    if target_idx < len(self.doc.paragraphs):
                        next_txt = self.doc.paragraphs[target_idx].text.strip()
                        # 如果下一段是相关段落（非新章节），替换它
                        if next_txt and not any(
                            kw in next_txt
                            for kw in ("结论", "建议", "技术标准", "勘察依据", "地基基础")
                        ):
                            set_para_text(self.doc.paragraphs[target_idx], ftext)
                        else:
                            # 模板段落不够，剩余文本追加到上一段末尾
                            prev_p = self.doc.paragraphs[target_idx - 1]
                            old_text = prev_p.text.rstrip()
                            set_para_text(prev_p, old_text + "\n" + ftext)
                    else:
                        # 模板段落不足，追加到最后一段
                        last_p = self.doc.paragraphs[len(self.doc.paragraphs) - 1]
                        old_text = last_p.text.rstrip()
                        set_para_text(last_p, old_text + "\n" + ftext)
                replaced = True
                break

        if replaced:
            logger.info(
                f"  场地评价: {placeholders['stability_grade']} / "
                f"{placeholders['suitability_grade']}"
            )
        else:
            logger.debug("  场地评价: 未找到匹配段落，跳过")

    # ---- 地基评价 (4.5.7 §2 均匀性 / §5 软弱下卧层 / §6 变形参数) ----

    def _fill_foundation_evaluation(self) -> None:
        """根据配置和数据生成地基评价段落"""
        eval_config = self.config.get_foundation_evaluation()
        if not eval_config:
            return

        paragraphs = eval_config.get("paragraphs", [])
        if not paragraphs:
            return

        # --- 自动数据收集 ---
        layer_stats = self.data.get("layers", {})
        bearing_config = {
            lid: self.config.get_bearing_values(lid)
            for lid in self.layer_ids
        }

        # §2 地基均匀性: 根据各层厚度变异系数判断
        auto_uniformity = self._auto_uniformity_text(layer_stats)

        # §5 软弱下卧层: 找出低承载力层
        auto_weak_layer = self._auto_weak_layer_text(bearing_config)

        # §6 变形计算参数: 引用 T18 表 Es 值
        bearing_idx = self.ti.get("bearing_capacity", 18)
        table_ref = f"表{bearing_idx}" if bearing_idx else "承载力建议值表"
        auto_deformation = (
            f"各土层压缩模量（Es）详见{table_ref}，"
            f"建议按《建筑地基基础设计规范》(GB 50007-2011)进行地基变形计算。"
        )

        # 准备占位符: 优先使用配置值，否则使用自动生成值
        placeholders: Dict[str, str] = {
            "uniformity_text": eval_config.get("uniformity_text", auto_uniformity),
            "weak_layer_text": eval_config.get("weak_layer_text", auto_weak_layer),
            "deformation_text": eval_config.get("deformation_text", auto_deformation),
        }

        # 附加占位符: 持力层信息
        bearing_layer = eval_config.get("bearing_layer", "")
        bearing_fak = eval_config.get("bearing_fak", "")
        placeholders["bearing_layer"] = bearing_layer
        placeholders["bearing_fak"] = bearing_fak

        # 格式化段落
        formatted: List[str] = []
        for tpl in paragraphs:
            try:
                text = tpl.format(**placeholders)
            except KeyError:
                text = tpl
            formatted.append(text)

        # 在模板中定位触发关键词段落并替换
        trigger_keywords = [
            "地基均匀性", "天然地基评价", "软弱下卧层",
            "地基评价", "地基土评价", "地基均匀",
        ]

        replaced = False
        for i, p in enumerate(self.doc.paragraphs):
            txt = p.text.strip()
            if any(kw in txt for kw in trigger_keywords):
                # 首段替换
                set_para_text(p, formatted[0])
                # 后续段落
                for j, ftext in enumerate(formatted[1:], start=1):
                    target_idx = i + j
                    if target_idx < len(self.doc.paragraphs):
                        next_txt = self.doc.paragraphs[target_idx].text.strip()
                        # 如果下一段是相关段落（非新章节标题），替换
                        if next_txt and not any(
                            kw in next_txt for kw in (
                                "结论", "建议", "技术标准", "勘察依据",
                                "场地稳定性", "地下水", "基坑",
                            )
                        ):
                            set_para_text(
                                self.doc.paragraphs[target_idx], ftext
                            )
                        else:
                            # 模板段落不够，追加到当前段
                            prev_p = self.doc.paragraphs[target_idx - 1]
                            old = prev_p.text.rstrip()
                            set_para_text(prev_p, old + "\n" + ftext)
                    else:
                        # 模板段落不足，追加到最后一段
                        last_p = self.doc.paragraphs[len(self.doc.paragraphs) - 1]
                        old = last_p.text.rstrip()
                        set_para_text(last_p, old + "\n" + ftext)
                replaced = True
                break

        if replaced:
            logger.info("  地基评价: 已填充 %d 段", len(formatted))
        else:
            logger.debug("  地基评价: 未找到匹配段落，跳过")

    def _auto_uniformity_text(
        self, layer_stats: Dict[str, Dict[str, Any]]
    ) -> str:
        """根据各层厚度变异系数自动判断地基均匀性"""
        if not layer_stats or not self.layer_ids:
            return "地基土均匀性需根据实际勘探资料判定。"

        # 计算各层厚度变异系数 (CV of thickness)
        cvs: List[float] = []
        for lid in self.layer_ids:
            ls = layer_stats.get(lid, {})
            thk_avg = ls.get("thick_avg")
            thk_min = ls.get("thick_min")
            thk_max = ls.get("thick_max")
            if thk_avg and thk_min is not None and thk_max is not None and thk_avg > 0:
                # 近似变异系数 = (max - min) / avg
                cv = (thk_max - thk_min) / thk_avg
                cvs.append(cv)

        if not cvs:
            return "地基土均匀性需根据实际勘探资料判定。"

        avg_cv = sum(cvs) / len(cvs)
        max_cv = max(cvs)

        if max_cv < 0.3:
            return (
                "拟建场地地基土为均匀地基，各土层分布较稳定，"
                "厚度变化较小，地基土均匀性较好。"
            )
        elif max_cv < 0.5:
            return (
                "拟建场地地基土为较均匀地基，各土层分布有一定变化，"
                "厚度变化一般，地基土均匀性一般。"
            )
        else:
            return (
                "拟建场地地基土为不均匀地基，各土层分布变化较大，"
                "厚度变化较大，地基土均匀性较差。"
            )

    def _auto_weak_layer_text(
        self, bearing_config: Dict[str, Optional[Dict[str, str]]]
    ) -> str:
        """自动检测软弱下卧层（承载力 < 100 kPa 的土层）"""
        weak_layers: List[Tuple[str, str, str]] = []  # (层号, 名称, fak)

        for lid in self.layer_ids:
            bc = bearing_config.get(lid)
            if not bc:
                continue
            fak_str = bc.get("fak", "")
            fak = safe_float(fak_str)
            if fak is not None and fak < 100 and fak > 0:
                name = self.layer_names.get(lid, "")
                weak_layers.append((lid, name, str(int(fak))))

        if weak_layers:
            desc_parts = []
            for lid, name, fak in weak_layers:
                desc_parts.append(f"{lid}层{name}（fak={fak}kPa）")
            layers_desc = "、".join(desc_parts)
            return (
                f"场地内局部分布{layers_desc}，承载力较低，"
                f"作为软弱下卧层需进行验算，"
                f"建议按《建筑地基基础设计规范》(GB 50007-2011)第5.2.7条进行软弱下卧层验算。"
            )
        else:
            return (
                "场地内未发现明显的软弱下卧层，"
                "各土层承载力可满足一般建筑天然地基要求。"
            )

    # ---- 第六章: 岩土工程分析评价 (配置驱动段落填充) ----

    def _fill_analysis_evaluation(self) -> None:
        """填充第六章岩土工程分析评价各子节

        配置结构 analysis_evaluation:
            layer_eval:       (一)岩土层逐层工程评价 — key为层号(如"1","1-1")
            anti_float:       (四)3 地下水的力学作用/抗浮
            foundation_text:  (五)1 基础选型补充说明
            pile_eval:        (五)2 桩基评价综合文字
            special_soils:    (六)特殊性岩土工程分析评价
            stability:        (八)地基稳定性评价
            excavation:       (九)基坑开挖有关问题 (has_basement=False时清空)
            risk:             (十)地质条件可能造成的工程风险 (has_basement=False时清空)
            deformation:      (十一)建筑物变形分析
        """
        ae = self.config.get_analysis_evaluation()
        if not ae:
            return

        # 有地下室判断 (基坑/风险章节条件控制)
        overview = self.config.get_project_overview()
        conditions = self._evaluate_standard_conditions()
        has_basement = (overview or {}).get(
            "has_basement", conditions.get("has_basement", False)
        )

        # 章节映射: (触发关键词列表, 配置键, 条件类型)
        #   cond=None: 始终处理
        #   cond="basement": 仅 has_basement=True 时处理, 否则清空
        sections: List[Tuple[List[str], str, Any]] = [
            (
                ["岩土层(体)工程分析", "岩土层(体)工程评价"],
                "layer_eval", None,
            ),
            (
                ["地下水的力学作用"],
                "anti_float", None,
            ),
            (
                ["地基及基础方案", "基础选型"],
                "foundation_text", None,
            ),
            (
                ["桩基评价", "桩基形式"],
                "pile_eval", None,
            ),
            (
                ["特殊性岩土"],
                "special_soils", None,
            ),
            (
                ["地基稳定性评价"],
                "stability", None,
            ),
            (
                ["基坑开挖有关问题", "基坑开挖"],
                "excavation", "basement",
            ),
            (
                ["工程风险"],
                "risk", "basement",
            ),
            (
                ["建筑物变形分析"],
                "deformation", None,
            ),
        ]

        current_section: Optional[str] = None
        current_config_key: Optional[str] = None
        section_start: int = 0
        replaced_count = 0

        for i, p in enumerate(self.doc.paragraphs):
            txt = p.text.strip()

            # --- 检测章节标题 (合并专用 + 通用检测) ---
            matched_section_key: Optional[str] = None

            # 1) 专用关键词匹配 (sections 列表)
            for keywords, config_key, cond in sections:
                if any(kw in txt for kw in keywords):
                    matched_section_key = config_key
                    break

            # 2) 通用标题模式: 括号格式 (X) 或已知非括号标题
            if not matched_section_key:
                if (
                    (txt.startswith("(") and ")" in txt[:6])
                    or txt in ("基坑开挖有关问题",)
                ):
                    matched_section_key = "__boundary__"

            # 处理检测结果
            if matched_section_key:
                # 先 flush 上一个章节
                if current_section:
                    replaced_count += self._process_analysis_section(
                        current_section, current_config_key,
                        section_start, i, ae, has_basement,
                    )

                if matched_section_key == "__boundary__":
                    # 不关心的章节, 重置状态
                    current_section = None
                    current_config_key = None
                else:
                    current_section = txt
                    current_config_key = matched_section_key
                    section_start = i + 1

        # 处理最后一个章节
        if current_section:
            replaced_count += self._process_analysis_section(
                current_section, current_config_key,
                section_start, len(self.doc.paragraphs), ae, has_basement,
            )

        if replaced_count:
            logger.info(f"  分析评价: 处理 {replaced_count} 个子节")
        else:
            logger.debug("  分析评价: 未找到匹配章节")

    def _process_analysis_section(
        self,
        heading: str,
        config_key: Optional[str],
        start: int,
        end: int,
        ae: Dict[str, Any],
        has_basement: bool,
    ) -> int:
        """处理第六章的一个子节

        返回处理的段落数。
        """
        if not config_key:
            return 0

        config_data = ae.get(config_key, None)

        # === 条件清空: 无基坑时清空基坑/风险章节 ===
        # 匹配条件: 标题中包含 "基坑" 或 "工程风险"
        if not has_basement and (
            "基坑" in heading or "工程风险" in heading
        ):
            cleared = 0
            for j in range(start, end):
                if self.doc.paragraphs[j].text.strip():
                    set_para_text(self.doc.paragraphs[j], "")
                    cleared += 1
            if cleared:
                logger.info(
                    f"    {config_key}: 无基坑, 清空 {cleared} 段"
                )
            return cleared

        if not config_data:
            return 0

        # === (一) 岩土层逐层评价: 按层号匹配 ===
        if config_key == "layer_eval" and isinstance(config_data, dict):
            replaced = 0
            for j in range(start, end):
                txt_j = self.doc.paragraphs[j].text.strip()
                if not txt_j:
                    continue
                for layer_key, layer_text in config_data.items():
                    if layer_text and txt_j.startswith(f"第{layer_key}层"):
                        set_para_text(self.doc.paragraphs[j], layer_text)
                        replaced += 1
                        break
            return replaced

        # === 通用段落替换 ===
        if not isinstance(config_data, list):
            config_data = [str(config_data)]

        replaced = 0
        for k, text in enumerate(config_data):
            idx = start + k
            if idx < end:
                set_para_text(self.doc.paragraphs[idx], text)
                replaced += 1
        # 多余模板段落清空
        for j in range(start + len(config_data), end):
            if self.doc.paragraphs[j].text.strip():
                set_para_text(self.doc.paragraphs[j], "")

        return replaced

    # ---- 技术标准列表 (配置驱动 + 条件过滤) ----

    def _fill_standards(self) -> None:
        """根据配置和实际数据生成技术标准列表段落"""
        standards_config = self.config.get_technical_standards()
        if not standards_config:
            return

        # 始终包含的标准
        always = standards_config.get("always", [])

        # 按条件过滤的标准
        conditional = standards_config.get("conditional", [])
        conditions = self._evaluate_standard_conditions()

        # 合并: 始终项 + 满足条件的条件项
        all_standards = list(always)
        for s in conditional:
            cond = s.get("condition", "always")
            if cond == "always" or conditions.get(cond, False):
                all_standards.append(s)

        # 法律法规和其他依据
        laws = standards_config.get("laws", [])
        other = standards_config.get("other", [])

        # 组装格式化文本
        lines: List[str] = []
        idx = 1

        if all_standards:
            lines.append("1、国家标准：")
            for s in all_standards:
                lines.append(f"{idx})《{s['name']}》({s['code']})")
                idx += 1

        if laws:
            lines.append("2、法律法规：")
            for law in laws:
                lines.append(f"{idx})《{law}》" if "《" not in law else f"{idx}){law}")
                idx += 1

        if other:
            lines.append("3、其他：")
            for item in other:
                lines.append(f"{idx}){item}")
                idx += 1

        full_text = "\n".join(lines)

        # 在模板中定位技术标准段落并替换
        trigger_keywords = ["执行的主要技术标准", "技术标准", "勘察依据", "依据的技术标准"]
        replaced = False

        for i, p in enumerate(self.doc.paragraphs):
            txt = p.text.strip()
            if any(kw in txt for kw in trigger_keywords):
                set_para_text(p, full_text)
                replaced = True
                # 清除后续已有的标准条目段落
                for j in range(i + 1, min(i + 40, len(self.doc.paragraphs))):
                    next_txt = self.doc.paragraphs[j].text.strip()
                    if not next_txt:
                        continue
                    # 匹配标准条目格式: 数字)《...》(编号) 或 纯标准名称
                    if (
                        re.match(r"^\d+[)）、]", next_txt)
                        or "《" in next_txt and "(" in next_txt
                        or next_txt.startswith(("国家标准", "行业标准", "地方标准", "法律法规", "其他"))
                    ):
                        set_para_text(self.doc.paragraphs[j], "")
                    elif any(kw in next_txt for kw in ("工程概况", "勘察目的", "场地", "拟建")):
                        break  # 已到达下一个章节
                break

        if replaced:
            included = [s["name"] for s in all_standards]
            excluded = [
                s["name"] for s in conditional
                if s.get("condition", "always") != "always"
                and not conditions.get(s["condition"], False)
            ]
            logger.info(f"  技术标准: 纳入 {len(all_standards)} 条, 排除 {len(excluded)} 条")
            if excluded:
                logger.debug(f"    排除: {', '.join(excluded)}")
        else:
            logger.debug("  技术标准: 未找到匹配段落，跳过")

    def _evaluate_standard_conditions(self) -> Dict[str, bool]:
        """评估标准条件，返回各条件是否满足"""
        layer_ids = self.layer_ids
        bh_info = self.data.get("borehole_info", {})
        buildings = self.data.get("buildings", [])
        water_samples = self.data.get("water_samples", [])

        # 有岩层: 层号含 "10-" 等
        has_rock_layers = any("-" in str(lid) for lid in layer_ids)

        # 有岩石试验样品
        has_rock_samples = (bh_info.get("rock", 0) or 0) > 0

        # 有水样
        has_water_samples = len(water_samples) > 0

        # 配置了桩基参数
        has_pile = any(
            self.config.get_pile_values(lid) is not None
            for lid in layer_ids
        )

        # 有地下室/地下结构
        has_basement = False
        for b in buildings:
            floors_str = b.get("floors", "")
            height_str = b.get("height", "")
            # 检查是否有地下层数标注 (如 "地上18层/地下1层")
            if "地下" in floors_str:
                has_basement = True
                break
            # 检查备注或结构字段
            for field in ["size", "span", "indoor_elv", "structure"]:
                if "地下" in str(b.get(field, "")):
                    has_basement = True
                    break
            if has_basement:
                break

        return {
            "has_rock_layers": has_rock_layers,
            "has_rock_samples": has_rock_samples,
            "has_water_samples": has_water_samples,
            "has_pile_foundation": has_pile,
            "has_basement": has_basement,
        }

    # ---- 日期替换 (配置驱动) ----

    def _apply_date_replacements(self) -> None:
        """按配置执行日期替换 (不再硬编码)"""
        date_repls = self.config.get_date_replacements()
        if not date_repls:
            return

        logger.info("  日期替换...")
        for old, new in date_repls:
            for p in self.doc.paragraphs:
                replace_in_para(p, old, new)


# ============================================================
# 主程序
# ============================================================

def load_layer_names(
    config: ProjectConfig,
    loader: SurveyDataLoader,
    layer_stats: Dict[str, Any],
    args_layers: Optional[str],
) -> Tuple[Dict[str, str], List[str]]:
    """加载地层名称映射和排序后的地层 ID 列表"""

    layer_names: Dict[str, str] = {}

    # 优先使用 --layers 参数
    if args_layers and os.path.isfile(args_layers):
        with open(args_layers, "r", encoding="utf-8") as f:
            layer_names = json.load(f)
    # 其次使用配置中的地层定义
    elif config.layers:
        for layer in config.layers:
            lid = str(layer.get("id", ""))
            name = layer.get("name", "")
            if lid and name:
                layer_names[lid] = name
    # 最后从物理力学表推断
    else:
        phys = loader.load_physical_stats()
        for lid in layer_stats:
            if not (any(c.isdigit() for c in str(lid)) and len(str(lid)) < 10):
                continue
            if lid in phys and phys[lid].get("name"):
                layer_names[lid] = phys[lid]["name"]
            else:
                layer_names[lid] = f"第{lid}层"

    # 排序层号
    valid_ids = [
        k
        for k in layer_stats.keys()
        if any(c.isdigit() for c in str(k)) and len(str(k)) < 10
    ]
    layer_ids = sorted(valid_ids, key=layer_sort_key)

    return layer_names, layer_ids


def main() -> None:
    parser = argparse.ArgumentParser(
        description=f"岩土工程勘察报告自动生成工具 v{__version__}"
    )
    parser.add_argument("--config", "-c", help="JSON 配置文件路径 (推荐)")
    parser.add_argument("--project", "-p", help="项目目录路径 (向后兼容)")
    parser.add_argument("--template", "-t", help="模板文件路径")
    parser.add_argument("--output", "-o", help="输出文件路径")
    parser.add_argument("--layers", help="地层名称映射 JSON 文件")
    parser.add_argument("--dry-run", action="store_true", help="仅加载数据，不生成报告")
    parser.add_argument("--verbose", "-v", action="store_true", help="输出详细日志")
    args = parser.parse_args()

    setup_logging(args.verbose)

    if not args.config and not args.project:
        parser.error("必须指定 --config 或 --project")

    # 检查第三方依赖 (dry-run 也需要 xlrd/openpyxl)
    _check_dependencies()

    # 1. 加载配置
    logger.info("=" * 60)
    logger.info(f"  岩土工程勘察报告自动生成工具 v{__version__}")
    logger.info("=" * 60)

    config = ProjectConfig(
        config_path=args.config,
        project_dir=args.project,
        template_override=args.template,
        output_override=args.output,
        layers_override=args.layers,
    )

    # 2. 加载数据
    loader = SurveyDataLoader(config.excel_dir)
    logger.info(f"\n[数据加载] Excel 目录: {config.excel_dir}")

    boreholes = loader.load_boreholes()
    bh_info = loader.classify_boreholes(boreholes)
    logger.info(
        f"  钻孔: {bh_info['total']} 个, "
        f"进尺 {bh_info['total_depth']:.2f}m"
    )
    logger.info(
        f"  控制孔 {bh_info['ctrl']}, 一般孔 {bh_info['general']}"
    )
    logger.info(
        f"  取土 {bh_info['qutu']}, 标贯 {bh_info['biaoguan']}, "
        f"一般 {bh_info['yiban']}, 动探 {bh_info['zhongtan']}, "
        f"波速 {bh_info['bosk']}"
    )

    layer_stats = loader.load_layer_stats()
    phys = loader.load_physical_stats()
    spt = loader.load_spt_stats()
    cpt = loader.load_cpt_stats()
    liq_data, liq_liq, liq_non = loader.load_liquefaction()
    rock = loader.load_rock_stats()
    buildings = loader.load_buildings()
    water_samples = loader.load_water_samples()
    salt_samples = loader.load_salt_samples()

    logger.info(
        f"  物理力学: {len(phys)} 层, "
        f"标贯: {len(spt)} 层, "
        f"动探: {len(cpt)} 层"
    )
    logger.info(
        f"  液化: {len(liq_data)} 点, "
        f"水样: {len(water_samples)} 件, "
        f"盐样: {len(salt_samples)} 件"
    )

    # 地层名称
    layer_names, layer_ids = load_layer_names(config, loader, layer_stats, args.layers)
    logger.info(f"  地层 ID 序列: {layer_ids}")

    # 华宁数据库 (可选)
    hn_data: Dict[str, Any] = {"available": False}
    hn_db_dir = config.get_hn_db_dir()
    if hn_db_dir and os.path.isdir(hn_db_dir):
        logger.info(f"\n[华宁数据库] 目录: {hn_db_dir}")
        project_code = config.get_hn_project_code()
        reader = HuaNingDBReader(hn_db_dir, project_code)
        hn_data = reader.read()
    elif hn_db_dir:
        logger.warning(f"  华宁数据库目录不存在: {hn_db_dir}")

    if args.dry_run:
        logger.info("\n[Dry-run] 数据加载完成，不生成报告。")
        return

    # 3. 填充报告
    logger.info(f"\n[模板] {config.template_path}")
    logger.info(f"[输出] {config.output_path}")

    data = {
        "boreholes": boreholes,
        "borehole_info": bh_info,
        "layers": layer_stats,
        "phys": phys,
        "spt": spt,
        "cpt": cpt,
        "liquefaction": (liq_data, liq_liq, liq_non),
        "rock": rock,
        "buildings": buildings,
        "water_samples": water_samples,
        "salt_samples": salt_samples,
        "hn_data": hn_data,
    }

    filler = ReportFiller(
        config.template_path,
        config.output_path,
        data,
        layer_names,
        layer_ids,
        config,
    )
    filler.fill_all()
    filler.save()

    logger.info(f"\n{'=' * 60}")
    logger.info(f"  生成完成: {config.output_path}")
    logger.info(f"{'=' * 60}")


if __name__ == "__main__":
    main()
