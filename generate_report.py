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
        self._fill_buildings_table()
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

    # ---- 地层描述段落 ----

    def _fill_layer_descriptions(self) -> None:
        logger.info("  地层描述段落...")
        layer_stats = self.data.get("layers", {})

        # 默认地层描述模板 (威海滨海区典型)
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

                # 优先使用配置文件中的描述，其次默认模板，最后通用模板
                desc_tpl = (
                    self.config.get_layer_description(lid)
                    or default_desc_templates.get(lid)
                )

                # 准备数据
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
                        n_val = fmt_val_int(ldata.get("n", ""))
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

        logger.info(f"    已更新: {sorted(filled)}")

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

    # ---- 结论段落 ----

    def _fill_conclusion(self) -> None:
        logger.info("  结论段落...")
        bh_info = self.data["borehole_info"]

        for p in self.doc.paragraphs:
            txt = p.text.strip()
            if "场地土主要由" in txt:
                names = "、".join(
                    self.layer_names[lid]
                    for lid in self.layer_ids
                    if lid in self.layer_names
                )
                if names:
                    set_para_text(p, f"场地土主要由{names}等组成。")
            if "钻孔孔口高程" in txt and "elv_min" in bh_info:
                set_para_text(p, (
                    f"所处地貌为山前海积、冲洪积小平原交界地带，"
                    f"场地经整平后地形较平缓，"
                    f"钻孔孔口高程{bh_info['elv_min']:.2f}～{bh_info['elv_max']:.2f}m"
                    f"(根据钻孔统计)。"
                ))

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
