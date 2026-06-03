#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
腐蚀性评价模块 — GB50021-2001 第12章
============================================================
独立模块，可单独使用或集成到报告生成工具中。

用法:
    from corrosion_eval import evaluate_corrosion
    result = evaluate_corrosion(water_samples, salt_samples)
    print(result['water']['II_conc'])  # {'SO4':'微','Mg':'微',...}
    print(result['soil']['steel'])     # '微'
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional


def _safe_get(sample: Dict[str, Any], key: str, default: float = 0) -> float:
    """安全获取样本数值，None 或非数值返回 default"""
    v = sample.get(key)
    if v is None:
        return default
    try:
        return float(v)
    except (ValueError, TypeError):
        return default


def _grade_by_limits(value: float, limits: tuple) -> str:
    """按阈值分级: (弱阈值, 中阈值, 强阈值) → 微/弱/中/强"""
    if value < limits[0]:
        return "微"
    if value < limits[1]:
        return "弱"
    if value < limits[2]:
        return "中"
    return "强"


def evaluate_corrosion(
    water_samples: List[Dict[str, Any]],
    salt_samples: List[Dict[str, Any]],
) -> Dict[str, Any]:
    """
    按 GB50021-2001 第12章评价水和土的腐蚀性。

    参数:
        water_samples: 水样列表，每个水样包含:
            'SO4', 'Mg', 'NH4', 'OH', 'TDS', 'Cl', 'pH', 'CO2_agg', 'field_id'(可选)
        salt_samples: 土样列表，每个土样包含:
            'SO4', 'Mg', 'Cl', 'pH', 'location'(可选)

    返回:
        dict，包含 'water' 和/或 'soil' 键，各含:
            - 'II_conc': dict — II类环境各指标腐蚀等级
            - 'perm': str — 地层渗透性等级
            - 'steel_wet' / 'steel_dry': str — 钢筋腐蚀等级 (仅水)
            - 'steel': str — 钢筋腐蚀等级 (仅土)
            - 各离子浓度最值
    """
    result: Dict[str, Any] = {}

    # ── 水腐蚀性评价 ──
    if water_samples:
        # 取所有水样中的最大浓度 (pH 取最小值)
        def worst_max(key: str) -> float:
            return max(_safe_get(w, key) for w in water_samples)

        def worst_min(key: str, default: float = 7.0) -> float:
            return min((_safe_get(w, key, default) for w in water_samples), default=default)

        so4 = worst_max("SO4")
        mg = worst_max("Mg")
        nh4 = worst_max("NH4")
        oh = worst_max("OH")
        tds = worst_max("TDS")
        cl = worst_max("Cl")
        ph = worst_min("pH")
        co2 = worst_max("CO2_agg")

        # 表 12.2.1: II类环境 水对混凝土结构腐蚀性等级
        # (弱, 中, 强) 阈值 — 低于弱阈值即为"微"
        limits_II = {
            "SO4": (300, 1500, 3000),
            "Mg": (2000, 3000, 4000),
            "NH4": (500, 800, 1000),
            "OH": (5000, 6000, 7000),
            "TDS": (20000, 50000, 60000),
        }
        II_conc = {
            k: _grade_by_limits(v, limits_II[k])
            for k, v in [("SO4", so4), ("Mg", mg), ("NH4", nh4), ("OH", oh), ("TDS", tds)]
        }

        # 表 12.2.2: 按地层渗透性评价
        if ph > 6.5 and co2 < 15:
            perm = "微"
        elif ph > 5.0:
            perm = "弱"
        else:
            perm = "中"

        # 表 12.2.3: 对钢筋混凝土中钢筋的腐蚀性 (Cl⁻)
        if cl < 100:
            steel_dry = "微"
        elif cl < 500:
            steel_dry = "弱"
        elif cl < 5000:
            steel_dry = "中"
        else:
            steel_dry = "强"

        steel_wet = "微" if cl < 10000 else "弱"

        result["water"] = {
            "II_conc": II_conc,
            "perm": perm,
            "steel_wet": steel_wet,
            "steel_dry": steel_dry,
            "SO4": so4, "Mg": mg, "NH4": nh4, "OH": oh,
            "TDS": tds, "Cl": cl, "pH": ph, "CO2": co2,
        }

    # ── 土腐蚀性评价 ──
    if salt_samples:
        def worst_max_s(key: str) -> float:
            return max(_safe_get(s, key) for s in salt_samples)

        def worst_min_s(key: str, default: float = 7.0) -> float:
            return min((_safe_get(s, key, default) for s in salt_samples), default=default)

        so4_s = worst_max_s("SO4")
        mg_s = worst_max_s("Mg")
        cl_s = worst_max_s("Cl")
        ph_s = worst_min_s("pH")

        # 表 12.2.4: II类环境 土对混凝土
        # SO4²⁻ 阈值 (mg/kg)
        if so4_s < 450:
            grade_so4 = "微"
        elif so4_s < 2250:
            grade_so4 = "弱"
        elif so4_s < 4500:
            grade_so4 = "中"
        else:
            grade_so4 = "强"

        # Mg²⁺ 阈值 (mg/kg)
        if mg_s < 3000:
            grade_mg = "微"
        elif mg_s < 4500:
            grade_mg = "弱"
        elif mg_s < 6000:
            grade_mg = "中"
        else:
            grade_mg = "强"

        II_conc_s = {"SO4": grade_so4, "Mg": grade_mg}

        # 表 12.2.5: 按地层渗透性
        perm_s = "微" if ph_s > 5.0 else "弱"

        # 对钢筋 (Cl⁻ mg/kg)
        if cl_s < 250:
            steel_s = "微"
        elif cl_s < 500:
            steel_s = "弱"
        else:
            steel_s = "中"

        result["soil"] = {
            "II_conc": II_conc_s,
            "perm": perm_s,
            "steel": steel_s,
            "SO4": so4_s, "Mg": mg_s, "Cl": cl_s, "pH": ph_s,
        }

    return result


def format_corrosion_report(corr: Dict[str, Any]) -> str:
    """将评价结果格式化为可读文本报告"""
    lines: List[str] = []

    if corr.get("water"):
        w = corr["water"]
        lines.append("【地下水腐蚀性评价 (GB50021-2001 第12章)】")
        lines.append(
            f"  SO4²⁻={w['SO4']:.0f}mg/L  Mg²⁺={w['Mg']:.0f}mg/L  "
            f"NH4⁺={w['NH4']:.1f}mg/L  OH⁻={w['OH']:.0f}mg/L  "
            f"总矿化度={w['TDS']:.0f}mg/L"
        )
        lines.append(
            f"  pH={w['pH']:.2f}  侵蚀性CO₂={w['CO2']:.0f}mg/L  "
            f"Cl⁻={w['Cl']:.1f}mg/L"
        )
        lines.append(f"  II类环境: 混凝土{w['II_conc']}")
        lines.append(f"  地层渗透性: {w['perm']}")
        lines.append(f"  对钢筋: 长期浸水{w['steel_wet']}, 干湿交替{w['steel_dry']}")

    if corr.get("soil"):
        s = corr["soil"]
        lines.append("【场地土腐蚀性评价 (GB50021-2001 第12章)】")
        lines.append(
            f"  SO4²⁻={s['SO4']:.0f}mg/kg  Mg²⁺={s['Mg']:.1f}mg/kg  "
            f"Cl⁻={s['Cl']:.1f}mg/kg  pH={s['pH']:.2f}"
        )
        lines.append(f"  II类环境: 混凝土{s['II_conc']}")
        lines.append(f"  地层渗透性: {s['perm']}")
        lines.append(f"  对钢筋: {s['steel']}")

    return "\n".join(lines)


# 快速测试
if __name__ == "__main__":
    import sys
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding="utf-8")

    w1 = {
        "SO4": 112, "Mg": 53, "NH4": 0, "OH": 0,
        "TDS": 2646, "Cl": 1309, "pH": 7.66, "CO2_agg": 0,
    }
    w2 = {
        "SO4": 212, "Mg": 246, "NH4": 0.43, "OH": 0,
        "TDS": 5851, "Cl": 3043, "pH": 7.23, "CO2_agg": 0,
    }
    s1 = {"SO4": 80, "Mg": 3.4, "Cl": 39, "pH": 7.2}

    corr = evaluate_corrosion([w1, w2], [s1])
    print(format_corrosion_report(corr))
