"""
剪切波速估算与场地类别判定模块
GB 50011-2010 §4.1.3~§4.1.6

方案一: 无实测波速数据时，按土层名称和状态估算 Vs
方案二: 有波速测试报告时，从数据提取 → 见 wave_velocity_parser.py

估算依据:
  - 《建筑抗震设计规范》GB 50011-2010 表4.1.3 (土的类型划分)
  - 威海地区勘察经验值
"""

from __future__ import annotations

import math
from typing import Any, Dict, List, Optional, Tuple

# ============================================================
# 土层名称 → 经验 Vs 值 (m/s)
# ============================================================

STATE_MODIFIERS: Dict[str, float] = {
    "流塑": 0.65, "软塑": 0.80, "可塑": 1.00, "硬塑": 1.25, "坚硬": 1.50,
    "松散": 0.70, "稍密": 0.85, "中密": 1.00, "密实": 1.20, "很密": 1.35,
    "全风化": 0.50, "强风化": 0.75, "中风化": 1.00, "微风化": 1.30, "未风化": 1.50,
}

# (基准Vs, 是否为砂土/碎石)
LAYER_VS_BASE: Dict[str, Tuple[float, bool]] = {
    "杂填土": (120, False), "素填土": (130, False), "填土": (125, False),
    "耕土": (110, False), "冲填土": (120, False),
    "淤泥": (80, False), "淤泥质粉质黏土": (95, False),
    "淤泥质黏土": (90, False), "淤泥质土": (92, False),
    "泥炭": (70, False), "泥炭质土": (75, False),
    "黏土": (220, False), "粉质黏土": (200, False), "粉土": (180, False),
    "粉砂": (160, True), "细砂": (200, True), "中砂": (280, True),
    "粗砂": (340, True), "砾砂": (400, True),
    "角砾": (420, True), "圆砾": (430, True), "卵石": (450, True),
    "碎石": (400, True), "块石": (420, True), "漂石": (430, True),
    "碎石土": (380, False), "砂质黏性土": (220, False), "砾质黏性土": (240, False),
    "残积土": (280, False), "残积黏性土": (260, False), "残积砂质黏性土": (300, False),
    # 花岗岩类 (精确条目优先 — 含"花岗"不再触发通用"强风化"modifier)
    "微风化花岗岩": (1200, False), "中风化花岗岩": (800, False),
    "碎块状强风化花岗岩": (600, False), "强风化花岗岩": (450, False),
    "全风化花岗岩": (350, False),
    # 花岗片麻岩类
    "微风化花岗片麻岩": (1200, False), "中风化花岗片麻岩": (800, False),
    "碎块状强风化花岗片麻岩": (600, False), "强风化花岗片麻岩": (450, False),
    "全风化花岗片麻岩": (350, False),
    # 片麻岩类
    "微风化片麻岩": (1200, False), "中风化片麻岩": (800, False),
    "碎块状强风化片麻岩": (600, False), "强风化片麻岩": (450, False),
    "全风化片麻岩": (350, False),
    # 砂岩类
    "中风化砂岩": (750, False), "强风化砂岩": (430, False),
    "全风化砂岩": (320, False),
    # 通用风化 (兜底, 匹配优先级最低)
    "微风化": (1200, False), "中风化": (800, False),
    "强风化": (450, False), "全风化": (350, False),
    "含黏性土砂": (180, True), "含砂黏性土": (190, False), "含砾黏性土": (210, False),
}


def _match_layer(name: str) -> Tuple[Optional[Tuple[float, bool]], bool]:
    """返回 ((基准Vs, 是否为砂土), 是否精确匹配)"""
    name_clean = name.strip().replace(" ", "").replace("\u3000", "")
    if name_clean in LAYER_VS_BASE:
        return (LAYER_VS_BASE[name_clean], True)
    best_len, best = 0, None
    for key, val in LAYER_VS_BASE.items():
        if key in name_clean and len(key) > best_len:
            best_len, best = len(key), val
    return (best, False)


def _detect_state(name: str) -> Optional[str]:
    for state in STATE_MODIFIERS:
        if state in name:
            return state
    return None


def estimate_vs(layer_name: str) -> float:
    (matched, is_exact) = _match_layer(layer_name)
    if matched is None:
        return 180.0
    base_vs = matched[0]
    # 精确匹配时不再乘状态系数 (状态已编码在条目中)
    if is_exact:
        return base_vs
    state = _detect_state(layer_name)
    if state and state in STATE_MODIFIERS:
        return round(base_vs * STATE_MODIFIERS[state])
    return base_vs


def is_sand_layer(layer_name: str) -> bool:
    (matched, _) = _match_layer(layer_name)
    if matched is None:
        sand_kw = ("砂", "砾", "碎石", "卵石", "角砾", "圆砾", "漂石", "块石")
        return any(kw in layer_name for kw in sand_kw)
    return matched[1]


def compute_equivalent_vs(
    layers: List[Dict[str, Any]],
    cover_thickness: Optional[float] = None,
) -> float:
    """计算等效剪切波速 (GB 50011-2010 4.1.5)

    νse = d₀ / t, d₀ = min(覆盖层厚度, 20m)
    """
    if not layers:
        return 0.0
    if cover_thickness is None:
        cover_thickness = compute_cover_thickness(layers)
    d0 = min(cover_thickness, 20.0)
    if d0 <= 0:
        return 0.0
    total_time = 0.0
    cumulative_depth = 0.0
    for layer in layers:
        vs = layer.get("vs") or estimate_vs(layer.get("name", ""))
        thickness = layer.get("thickness", 0)
        if cumulative_depth + thickness > d0:
            total_time += (d0 - cumulative_depth) / vs
            break
        else:
            total_time += thickness / vs
            cumulative_depth += thickness
    return round(d0 / total_time, 1) if total_time > 0 else 0.0


def compute_cover_thickness(
    layers: List[Dict[str, Any]],
    vs_threshold: float = 500.0,
) -> float:
    cumulative = 0.0
    for i, layer in enumerate(layers):
        vs = layer.get("vs") or estimate_vs(layer.get("name", ""))
        thickness = layer.get("thickness", 0)
        if vs >= vs_threshold:
            # 确认下面所有层都 >= 500
            below_ok = True
            below_cum = cumulative + thickness
            for later in layers[i + 1:]:
                lvs = later.get("vs") or estimate_vs(later.get("name", ""))
                if lvs < vs_threshold:
                    below_ok = False
                    break
            if below_ok:
                return cumulative
        cumulative += thickness
    return sum(l.get("thickness", 0) for l in layers) + 1


def classify_site(vse: float, cover_thickness: float) -> Tuple[str, str]:
    d = cover_thickness
    if vse > 800:
        return ("I\u2080", f"\u03bdse={vse}m/s > 800, d={d}m \u2192 \u2160\u2080\u7c7b")
    elif vse > 500:
        if d == 0:
            return ("I\u2080", f"500 < \u03bdse={vse} \u2264 800, d=0 \u2192 \u2160\u2080\u7c7b")
        else:
            return ("I\u2081", f"500 < \u03bdse={vse} \u2264 800, d={d}m > 0 \u2192 \u2160\u2081\u7c7b")
    elif vse > 250:
        if d < 5:
            return ("I\u2081", f"250 < \u03bdse={vse} \u2264 500, d={d}m < 5 \u2192 \u2160\u2081\u7c7b")
        elif 5 <= d <= 50:
            return ("\u2161", f"250 < \u03bdse={vse} \u2264 500, d={d}m \u2208 [5,50] \u2192 \u2161\u7c7b")
        else:
            return ("\u2162", f"250 < \u03bdse={vse} \u2264 500, d={d}m > 50 \u2192 \u2162\u7c7b")
    elif vse > 150:
        if d < 3:
            return ("I\u2081", f"150 < \u03bdse={vse} \u2264 250, d={d}m < 3 \u2192 \u2160\u2081\u7c7b")
        elif 3 <= d <= 50:
            return ("\u2161", f"150 < \u03bdse={vse} \u2264 250, d={d}m \u2208 [3,50] \u2192 \u2161\u7c7b")
        else:
            return ("\u2162", f"150 < \u03bdse={vse} \u2264 250, d={d}m > 50 \u2192 \u2162\u7c7b")
    else:
        if d < 3:
            return ("I\u2081", f"\u03bdse={vse} \u2264 150, d={d}m < 3 \u2192 \u2160\u2081\u7c7b")
        elif 3 <= d <= 15:
            return ("\u2161", f"\u03bdse={vse} \u2264 150, d={d}m \u2208 [3,15] \u2192 \u2161\u7c7b")
        elif 15 < d <= 80:
            return ("\u2162", f"\u03bdse={vse} \u2264 150, d={d}m \u2208 (15,80] \u2192 \u2162\u7c7b")
        else:
            return ("\u2163", f"\u03bdse={vse} \u2264 150, d={d}m > 80 \u2192 \u2163\u7c7b")


def evaluate_site_class_from_layers(
    layers: List[Dict[str, Any]],
) -> Dict[str, Any]:
    total_depth = sum(l.get("thickness", 0) for l in layers)
    cover = compute_cover_thickness(layers)
    vse = compute_equivalent_vs(layers, cover)
    site_class, explanation = classify_site(vse, cover)
    detail = []
    cumulative = 0.0
    for l in layers:
        vs = l.get("vs") or estimate_vs(l.get("name", ""))
        thick = l.get("thickness", 0)
        t = thick / vs if vs > 0 else 0
        detail.append({
            "name": l.get("name", ""),
            "thickness": thick,
            "vs": vs,
            "time": round(t, 6),
            "depth_top": round(cumulative, 1),
            "depth_bottom": round(cumulative + thick, 1),
        })
        cumulative += thick
    return {
        "vse": vse, "cover_thickness": cover, "total_depth": total_depth,
        "site_class": site_class, "explanation": explanation,
        "layers_detail": detail,
    }


if __name__ == "__main__":
    test_2 = [
        {"name": "素填土", "thickness": 2.2},
        {"name": "细砂", "thickness": 1.3},
        {"name": "淤泥质粉质黏土", "thickness": 3.5},
        {"name": "粉质黏土", "thickness": 2.2},
        {"name": "中砂", "thickness": 8.1},
        {"name": "粉质黏土", "thickness": 2.7},
    ]
    test_49 = [
        {"name": "素填土", "thickness": 4.3},
        {"name": "细砂", "thickness": 1.2},
        {"name": "淤泥质粉质黏土", "thickness": 4.7},
        {"name": "粉质黏土", "thickness": 1.5},
        {"name": "中砂", "thickness": 7.9},
        {"name": "粉质黏土", "thickness": 0.4},
    ]
    for label, layers in [("钻孔2", test_2), ("钻孔49", test_49)]:
        r = evaluate_site_class_from_layers(layers)
        print(f"{label}: vse={r['vse']}m/s, d={r['cover_thickness']}m, 场地类别={r['site_class']}")
        print(f"  说明: {r['explanation']}")
        for l in r["layers_detail"]:
            print(f"    {l['name']:12s} 厚{l['thickness']:5.1f}m Vs={l['vs']:4.0f}m/s t={l['time']:.4f}s")
