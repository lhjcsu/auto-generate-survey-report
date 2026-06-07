#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""基础冒烟测试 — 验证各模块可正常导入和实例化"""

import sys
import os

# 确保项目目录在 path 中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def test_imports():
    """验证所有模块可正常导入"""
    from corrosion_eval import evaluate_corrosion
    from seismic_lookup import get_seismic_params
    from wave_velocity_estimate import estimate_vs, compute_equivalent_vs
    from wave_velocity_parser import load_wave_velocity_data
    print("[PASS] 所有模块导入成功")


def test_corrosion_eval():
    """验证腐蚀性评价模块基本功能"""
    from corrosion_eval import evaluate_corrosion
    result = evaluate_corrosion([], [])
    assert isinstance(result, dict)
    print("[PASS] corrosion_eval 基本功能")


def test_seismic_lookup():
    """验证地震参数查表模块"""
    from seismic_lookup import get_seismic_params
    result = get_seismic_params("环翠区", "怡园街道")
    assert result is not None
    assert "pga" in result or "intensity" in result
    print(f"[PASS] seismic_lookup: {result}")


def test_wave_velocity_estimate():
    """验证波速估算模块"""
    from wave_velocity_estimate import estimate_vs
    vs = estimate_vs("粉质黏土")
    assert vs > 0
    print(f"[PASS] wave_velocity_estimate: 粉质黏土 -> Vs={vs}")


def test_generate_report_syntax():
    """验证主程序语法正确"""
    import py_compile
    py_compile.compile("generate_report.py", doraise=True)
    py_compile.compile("corrosion_eval.py", doraise=True)
    py_compile.compile("seismic_lookup.py", doraise=True)
    py_compile.compile("wave_velocity_estimate.py", doraise=True)
    py_compile.compile("wave_velocity_parser.py", doraise=True)
    print("[PASS] 所有 .py 语法检查通过")

def test_config_example_valid():
    """验证示例配置文件是合法 JSON"""
    import json
    with open("project_config.example.json", "r", encoding="utf-8") as f:
        cfg = json.load(f)
    assert "project_name" in cfg
    assert "layers" in cfg
    assert len(cfg["layers"]) > 0
    print(f"[PASS] 配置文件合法, {len(cfg['layers'])} 个地层")


if __name__ == "__main__":
    tests = [
        test_imports,
        test_corrosion_eval,
        test_seismic_lookup,
        test_wave_velocity_estimate,
        test_generate_report_syntax,
        test_config_example_valid,
    ]
    
    passed = 0
    failed = 0
    for t in tests:
        try:
            t()
            passed += 1
        except Exception as e:
            print(f"[FAIL] {t.__name__}: {e}")
            failed += 1
    
    print(f"\n{'='*40}")
    print(f"  通过: {passed}/{passed+failed}")
    if failed:
        print(f"  失败: {failed}")
    print(f"{'='*40}")
    sys.exit(1 if failed else 0)
