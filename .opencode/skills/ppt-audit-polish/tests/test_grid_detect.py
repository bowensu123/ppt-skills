"""Unit tests for _grid_detect.py — 2D grid detection + outlier identification."""
from __future__ import annotations

import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from _grid_detect import (
    _cluster_1d,
    _candidate_panels,
    _largest_size_cohort,
    _detect_header_strip_outliers,
    detect_grid,
    detect_grids_nested,
    diagnose_grid_repair,
)


def _panel(sid, left, top, w=5400000, h=1700000, kind="container", name=None):
    """Default size matches a "panel-like" container."""
    return {
        "shape_id": sid,
        "name": name or f"shape-{sid}",
        "kind": kind,
        "left": left, "top": top, "width": w, "height": h,
        "anomalous": False,
        "fill_hex": "111827",
    }


# ---------- _cluster_1d ----------

def test_cluster_simple():
    pts = [(100, 1), (110, 2), (300, 3), (310, 4)]
    clusters = _cluster_1d(pts, eps=50)
    assert len(clusters) == 2
    assert sorted(clusters[0]) == [1, 2]
    assert sorted(clusters[1]) == [3, 4]


def test_cluster_single_value():
    pts = [(100, 1)]
    assert _cluster_1d(pts, eps=50) == [[1]]


def test_cluster_empty():
    assert _cluster_1d([], eps=50) == []


# ---------- candidate filtering ----------

def test_candidates_excludes_small():
    objs = [
        _panel(1, 0, 0, w=900000, h=700000),     # ok
        _panel(2, 0, 0, w=300000, h=300000),     # too small
        _panel(3, 0, 0, w=900000, h=200000),     # too short
    ]
    cands = _candidate_panels(objs)
    assert {c["shape_id"] for c in cands} == {1}


# ---------- detect_grid: perfect 2x3 ----------

def test_perfect_2x3_no_outliers():
    panels = []
    for r, top in enumerate([1000000, 3000000, 5000000]):
        for c, left in enumerate([500000, 6500000]):
            panels.append(_panel(r * 10 + c, left, top))
    grid = detect_grid(panels)
    assert grid is not None
    assert grid["rows"] == 3
    assert grid["cols"] == 2
    assert grid["outliers"] == []


# ---------- detect_grid: 2x3 with one outlier (vlm-style) ----------

def test_2x3_with_outlier():
    panels = [
        _panel(6, 700000, 400000),       # ← OUTLIER (should be 500000, 1000000)
        _panel(21, 6500000, 1000000),
        _panel(36, 500000, 3000000),
        _panel(51, 6500000, 3000000),
        _panel(66, 500000, 5000000),
        _panel(81, 6500000, 5000000),
    ]
    grid = detect_grid(panels)
    assert grid is not None
    assert grid["rows"] == 3 and grid["cols"] == 2
    outlier_ids = {o["shape_id"] for o in grid["outliers"]}
    assert 6 in outlier_ids
    # The fix should snap sid=6 toward (500000, 1000000).
    fix_for_6 = next(o for o in grid["outliers"] if o["shape_id"] == 6)
    target_left, target_top = fix_for_6["target"]
    assert abs(target_left - 500000) < 100000
    assert abs(target_top - 1000000) < 100000


# ---------- detect_grid: 3x3 ----------

def test_3x3_perfect():
    panels = []
    for r, top in enumerate([1000000, 3000000, 5000000]):
        for c, left in enumerate([500000, 4000000, 7500000]):
            panels.append(_panel(r * 10 + c, left, top, w=3000000))
    grid = detect_grid(panels)
    assert grid is not None
    assert grid["rows"] == 3 and grid["cols"] == 3
    assert grid["outliers"] == []


# ---------- detect_grid: incomplete grid ----------

def test_too_few_panels_no_grid():
    panels = [_panel(1, 0, 0), _panel(2, 6000000, 0)]
    assert detect_grid(panels) is None


def test_three_panels_rejected_too_few():
    # Smallest grid we recognize is 2x2. 3 panels can't form one.
    panels = [
        _panel(1, 500000, 1000000),
        _panel(2, 4000000, 1000000),
        _panel(3, 500000, 5000000),
    ]
    grid = detect_grid(panels)
    assert grid is None


# ---------- size-cohort filtering ----------

def test_size_cohort_excludes_outlier_size():
    panels = [
        _panel(1, 500000, 1000000, w=5400000, h=1700000),
        _panel(2, 6500000, 1000000, w=5400000, h=1700000),
        _panel(3, 500000, 3000000, w=5400000, h=1700000),
        _panel(4, 6500000, 3000000, w=5400000, h=1700000),
        _panel(5, 0, 0, w=12000000, h=400000),  # huge title bar — different size
    ]
    cohort = _largest_size_cohort(panels)
    assert 5 not in {p["shape_id"] for p in cohort}


# ---------- nested ----------

def test_nested_grids_finds_sub_grid():
    """Outer 2x2 of big panels, each containing 2x2 of sub-cells."""
    objs = []
    big_w, big_h = 5000000, 3000000
    sub_w, sub_h = 2000000, 1000000
    for big_r, big_top in enumerate([0, big_h + 200000]):
        for big_c, big_left in enumerate([0, big_w + 200000]):
            big_sid = 100 + big_r * 10 + big_c
            objs.append(_panel(big_sid, big_left, big_top, w=big_w, h=big_h))
            # 2x2 sub-cells inside
            for sub_r, sub_off_top in enumerate([200000, 1500000]):
                for sub_c, sub_off_left in enumerate([200000, 2700000]):
                    sub_sid = big_sid * 10 + sub_r * 2 + sub_c
                    objs.append(_panel(
                        sub_sid,
                        big_left + sub_off_left,
                        big_top + sub_off_top,
                        w=sub_w, h=sub_h,
                    ))
    grids = detect_grids_nested(objs, max_depth=2)
    assert len(grids) >= 1
    outer = next((g for g in grids if g["depth"] == 0), None)
    assert outer is not None
    assert outer["rows"] == 2 and outer["cols"] == 2
    nested = [g for g in grids if g["depth"] == 1]
    assert len(nested) >= 1


# ---------- header-strip relative-offset outliers ----------

def _strip(sid, left, top, w=5400000, h=200000, fill="00D4FF"):
    return {
        "shape_id": sid,
        "name": f"strip-{sid}",
        "kind": "container",
        "left": left, "top": top, "width": w, "height": h,
        "anomalous": False,
        "fill_hex": fill,
    }


def test_header_strip_outlier_detected():
    """A 2x3 grid where one panel's header strip floats 274K EMU above the
    panel top (vlm_data_agent_infra reproducer). All other strips sit at
    panel.top + 0. Detector should flag the outlier and propose a snap to
    the peer-median offset."""
    panels = []
    strips = []
    sid = 100
    for r, top in enumerate([1000000, 3000000, 5000000]):
        for c, left in enumerate([500000, 6500000]):
            panels.append(_panel(sid, left, top))
            strips.append(_strip(sid + 1, left, top))
            sid += 10

    # Make one strip float 274320 EMU above its panel.
    bad_strip = strips[2]   # row 1, col 0
    bad_panel = panels[2]
    bad_strip["top"] = bad_panel["top"] - 274320

    grid = detect_grid(panels)
    assert grid is not None
    panel_lookup = {p["shape_id"]: p for p in panels}
    fixes = _detect_header_strip_outliers(grid, panel_lookup, panels + strips)
    assert any(f["shape_id"] == bad_strip["shape_id"] for f in fixes)
    fix = next(f for f in fixes if f["shape_id"] == bad_strip["shape_id"])
    assert "top" in fix
    assert abs(fix["top"] - bad_panel["top"]) < 50000  # snap to panel top


def test_header_strip_no_outlier_when_all_aligned():
    """All strips sit at the same relative offset → no fixes."""
    panels = []
    strips = []
    sid = 100
    for r, top in enumerate([1000000, 3000000, 5000000]):
        for c, left in enumerate([500000, 6500000]):
            panels.append(_panel(sid, left, top))
            strips.append(_strip(sid + 1, left, top))
            sid += 10
    grid = detect_grid(panels)
    panel_lookup = {p["shape_id"]: p for p in panels}
    fixes = _detect_header_strip_outliers(grid, panel_lookup, panels + strips)
    assert fixes == []


# ---------- diagnose API shape ----------

def test_diagnose_grid_repair_format():
    panels = [
        _panel(6, 700000, 400000),
        _panel(21, 6500000, 1000000),
        _panel(36, 500000, 3000000),
        _panel(51, 6500000, 3000000),
        _panel(66, 500000, 5000000),
        _panel(81, 6500000, 5000000),
    ]
    plan = diagnose_grid_repair({"objects": panels})
    assert "rows" in plan
    assert plan["rows"], "expected at least one row entry"
    first = plan["rows"][0]
    assert "card_box_fixes" in first
    assert any(fix["shape_id"] == 6 for fix in first["card_box_fixes"])
