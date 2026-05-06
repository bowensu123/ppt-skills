"""Unit tests for the 6 blind-spot modules added today."""
from __future__ import annotations

import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from _color_harmony import detect_color_issues, _pick_legible_text_color, _find_container
from _visual_balance import compute_balance
from _image_check import detect_image_issues
from _table_chart import detect_table_chart_issues


def _shape(sid, left, top, w, h, kind="container", text="", fill=None, text_color=None, font_pt=None, **extra):
    obj = {
        "shape_id": sid,
        "name": f"shape-{sid}",
        "kind": kind,
        "left": left, "top": top, "width": w, "height": h,
        "anomalous": False, "text": text,
        "font_sizes": [int(font_pt * 12700)] if font_pt else [],
        "fill_hex": fill, "text_color": text_color,
    }
    obj.update(extra)
    return obj


SLIDE_W = 12192000
SLIDE_H = 6858000


# ---------- color harmony ----------

def test_low_contrast_dark_on_dark():
    text = _shape(1, 100000, 100000, 2000000, 500000, kind="text", text="hello",
                  text_color="555555", font_pt=11.0)
    bg = _shape(2, 0, 0, 4000000, 4000000, kind="container", fill="333333")
    slide = {"objects": [bg, text], "width_emu": SLIDE_W, "height_emu": SLIDE_H}
    issues = detect_color_issues(slide)
    cats = [i["category"] for i in issues]
    assert "low-contrast-text" in cats


def test_high_contrast_no_issue():
    text = _shape(1, 100000, 100000, 2000000, 500000, kind="text", text="hello",
                  text_color="161616", font_pt=11.0)
    bg = _shape(2, 0, 0, 4000000, 4000000, kind="container", fill="FFFFFF")
    slide = {"objects": [bg, text], "width_emu": SLIDE_W, "height_emu": SLIDE_H}
    issues = detect_color_issues(slide)
    assert all(i["category"] != "low-contrast-text" for i in issues)


def test_palette_too_busy_flagged():
    objects = [
        _shape(i, 0, 0, 100000, 100000, fill=f"{i:02X}{i:02X}{i:02X}")
        for i in range(1, 9)
    ]
    slide = {"objects": objects, "width_emu": SLIDE_W, "height_emu": SLIDE_H}
    issues = detect_color_issues(slide)
    assert any(i["category"] == "palette-too-busy" for i in issues)


def test_pick_legible_text_color():
    assert _pick_legible_text_color("000000") == "FFFFFF"
    assert _pick_legible_text_color("FFFFFF") == "161616"


def test_find_container_picks_smallest():
    text = _shape(1, 1000000, 1000000, 500000, 200000, kind="text", text="x", text_color="000000")
    big = _shape(2, 0, 0, 8000000, 8000000, fill="FFFFFF")
    small = _shape(3, 800000, 800000, 1000000, 1000000, fill="EEEEEE")
    container = _find_container(text, [text, big, small])
    assert container["shape_id"] == 3


# ---------- visual balance ----------

def test_centered_layout_high_score():
    # one big shape exactly in slide center
    obj = _shape(1, SLIDE_W // 4, SLIDE_H // 4, SLIDE_W // 2, SLIDE_H // 2)
    slide = {"objects": [obj], "width_emu": SLIDE_W, "height_emu": SLIDE_H}
    result = compute_balance(slide)
    assert result["score"] >= 95.0
    assert result["issue"] is None


def test_far_left_imbalance_flagged():
    # shape entirely on the left side
    obj = _shape(1, 0, SLIDE_H // 4, SLIDE_W // 6, SLIDE_H // 2)
    slide = {"objects": [obj], "width_emu": SLIDE_W, "height_emu": SLIDE_H}
    result = compute_balance(slide)
    assert result["imbalance_x"] > 0.5
    assert result["issue"] is not None
    assert result["issue"]["category"] == "visual-imbalance"


def test_empty_slide_neutral_score():
    slide = {"objects": [], "width_emu": SLIDE_W, "height_emu": SLIDE_H}
    result = compute_balance(slide)
    assert result["score"] == 50.0


# ---------- image check ----------

def test_picture_heavy_crop_flagged():
    pic = _shape(1, 0, 0, 1000000, 1000000, kind="picture",
                 crop={"l": 35000, "t": 0, "r": 0, "b": 0})  # 35% left crop
    slide = {"objects": [pic]}
    issues = detect_image_issues(slide)
    cats = [i["category"] for i in issues]
    assert "picture-heavy-crop" in cats


def test_picture_light_crop_info():
    pic = _shape(1, 0, 0, 1000000, 1000000, kind="picture",
                 crop={"l": 15000, "t": 0, "r": 0, "b": 0})  # 15% left
    slide = {"objects": [pic]}
    issues = detect_image_issues(slide)
    assert any(i["category"] == "picture-cropped" for i in issues)


def test_picture_no_crop_quiet():
    pic = _shape(1, 0, 0, 1000000, 1000000, kind="picture",
                 crop={"l": 0, "t": 0, "r": 0, "b": 0})
    slide = {"objects": [pic]}
    assert detect_image_issues(slide) == []


# ---------- table / chart ----------

def test_table_too_many_rows_flagged():
    tbl = _shape(1, 0, 0, 5000000, 4000000, kind="table",
                 table_info={"rows": 15, "cols": 3, "empty_cells": 0})
    slide = {"objects": [tbl]}
    issues = detect_table_chart_issues(slide)
    assert any(i["category"] == "table-too-many-rows" for i in issues)


def test_table_many_empty_cells_flagged():
    tbl = _shape(1, 0, 0, 5000000, 4000000, kind="table",
                 table_info={"rows": 5, "cols": 4, "empty_cells": 8})  # 40% empty
    slide = {"objects": [tbl]}
    issues = detect_table_chart_issues(slide)
    assert any(i["category"] == "table-many-empty-cells" for i in issues)


def test_chart_too_many_series_flagged():
    chart = _shape(1, 0, 0, 5000000, 4000000, kind="chart",
                   chart_info={"chart_type": "LINE", "series_count": 12, "has_legend": True})
    slide = {"objects": [chart]}
    issues = detect_table_chart_issues(slide)
    assert any(i["category"] == "chart-too-many-series" for i in issues)


def test_chart_missing_legend_with_many_series_flagged():
    chart = _shape(1, 0, 0, 5000000, 4000000, kind="chart",
                   chart_info={"chart_type": "BAR", "series_count": 5, "has_legend": False})
    slide = {"objects": [chart]}
    issues = detect_table_chart_issues(slide)
    assert any(i["category"] == "chart-missing-legend" for i in issues)


def test_simple_chart_no_issue():
    chart = _shape(1, 0, 0, 5000000, 4000000, kind="chart",
                   chart_info={"chart_type": "PIE", "series_count": 1, "has_legend": True})
    slide = {"objects": [chart]}
    assert detect_table_chart_issues(slide) == []


# ---------- group recursion (smoke; full coverage via inspection) ----------

def test_group_recurse_module_imports():
    from _group_recurse import walk_group_children, _read_group_xfrm, _apply
    # transforms should be identity-ish when off==chOff and ext==chExt
    transform = {"off_x": 0, "off_y": 0, "ext_cx": 1000, "ext_cy": 1000, "chOff_x": 0, "chOff_y": 0, "chExt_cx": 1000, "chExt_cy": 1000}
    assert _apply(transform, 100, 200) == (100, 200)
    # scaled transform: child space is half of slide space
    transform2 = {"off_x": 5000, "off_y": 5000, "ext_cx": 2000, "ext_cy": 2000, "chOff_x": 0, "chOff_y": 0, "chExt_cx": 1000, "chExt_cy": 1000}
    # child at (100, 200) in 1000-space -> (100*2, 200*2) + offset
    assert _apply(transform2, 100, 200) == (5200, 5400)


def test_master_inherit_module_imports():
    from _master_inherit import has_inherited_title, collect_inheritance_info
    assert has_inherited_title({}) is False
    assert has_inherited_title({"layout_placeholders": [{"type": "TITLE", "name": "t", "idx": 0}]}) is True
