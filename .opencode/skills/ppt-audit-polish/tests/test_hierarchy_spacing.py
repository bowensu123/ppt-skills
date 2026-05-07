"""Unit tests for _hierarchy_spacing.py — vertical-rhythm rules between
title/subtitle/body roles."""
from __future__ import annotations

import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from _hierarchy_spacing import detect_hierarchy_spacing_issues


def _shape(sid, top, height, kind="text"):
    return {
        "shape_id": sid, "name": f"sh-{sid}", "kind": kind,
        "left": 0, "top": top, "width": 12192000, "height": height,
        "anomalous": False,
    }


def _slide(*shapes):
    return {
        "slide_index": 1,
        "width_emu": 12192000,
        "height_emu": 6858000,
        "objects": list(shapes),
    }


def _roles(**role_to_sid):
    return {"shapes": [{"shape_id": sid, "role": role}
                       for role, sid in role_to_sid.items()]}


# ---- title → subtitle ratio ----

def test_title_to_subtitle_ideal_no_issue():
    """Title bottom at 800K, subtitle at top 1.2M (gap=400K).
    Title height=600K → ratio = 400K/600K = 0.67, in [0.30, 0.80] → OK."""
    title    = _shape(1, top=200000,  height=600000)
    subtitle = _shape(2, top=1200000, height=300000)
    issues = detect_hierarchy_spacing_issues(
        _slide(title, subtitle), _roles(title=1, subtitle=2),
    )
    assert issues == []


def test_title_to_subtitle_too_tight_fires():
    """Tiny gap below title → ratio < 0.30 → issue, dy positive (move down)."""
    title    = _shape(1, top=200000, height=1000000)   # bottom at 1.2M
    subtitle = _shape(2, top=1250000, height=300000)   # gap = 50000
    issues = detect_hierarchy_spacing_issues(
        _slide(title, subtitle), _roles(title=1, subtitle=2),
    )
    assert len(issues) == 1
    issue = issues[0]
    assert issue["category"] == "hierarchy-spacing-drift"
    argv = issue["suggested_argv"]
    assert argv[0] == "nudge"
    assert int(argv[argv.index("--dy") + 1]) > 0  # move subtitle DOWN


def test_title_to_subtitle_too_loose_fires():
    """Big gap below title → ratio > 0.80 → issue, dy negative (move up)."""
    title    = _shape(1, top=200000,  height=600000)    # bottom at 800K
    subtitle = _shape(2, top=2500000, height=300000)    # gap = 1.7M = 2.83x
    issues = detect_hierarchy_spacing_issues(
        _slide(title, subtitle), _roles(title=1, subtitle=2),
    )
    assert len(issues) == 1
    argv = issues[0]["suggested_argv"]
    assert int(argv[argv.index("--dy") + 1]) < 0  # move subtitle UP


# ---- subtitle → body ratio ----

def test_subtitle_to_body_too_loose_fires():
    title    = _shape(1, top=200000,  height=600000)
    subtitle = _shape(2, top=1100000, height=300000)
    body     = _shape(3, top=2500000, height=2000000)  # gap=1.1M, ratio 3.67x
    issues = detect_hierarchy_spacing_issues(
        _slide(title, subtitle, body),
        _roles(title=1, subtitle=2, body=3),
    )
    cats = [i["category"] for i in issues]
    # title-subtitle should be OK (gap=300K, ratio 0.5x), only subtitle-body fires.
    drift_issues = [i for i in issues if i["category"] == "hierarchy-spacing-drift"]
    assert any(i["rhythm_report"]["below_role"] == "body" for i in drift_issues)


# ---- title → body (no subtitle) ----

def test_title_to_body_no_subtitle_works():
    title = _shape(1, top=200000,  height=600000)   # bottom 800K
    body  = _shape(2, top=900000,  height=2000000)  # gap=100K, ratio 0.17x
    issues = detect_hierarchy_spacing_issues(
        _slide(title, body), _roles(title=1, body=2),
    )
    # 0.17 < 0.80 → too tight, fires.
    assert len(issues) == 1
    assert issues[0]["rhythm_report"]["below_role"] == "body"


# ---- silence cases ----

def test_no_title_no_issue():
    body = _shape(1, top=500000, height=300000)
    issues = detect_hierarchy_spacing_issues(_slide(body), _roles(body=1))
    assert issues == []


def test_overlapping_shapes_skipped():
    """If shapes overlap (gap < MIN_MEASURABLE_GAP_EMU), don't fire — that's
    a different kind of problem (handled by overlap detector)."""
    title    = _shape(1, top=200000, height=600000)   # bottom 800K
    subtitle = _shape(2, top=750000, height=300000)   # OVERLAPS (gap=-50000)
    issues = detect_hierarchy_spacing_issues(
        _slide(title, subtitle), _roles(title=1, subtitle=2),
    )
    assert issues == []


def test_rhythm_report_in_issue():
    """Each issue should include the full numerical context for the agent."""
    title    = _shape(1, top=200000,  height=600000)
    subtitle = _shape(2, top=2500000, height=300000)
    issues = detect_hierarchy_spacing_issues(
        _slide(title, subtitle), _roles(title=1, subtitle=2),
    )
    rep = issues[0]["rhythm_report"]
    assert rep["above_shape_id"] == 1
    assert rep["below_shape_id"] == 2
    assert rep["above_role"] == "title"
    assert rep["below_role"] == "subtitle"
    assert rep["current_ratio"] > 0.80
    assert "proposed_dy" in rep
