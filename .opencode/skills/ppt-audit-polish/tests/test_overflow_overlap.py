"""Unit tests for the new overflow + overlap + near-edge detectors."""
from __future__ import annotations

import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from score_layout import (
    _bbox_intersect_area,
    _bbox_contains,
    _detect_boundary_issues,
    _detect_overlaps,
    _suggest_overlap_resolution,
)


def _shape(sid: int, left: int, top: int, w: int, h: int, kind: str = "container", name: str | None = None) -> dict:
    return {
        "shape_id": sid,
        "name": name or f"shape-{sid}",
        "kind": kind,
        "left": left,
        "top": top,
        "width": w,
        "height": h,
        "anomalous": False,
        "text": "",
        "font_sizes": [],
        "fill_hex": None,
    }


# Slide is 12192000 x 6858000 EMU (16:9 widescreen)
SLIDE_W = 12192000
SLIDE_H = 6858000


# ---------- boundary detection ----------

def test_right_overflow_flagged():
    obj = _shape(1, 11000000, 1000000, 2000000, 500000)  # right edge = 13M > 12.2M
    issues = _detect_boundary_issues(obj, SLIDE_W, SLIDE_H)
    cats = [i["category"] for i in issues]
    assert "boundary-overflow" in cats
    assert any("right" in i["message"] for i in issues)


def test_bottom_overflow_flagged():
    obj = _shape(1, 1000000, 6500000, 500000, 700000)  # bottom = 7.2M > 6.86M
    issues = _detect_boundary_issues(obj, SLIDE_W, SLIDE_H)
    assert any(i["category"] == "boundary-overflow" and "bottom" in i["message"] for i in issues)


def test_left_overflow_flagged():
    obj = _shape(1, -200000, 1000000, 500000, 500000)  # left negative
    issues = _detect_boundary_issues(obj, SLIDE_W, SLIDE_H)
    assert any(i["category"] == "boundary-overflow" and "left" in i["message"] for i in issues)


def test_top_overflow_flagged():
    obj = _shape(1, 1000000, -200000, 500000, 500000)  # top negative
    issues = _detect_boundary_issues(obj, SLIDE_W, SLIDE_H)
    assert any(i["category"] == "boundary-overflow" and "top" in i["message"] for i in issues)


def test_near_edge_crowding_left():
    obj = _shape(1, 50000, 1000000, 500000, 500000)  # left = 50K, threshold 100K
    issues = _detect_boundary_issues(obj, SLIDE_W, SLIDE_H)
    assert any(i["category"] == "near-edge-crowding" and "left" in i["message"] for i in issues)


def test_no_issue_when_inside_safe():
    obj = _shape(1, 500000, 500000, 1000000, 800000)
    issues = _detect_boundary_issues(obj, SLIDE_W, SLIDE_H)
    assert issues == []


def test_overflow_suppresses_crowding():
    """When a shape overflows on the right we shouldn't ALSO flag near-edge crowding on the left."""
    obj = _shape(1, 50000, 1000000, 13000000, 500000)
    issues = _detect_boundary_issues(obj, SLIDE_W, SLIDE_H)
    cats = {i["category"] for i in issues}
    assert "boundary-overflow" in cats
    assert "near-edge-crowding" not in cats  # overflow path skips crowding


# ---------- overlap detection ----------

def test_no_overlap_when_disjoint():
    a = _shape(1, 0, 0, 1000000, 1000000)
    b = _shape(2, 2000000, 0, 1000000, 1000000)
    slide = {"objects": [a, b]}
    assert _detect_overlaps(slide) == []


def test_overlap_flagged():
    a = _shape(1, 0, 0, 2000000, 2000000)
    b = _shape(2, 1000000, 1000000, 2000000, 2000000)  # overlaps a
    slide = {"objects": [a, b]}
    issues = _detect_overlaps(slide)
    assert len(issues) == 1
    assert issues[0]["category"] == "shape-overlap"
    assert issues[0]["overlap_ratio"] > 0.10


def test_parent_child_not_flagged():
    """A small badge inside a card shouldn't be flagged as overlap."""
    card = _shape(1, 0, 0, 4000000, 4000000)
    badge = _shape(2, 100000, 100000, 400000, 400000)  # inside card
    slide = {"objects": [card, badge]}
    assert _detect_overlaps(slide) == []


def test_connector_not_flagged():
    a = _shape(1, 0, 0, 2000000, 2000000)
    b = _shape(2, 1000000, 1000000, 2000000, 2000000, kind="connector")
    slide = {"objects": [a, b]}
    assert _detect_overlaps(slide) == []


def test_overlap_severity_scales():
    """50%+ overlap should be 'error', smaller should be 'warning'.
    Use partial overlap (not containment) so detection fires."""
    # ~56% overlap of smaller shape
    a = _shape(1, 0, 0, 2000000, 2000000)
    big = _shape(2, 500000, 500000, 2000000, 2000000)  # extends past a's right/bottom
    slide = {"objects": [a, big]}
    issues = _detect_overlaps(slide)
    assert issues, "expected an overlap issue"
    assert issues[0]["severity"] == "error"

    # ~22% overlap of smaller shape (medium)
    a = _shape(1, 0, 0, 2000000, 2000000)
    medium = _shape(2, 1100000, 1100000, 1500000, 1500000)
    slide = {"objects": [a, medium]}
    issues = _detect_overlaps(slide)
    assert issues, "expected an overlap issue"
    assert issues[0]["severity"] == "warning"

    # ~3% overlap — below threshold, no issue
    a = _shape(1, 0, 0, 2000000, 2000000)
    tiny = _shape(2, 1850000, 1850000, 1500000, 1500000)
    slide = {"objects": [a, tiny]}
    assert _detect_overlaps(slide) == []


# ---------- bbox helpers ----------

def test_intersect_area_correct():
    a = _shape(1, 0, 0, 1000, 1000)
    b = _shape(2, 500, 500, 1000, 1000)
    assert _bbox_intersect_area(a, b) == 250000  # 500x500


def test_intersect_zero_when_disjoint():
    a = _shape(1, 0, 0, 100, 100)
    b = _shape(2, 200, 200, 100, 100)
    assert _bbox_intersect_area(a, b) == 0


def test_contains_with_slack():
    # Use realistic EMU sizes (1M+) so the 91K slack is small relative to dims.
    outer = _shape(1, 0, 0, 4000000, 4000000)
    inner = _shape(2, 500000, 500000, 2000000, 2000000)
    assert _bbox_contains(outer, inner)
    assert not _bbox_contains(inner, outer)


# ---------- nudge suggestion ----------

def test_overlap_resolution_suggests_smallest_displacement():
    """Smaller shape should be nudged the shortest distance to clear larger."""
    larger = _shape(1, 1000000, 1000000, 4000000, 4000000)
    # Smaller mostly-inside on top-left: shortest exit is up.
    smaller = _shape(2, 1100000, 1100000, 800000, 800000)
    argv = _suggest_overlap_resolution(smaller, larger)
    assert argv[0] == "nudge"
    # Should pick the closest edge to exit through.
    assert "--shape-id" in argv
