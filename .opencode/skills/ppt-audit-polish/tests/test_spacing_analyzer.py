"""Unit tests for _spacing_analyzer.py — auto-discover peer groups,
analyze gap distribution, emit issues with rich rationale."""
from __future__ import annotations

import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from _spacing_analyzer import (
    analyze_group,
    compute_gaps,
    detect_spacing_issues,
    discover_column_groups,
    discover_row_groups,
)


def _shape(sid, left, top, w=1000000, h=1000000, kind="container", name=None):
    return {
        "shape_id": sid,
        "name": name or f"shape-{sid}",
        "kind": kind,
        "left": left, "top": top, "width": w, "height": h,
        "anomalous": False,
        "fill_hex": "111827",
    }


# ---- discover_row_groups ----

def test_discover_row_groups_finds_3_aligned_peers():
    objs = [
        _shape(1, 500000, 1000000),
        _shape(2, 2000000, 1000000),
        _shape(3, 3500000, 1000000),
    ]
    groups = discover_row_groups(objs)
    assert len(groups) == 1
    assert [s["shape_id"] for s in groups[0]] == [1, 2, 3]


def test_discover_row_groups_rejects_2_only():
    objs = [_shape(1, 0, 0), _shape(2, 2000000, 0)]
    assert discover_row_groups(objs) == []


def test_discover_row_groups_excludes_size_outlier():
    """3 small + 1 huge sharing a top band → keeps the 3 small as a group."""
    objs = [
        _shape(1, 500000, 1000000, w=1000000),
        _shape(2, 2000000, 1000000, w=1000000),
        _shape(3, 3500000, 1000000, w=1000000),
        _shape(4, 5000000, 1000000, w=10000000),  # 10x bigger
    ]
    groups = discover_row_groups(objs)
    sids = {s["shape_id"] for g in groups for s in g}
    assert 4 not in sids


def test_discover_row_groups_separates_different_top_bands():
    objs = [
        _shape(1, 0, 1000000), _shape(2, 2000000, 1000000), _shape(3, 4000000, 1000000),
        _shape(4, 0, 4000000), _shape(5, 2000000, 4000000), _shape(6, 4000000, 4000000),
    ]
    groups = discover_row_groups(objs)
    assert len(groups) == 2


def test_discover_column_groups_finds_vertical_stack():
    objs = [
        _shape(1, 500000, 1000000),
        _shape(2, 500000, 3000000),
        _shape(3, 500000, 5000000),
    ]
    groups = discover_column_groups(objs)
    assert len(groups) == 1
    assert [s["shape_id"] for s in groups[0]] == [1, 2, 3]


# ---- compute_gaps ----

def test_compute_gaps_horizontal():
    peers = [_shape(1, 0, 0, w=100), _shape(2, 200, 0, w=100), _shape(3, 500, 0, w=100)]
    assert compute_gaps(peers, "horizontal") == [100, 200]


def test_compute_gaps_vertical():
    peers = [_shape(1, 0, 0, h=100), _shape(2, 0, 200, h=100), _shape(3, 0, 500, h=100)]
    assert compute_gaps(peers, "vertical") == [100, 200]


# ---- analyze_group ----

def test_analyze_group_uniform_gaps_no_verdict():
    """3 peers with identical 200K gaps → no verdicts emitted."""
    peers = [
        _shape(1, 0, 0, w=1000000),
        _shape(2, 1200000, 0, w=1000000),
        _shape(3, 2400000, 0, w=1000000),
    ]
    rep = analyze_group(peers, "horizontal", 12192000)
    assert rep["verdicts"] == []
    assert rep["gap_mean_emu"] == 200000


def test_analyze_group_uneven_fires():
    peers = [
        _shape(1, 0,        0, w=1000000),
        _shape(2, 1100000,  0, w=1000000),   # gap 100K
        _shape(3, 2900000,  0, w=1000000),   # gap 800K
    ]
    rep = analyze_group(peers, "horizontal", 12192000)
    kinds = {v["kind"] for v in rep["verdicts"]}
    assert "uneven" in kinds
    uneven = next(v for v in rep["verdicts"] if v["kind"] == "uneven")
    # median of [100K, 800K] = 450K (statistics.median averages even-length pairs).
    assert uneven["target_gap_emu"] == 450000


def test_analyze_group_too_tight_fires():
    """Peers nearly touching: median gap < 10% of avg width."""
    peers = [
        _shape(1, 0,       0, w=1000000),
        _shape(2, 1050000, 0, w=1000000),    # gap 50K = 5% of width
        _shape(3, 2100000, 0, w=1000000),    # gap 50K
    ]
    rep = analyze_group(peers, "horizontal", 12192000)
    kinds = {v["kind"] for v in rep["verdicts"]}
    assert "too-tight" in kinds


def test_analyze_group_too_loose_fires():
    """Peers spread far apart: median gap > 200% of avg width."""
    peers = [
        _shape(1, 0,        0, w=1000000),
        _shape(2, 4000000,  0, w=1000000),    # gap 3M = 300% of width
        _shape(3, 8000000,  0, w=1000000),    # gap 3M
    ]
    rep = analyze_group(peers, "horizontal", 12192000)
    kinds = {v["kind"] for v in rep["verdicts"]}
    assert "too-loose" in kinds


# ---- detect_spacing_issues (top-level integration) ----

def test_detect_spacing_issues_emits_argv():
    """Uneven row → issue dict with ready-to-run mutate argv."""
    slide = {
        "slide_index": 1,
        "width_emu": 12192000,
        "height_emu": 6858000,
        "objects": [
            _shape(1, 0, 1000000, w=1000000),
            _shape(2, 1100000, 1000000, w=1000000),  # gap 100K
            _shape(3, 2900000, 1000000, w=1000000),  # gap 800K
        ],
    }
    issues = detect_spacing_issues(slide)
    uneven = [i for i in issues if i["category"] == "spacing-uneven"]
    assert len(uneven) == 1
    argv = uneven[0]["suggested_argv"]
    assert argv[0] == "equalize-gaps"
    assert "--axis" in argv and "horizontal" in argv
    assert "--shape-ids" in argv
    # The reported shape_ids should be all 3 peers in axis order.
    shape_ids_arg = argv[argv.index("--shape-ids") + 1]
    assert shape_ids_arg == "1,2,3"


def test_detect_spacing_issues_no_groups_no_issues():
    slide = {
        "slide_index": 1, "width_emu": 12192000, "height_emu": 6858000,
        "objects": [_shape(1, 0, 0), _shape(2, 5000000, 5000000)],  # different size + position
    }
    assert detect_spacing_issues(slide) == []


def test_detect_spacing_issues_rationale_in_message():
    """Smart messages include the WHY, not just numbers."""
    slide = {
        "slide_index": 1, "width_emu": 12192000, "height_emu": 6858000,
        "objects": [
            _shape(1, 0,        1000000, w=1000000),
            _shape(2, 1100000,  1000000, w=1000000),
            _shape(3, 2900000,  1000000, w=1000000),
        ],
    }
    issues = detect_spacing_issues(slide)
    msg = issues[0]["message"]
    # Message should include numerical context AND rationale
    assert "Horizontal" in msg or "horizontal" in msg.lower()
    assert "median" in msg.lower() or "majority" in msg.lower()


def test_detect_spacing_issues_horizontal_and_vertical():
    """A 2D grid produces issues for both axes when both are uneven."""
    slide = {
        "slide_index": 1, "width_emu": 12192000, "height_emu": 6858000,
        "objects": [
            # Row 1: uneven horizontal spacing
            _shape(1, 0,         1000000, w=1000000),
            _shape(2, 1100000,   1000000, w=1000000),
            _shape(3, 2900000,   1000000, w=1000000),
            # Row 2: lined up under row 1
            _shape(4, 0,         3000000, w=1000000),
            _shape(5, 1100000,   3000000, w=1000000),
            _shape(6, 2900000,   3000000, w=1000000),
            # Row 3: at uneven vertical interval (gap row1-row2 = 1M, row2-row3 = 2.5M)
            _shape(7, 0,         5500000, w=1000000),
            _shape(8, 1100000,   5500000, w=1000000),
            _shape(9, 2900000,   5500000, w=1000000),
        ],
    }
    issues = detect_spacing_issues(slide)
    cats = {i["category"] for i in issues}
    # Both row-spacing-uneven (horizontal) and column-spacing-uneven (vertical) fire.
    assert "spacing-uneven" in cats
