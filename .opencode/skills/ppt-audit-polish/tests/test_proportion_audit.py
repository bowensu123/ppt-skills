"""Tests for _proportion_audit.py — visual composition audits."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from _proportion_audit import (
    ICON_AREA_RATIO_MAX,
    ICON_AREA_RATIO_MIN,
    ICON_AREA_RATIO_TARGET,
    _is_card,
    _is_icon_like,
    _max_font_pt,
    detect_proportion_issues,
)


# ---- shape classification ----

def test_is_card_yes():
    obj = {"kind": "container", "width": 5000000, "height": 3000000,
           "fill_hex": "111827", "anomalous": False}
    assert _is_card(obj)


def test_is_card_no_too_small():
    obj = {"kind": "container", "width": 100000, "height": 100000,
           "fill_hex": "111827", "anomalous": False}
    assert not _is_card(obj)


def test_is_icon_like_picture():
    """Square-ish picture = icon."""
    assert _is_icon_like({"kind": "picture", "width": 500000, "height": 500000})


def test_is_icon_like_picture_wide_bar_no():
    """Wide bar (aspect > 2.5) is NOT an icon, even if it's a picture."""
    assert not _is_icon_like(
        {"kind": "picture", "width": 5000000, "height": 100000},
    )


def test_is_icon_like_emoji_text():
    """Single emoji char with big font in roughly-square bbox = icon."""
    obj = {"kind": "text", "text": "💬",
           "width": 600000, "height": 600000,
           "font_sizes": [48 * 12700]}  # 48pt
    assert _is_icon_like(obj)


def test_is_icon_like_normal_text_no():
    obj = {"kind": "text", "text": "This is a long description",
           "width": 4000000, "height": 200000,
           "font_sizes": [12 * 12700]}
    assert not _is_icon_like(obj)


def test_is_icon_like_short_but_small_font_no():
    """Short text with small font ≠ icon (could be a label)."""
    obj = {"kind": "text", "text": "OK",
           "width": 200000, "height": 200000,
           "font_sizes": [10 * 12700]}
    assert not _is_icon_like(obj)


# ---- icon undersized detection ----

def test_detect_icon_undersized():
    """Tiny icon in big card → flag + resize argv."""
    card = {"shape_id": 1, "kind": "container",
            "left": 0, "top": 0, "width": 5000000, "height": 3000000,
            "fill_hex": "111827", "anomalous": False}
    icon = {"shape_id": 2, "kind": "picture",
            "left": 2400000, "top": 1400000,
            "width": 200000, "height": 200000,
            "anomalous": False}
    slide = {"objects": [card, icon]}
    issues = detect_proportion_issues(slide)
    underflows = [i for i in issues if i["category"] == "icon-undersized"]
    assert len(underflows) == 1
    assert underflows[0]["shape_id"] == 2
    argv = underflows[0]["suggested_argv"]
    assert argv[0] == "resize"
    new_w = int(argv[argv.index("--width") + 1])
    # Target 12% of 15M EMU² area = 1.8M EMU² → side ~1.34M EMU
    assert new_w > 1000000   # significantly bigger than 200000


def test_detect_icon_oversized():
    """Huge icon eating the card → flag + resize argv to shrink."""
    card = {"shape_id": 1, "kind": "container",
            "left": 0, "top": 0, "width": 5000000, "height": 3000000,
            "fill_hex": "111827", "anomalous": False}
    icon = {"shape_id": 2, "kind": "picture",
            "left": 200000, "top": 200000,
            "width": 4000000, "height": 2500000,   # 67% of card
            "anomalous": False}
    slide = {"objects": [card, icon]}
    issues = detect_proportion_issues(slide)
    overflows = [i for i in issues if i["category"] == "icon-oversized"]
    assert len(overflows) == 1
    argv = overflows[0]["suggested_argv"]
    assert argv[0] == "resize"
    new_w = int(argv[argv.index("--width") + 1])
    assert new_w < 4000000   # smaller than current


# ---- card too empty ----

def test_detect_card_too_empty():
    """Card with content occupying < 45% area → too empty."""
    card = {"shape_id": 1, "kind": "container",
            "left": 0, "top": 0, "width": 5000000, "height": 3000000,
            "fill_hex": "111827", "anomalous": False}
    icon = {"shape_id": 2, "kind": "picture",
            "left": 2400000, "top": 1400000,
            "width": 200000, "height": 200000,
            "anomalous": False}
    slide = {"objects": [card, icon]}
    issues = detect_proportion_issues(slide)
    empty = [i for i in issues if i["category"] == "card-too-empty"]
    assert len(empty) == 1
    assert empty[0]["proportion_report"]["empty_pct"] > 55


# ---- icon-text size mismatch ----

def test_detect_icon_text_mismatch():
    """100pt emoji + 9pt body = mismatch."""
    card = {"shape_id": 1, "kind": "container",
            "left": 0, "top": 0, "width": 5000000, "height": 3000000,
            "fill_hex": "111827", "anomalous": False}
    icon = {"shape_id": 2, "kind": "text", "text": "💬",
            "left": 100000, "top": 100000,
            "width": 1000000, "height": 1000000,
            "font_sizes": [100 * 12700],
            "anomalous": False}
    title = {"shape_id": 3, "kind": "text", "text": "Description text body",
             "left": 100000, "top": 1500000,
             "width": 4000000, "height": 300000,
             "font_sizes": [9 * 12700],
             "anomalous": False}
    slide = {"objects": [card, icon, title]}
    issues = detect_proportion_issues(slide)
    mismatch = [i for i in issues if i["category"] == "icon-text-size-mismatch"]
    assert len(mismatch) == 1
    rep = mismatch[0]["proportion_report"]
    assert rep["icon_pt"] > 90
    assert rep["title_pt"] < 12
    assert rep["ratio"] > 8


# ---- top-heavy composition ----

def test_detect_top_heavy():
    """All content in top half → top-heavy."""
    card = {"shape_id": 1, "kind": "container",
            "left": 0, "top": 0, "width": 5000000, "height": 3000000,
            "fill_hex": "111827", "anomalous": False}
    # Icon AND text both in top quarter
    icon = {"shape_id": 2, "kind": "picture",
            "left": 100000, "top": 100000,
            "width": 1000000, "height": 500000,
            "anomalous": False}
    text = {"shape_id": 3, "kind": "text", "text": "Header text long enough",
            "left": 100000, "top": 200000,
            "width": 4000000, "height": 400000,
            "font_sizes": [12 * 12700],
            "anomalous": False}
    slide = {"objects": [card, icon, text]}
    issues = detect_proportion_issues(slide)
    top_heavy = [i for i in issues if i["category"] == "top-heavy-composition"]
    assert len(top_heavy) == 1


# ---- no false positives ----

def test_no_issues_for_balanced_card():
    """A well-proportioned card produces no issues."""
    card = {"shape_id": 1, "kind": "container",
            "left": 0, "top": 0, "width": 5000000, "height": 3000000,
            "fill_hex": "111827", "anomalous": False}
    # Icon at 12% of area
    icon = {"shape_id": 2, "kind": "picture",
            "left": 1900000, "top": 600000,
            "width": 1300000, "height": 1300000,
            "anomalous": False}
    # Text below icon
    text = {"shape_id": 3, "kind": "text", "text": "Description here long enough",
            "left": 200000, "top": 2200000,
            "width": 4600000, "height": 700000,
            "font_sizes": [14 * 12700],
            "anomalous": False}
    slide = {"objects": [card, icon, text]}
    issues = detect_proportion_issues(slide)
    # May still emit info-level "card-too-empty" but the icon shouldn't be
    # flagged as undersized or oversized.
    assert not any(i["category"] in ("icon-undersized", "icon-oversized")
                    for i in issues)
