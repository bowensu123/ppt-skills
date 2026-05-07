"""Tests for _proportion_audit.py — pure composition descriptor (no
hardcoded thresholds, no auto-fix; agent reads the output + design
principles + render and judges)."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from _proportion_audit import (
    _classify_role,
    _icon_visible_area,
    _is_card,
    _max_font_pt,
    describe_composition,
)


# ---- card identification ----

def test_is_card_yes():
    obj = {"kind": "container", "width": 5000000, "height": 3000000,
           "fill_hex": "111827", "anomalous": False}
    assert _is_card(obj)


def test_is_card_no_too_small():
    obj = {"kind": "container", "width": 100000, "height": 100000,
           "fill_hex": "111827", "anomalous": False}
    assert not _is_card(obj)


def test_is_card_filters_slide_background():
    """Container covering > 80% of slide area is slide-bg, not a card."""
    obj = {"kind": "container", "width": 12192000, "height": 6858000,
           "fill_hex": "0F0F0F", "anomalous": False}
    assert not _is_card(obj, slide_w=12192000, slide_h=6858000)


def test_is_card_no_fill_skipped():
    obj = {"kind": "container", "width": 5000000, "height": 3000000,
           "fill_hex": None, "anomalous": False}
    assert not _is_card(obj)


# ---- role classification ----

def test_classify_role_picture_square_is_icon():
    card = {"left": 0, "top": 0, "width": 5000000, "height": 3000000}
    child = {"kind": "picture", "width": 500000, "height": 500000}
    assert _classify_role(child, card) == "icon"


def test_classify_role_picture_wide_bar_not_icon():
    card = {"left": 0, "top": 0, "width": 5000000, "height": 3000000}
    child = {"kind": "picture", "width": 4000000, "height": 200000}
    assert _classify_role(child, card) == "image"


def test_classify_role_emoji_text_is_icon():
    card = {"left": 0, "top": 0, "width": 5000000, "height": 3000000}
    child = {"kind": "text", "text": "💬",
             "width": 600000, "height": 600000,
             "font_sizes": [80 * 12700]}
    assert _classify_role(child, card) == "icon"


def test_classify_role_title_text():
    card = {"left": 0, "top": 0, "width": 5000000, "height": 3000000}
    child = {"kind": "text", "text": "Section Header",
             "width": 4000000, "height": 400000,
             "font_sizes": [24 * 12700]}
    assert _classify_role(child, card) == "title"


def test_classify_role_body_text():
    card = {"left": 0, "top": 0, "width": 5000000, "height": 3000000}
    child = {"kind": "text", "text": "Description text body line",
             "width": 4000000, "height": 400000,
             "font_sizes": [11 * 12700]}
    assert _classify_role(child, card) == "body"


def test_classify_role_badge_short_small():
    card = {"left": 0, "top": 0, "width": 5000000, "height": 3000000}
    child = {"kind": "text", "text": "01",
             "width": 200000, "height": 200000,
             "font_sizes": [11 * 12700]}
    assert _classify_role(child, card) == "badge"


# ---- icon visible area ----

def test_icon_visible_area_uses_glyph_for_emoji():
    """For text icons, returns (font_pt × 1.2)² — independent of bbox.
    A wide text-frame around a small glyph reports glyph area, not bbox."""
    icon = {"kind": "text", "text": "💬",
            "width": 5000000, "height": 1500000,    # huge frame
            "font_sizes": [80 * 12700]}
    area = _icon_visible_area(icon)
    expected = int(80 * 12700 * 1.2) ** 2
    assert area == expected


def test_icon_visible_area_uses_bbox_for_picture():
    icon = {"kind": "picture", "width": 1000000, "height": 1000000}
    assert _icon_visible_area(icon) == 1_000_000_000_000


# ---- describe_composition end-to-end ----

def test_describe_composition_emits_facts_no_judgments():
    """The descriptor reports geometry, not 'good/bad' verdicts."""
    slide = {
        "slide_index": 1,
        "width_emu": 12192000, "height_emu": 6858000,
        "objects": [
            # Slide-bg (should be filtered out as a "card")
            {"shape_id": 1, "kind": "container",
             "left": 0, "top": 0, "width": 12192000, "height": 6858000,
             "fill_hex": "0F0F0F", "anomalous": False},
            # Real card
            {"shape_id": 2, "kind": "container",
             "left": 457200, "top": 1000000,
             "width": 5400000, "height": 2680000,
             "fill_hex": "1A1A1A", "anomalous": False},
            # Emoji icon inside the card
            {"shape_id": 3, "kind": "text", "text": "💬",
             "left": 2400000, "top": 1800000,
             "width": 600000, "height": 600000,
             "font_sizes": [80 * 12700],
             "anomalous": False},
            # Title above the card
            {"shape_id": 4, "kind": "text", "text": "Section title here",
             "left": 457200, "top": 200000,
             "width": 8000000, "height": 400000,
             "font_sizes": [24 * 12700],
             "anomalous": False},
        ],
    }
    desc = describe_composition(slide)
    # Slide bg filtered out
    assert len(desc["cards"]) == 1
    assert desc["cards"][0]["card_id"] == 2
    # Title detected
    assert len(desc["title_candidates"]) == 1
    assert desc["title_candidates"][0]["shape_id"] == 4
    # Card has 1 child (the emoji)
    card = desc["cards"][0]
    assert card["composition_summary"]["n_children"] == 1
    icon = card["children"][0]
    assert icon["role_hint"] == "icon"
    # The descriptor uses glyph² area for emoji, not bbox.
    # 80pt → glyph_emu = 80 × 12700 × 1.2 = 1219200 → glyph_area ≈ 1.49e12
    # card area = 5400000 × 2680000 = 1.4472e13
    # ratio ≈ 10.3%
    assert 8.0 <= icon["visible_area_pct_of_card"] <= 13.0


def test_describe_composition_includes_alignment_grid():
    """The grid detector clusters lefts/tops where ≥2 shapes share an
    alignment. Need 2+ shapes per gridline for it to be "detected"."""
    slide = {
        "slide_index": 1,
        "width_emu": 12192000, "height_emu": 6858000,
        "objects": [
            # 4 shapes: 2 cards top-aligned, 2 child shapes also aligned
            {"shape_id": 2, "kind": "container",
             "left": 457200, "top": 1000000,
             "width": 5400000, "height": 2680000,
             "fill_hex": "111827", "anomalous": False},
            {"shape_id": 3, "kind": "container",
             "left": 6500000, "top": 1000000,
             "width": 5400000, "height": 2680000,
             "fill_hex": "111827", "anomalous": False},
            # Children inside left card aligned at left=457200
            {"shape_id": 4, "kind": "text", "text": "Hi",
             "left": 457200, "top": 1500000,
             "width": 1000000, "height": 200000,
             "anomalous": False},
            # Children inside right card aligned at left=6500000
            {"shape_id": 5, "kind": "text", "text": "Hi",
             "left": 6500000, "top": 1500000,
             "width": 1000000, "height": 200000,
             "anomalous": False},
        ],
    }
    desc = describe_composition(slide)
    grid = desc["global_alignment"]
    # Now each column has 2 shapes sharing a left → 2 columns detected
    assert len(grid["detected_column_lefts_emu"]) == 2


def test_describe_composition_no_hardcoded_judgments():
    """The descriptor never returns 'icon-undersized' or other verdicts —
    those are agent decisions based on DESIGN_PRINCIPLES.md."""
    slide = {
        "slide_index": 1,
        "width_emu": 12192000, "height_emu": 6858000,
        "objects": [
            {"shape_id": 1, "kind": "container",
             "left": 0, "top": 0, "width": 5400000, "height": 2680000,
             "fill_hex": "111827", "anomalous": False},
            {"shape_id": 2, "kind": "picture",
             "left": 100000, "top": 100000,
             "width": 100000, "height": 100000,
             "anomalous": False},
        ],
    }
    desc = describe_composition(slide)
    # No "issues" array, no severity, no suggested_argv
    assert "issues" not in desc
    assert "actions" not in desc
    # Just facts per child
    assert "visible_area_pct_of_card" in desc["cards"][0]["children"][0]
