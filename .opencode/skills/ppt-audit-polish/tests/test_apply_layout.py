"""Unit tests for apply_layout.py — free-form layout renderer."""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from apply_layout import _resolve_ref, render_layout


# ---- _resolve_ref ----

def test_resolve_ref_top_level():
    assert _resolve_ref({"title": "Hello"}, "title") == "Hello"


def test_resolve_ref_nested_dict():
    content = {"meta": {"author": "Alice"}}
    assert _resolve_ref(content, "meta.author") == "Alice"


def test_resolve_ref_list_index():
    content = {"items": [{"name": "a"}, {"name": "b"}, {"name": "c"}]}
    assert _resolve_ref(content, "items.1.name") == "b"


def test_resolve_ref_missing_returns_none():
    assert _resolve_ref({"a": 1}, "b") is None
    assert _resolve_ref({"a": {"b": 2}}, "a.c") is None
    assert _resolve_ref({"items": []}, "items.0.name") is None


def test_resolve_ref_image_dict():
    """An item.image is a dict — ref returns the whole dict, caller pulls .path"""
    content = {"items": [{"image": {"path": "assets/a.png", "asset_id": "a01"}}]}
    img = _resolve_ref(content, "items.0.image")
    assert img == {"path": "assets/a.png", "asset_id": "a01"}


# ---- render_layout end-to-end ----

def test_render_layout_minimal_deck(tmp_path):
    """A layout with one rect + one text renders without error."""
    layout = {
        "elements": [
            {"kind": "rect", "bbox": [0, 0, 12192000, 6858000],
             "fill": "FFFFFF"},
            {"kind": "text", "bbox": [457200, 457200, 11000000, 411480],
             "ref": "title", "size_pt": 24, "color": "161616"},
        ],
    }
    content = {"title": "Test Title"}
    out = tmp_path / "out.pptx"
    result = render_layout(layout, content, out)
    assert out.exists()
    assert result["rendered"] == 2
    assert result["skipped"] == []


def test_render_layout_skips_missing_ref(tmp_path):
    """Element with a ref that doesn't resolve is skipped, not crashed."""
    layout = {
        "elements": [
            {"kind": "text", "bbox": [0, 0, 1000, 1000],
             "ref": "nonexistent", "size_pt": 12},
        ],
    }
    content = {"title": "Hello"}
    out = tmp_path / "out.pptx"
    result = render_layout(layout, content, out)
    assert result["rendered"] == 0
    assert any(s["reason"] == "ref-missing" for s in result["skipped"])


def test_render_layout_with_literal_text(tmp_path):
    """Element with `content` literal string overrides ref."""
    layout = {
        "elements": [
            {"kind": "text", "bbox": [0, 0, 1000, 1000],
             "content": "literal", "ref": "title", "size_pt": 12},
        ],
    }
    content = {"title": "Should Not See"}
    out = tmp_path / "out.pptx"
    result = render_layout(layout, content, out)
    assert result["rendered"] == 1


def test_render_layout_image_missing_file(tmp_path):
    """Image element pointing at nonexistent file is skipped, not crashed."""
    layout = {
        "elements": [
            {"kind": "image", "bbox": [0, 0, 1000000, 1000000],
             "path": "missing.png"},
        ],
    }
    out = tmp_path / "out.pptx"
    result = render_layout(layout, {}, out, assets_base=tmp_path)
    assert any(s["reason"] == "image-not-found" for s in result["skipped"])


def test_render_layout_handles_all_kinds(tmp_path):
    """fill / rect / rounded_rect / circle / line / text all render."""
    layout = {
        "elements": [
            {"kind": "fill", "bbox": [0, 0, 1000, 1000], "color": "111111"},
            {"kind": "rect", "bbox": [100, 100, 200, 200], "fill": "FF0000"},
            {"kind": "rounded_rect", "bbox": [100, 400, 200, 200],
             "fill": "00FF00", "corner_ratio": 0.04},
            {"kind": "circle", "bbox": [400, 400, 200, 200], "fill": "0000FF"},
            {"kind": "line", "bbox": [0, 800, 1000, 800],
             "color": "888888", "width_pt": 1.0},
            {"kind": "text", "bbox": [100, 100, 800, 200],
             "content": "Hello", "size_pt": 14},
        ],
    }
    out = tmp_path / "out.pptx"
    result = render_layout(layout, {}, out)
    # All 6 elements should render.
    assert result["rendered"] == 6
    assert result["skipped"] == []


def test_render_layout_skips_unknown_kind(tmp_path):
    layout = {"elements": [{"kind": "polygon", "bbox": [0, 0, 100, 100]}]}
    out = tmp_path / "out.pptx"
    result = render_layout(layout, {}, out)
    assert any(s["reason"] == "unknown-kind" for s in result["skipped"])


def test_render_layout_background_shorthand(tmp_path):
    """Top-level `background` produces a full-bleed fill rect."""
    layout = {"background": "0F0F0F", "elements": []}
    out = tmp_path / "out.pptx"
    result = render_layout(layout, {}, out)
    # Background isn't counted in `rendered` (it's pre-elements), but the
    # PPTX should save successfully.
    assert out.exists()


def test_render_layout_custom_slide_dims(tmp_path):
    layout = {
        "slide_dims": {"width": 9144000, "height": 6858000},   # 4:3
        "elements": [],
    }
    out = tmp_path / "out.pptx"
    render_layout(layout, {}, out)
    # Re-open and check
    from pptx import Presentation
    prs = Presentation(str(out))
    assert int(prs.slide_width) == 9144000
    assert int(prs.slide_height) == 6858000


def test_render_layout_image_via_ref(tmp_path):
    """Image element with a ref into items[].image renders the dict's path."""
    from PIL import Image
    asset_dir = tmp_path / "assets"; asset_dir.mkdir()
    Image.new("RGB", (8, 8), "white").save(asset_dir / "icon.png", "PNG")

    content = {"items": [{
        "name": "Item 0",
        "image": {"path": "assets/icon.png", "asset_id": "a01"},
    }]}
    layout = {
        "elements": [
            {"kind": "image", "bbox": [100000, 100000, 1000000, 1000000],
             "ref": "items.0.image"},
        ],
    }
    out = tmp_path / "out.pptx"
    result = render_layout(layout, content, out, assets_base=tmp_path)
    assert result["rendered"] == 1, f"skipped: {result['skipped']}"
