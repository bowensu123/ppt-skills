"""Smoke tests for _common.py utilities."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from _common import (
    Theme,
    best_text_on,
    contrast_ratio,
    emu_to_inches,
    hex_to_rgb,
    inches_to_emu,
    load_theme,
    normalize_geometry,
    parse_shape_ids,
    rgb_to_hex,
)


def test_hex_round_trip():
    rgb = hex_to_rgb("0F62FE")
    assert rgb_to_hex(rgb) == "0F62FE"


def test_hex_strips_pound():
    assert rgb_to_hex(hex_to_rgb("#FFFFFF")) == "FFFFFF"


def test_hex_invalid():
    with pytest.raises(ValueError):
        hex_to_rgb("XYZ")


def test_emu_inches_round_trip():
    assert inches_to_emu(emu_to_inches(914400)) == 914400


def test_normalize_geometry_flips_negative():
    left, top, w, h, flipped = normalize_geometry(1000, 500, -200, 100)
    assert flipped is True
    assert left == 800 and w == 200
    assert top == 500 and h == 100


def test_normalize_geometry_passes_positive():
    out = normalize_geometry(0, 0, 100, 100)
    assert out == (0, 0, 100, 100, False)


def test_contrast_ratio_white_black():
    assert round(contrast_ratio("FFFFFF", "000000"), 1) == 21.0


def test_best_text_on_dark_picks_white():
    assert best_text_on("0F172A") == "FFFFFF"


def test_best_text_on_light_picks_dark():
    assert best_text_on("FFFFFF") == "161616"


def test_parse_shape_ids():
    assert parse_shape_ids("5,6,7") == [5, 6, 7]
    assert parse_shape_ids("5 6 7") == [5, 6, 7]
    assert parse_shape_ids("12") == [12]


def test_load_theme_clean_tech():
    theme = load_theme()
    assert isinstance(theme, Theme)
    assert theme.name == "clean-tech"
    assert "primary" in theme.palette
    assert theme.font_size_pt("title") == 28


def test_load_theme_invalid_raises(tmp_path: Path):
    bad = tmp_path / "bad.json"
    bad.write_text(json.dumps({"name": "x"}), encoding="utf-8")
    with pytest.raises(ValueError):
        load_theme(bad)
