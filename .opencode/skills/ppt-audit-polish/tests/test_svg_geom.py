"""Unit tests for _svg_geom.py — SVG-derived post-render signals."""
from __future__ import annotations

import sys
import textwrap
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

import pytest

from _svg_geom import (
    SVG_UNIT_TO_EMU,
    SvgShape,
    SvgTextRun,
    _iou,
    _normalize_font,
    _path_endpoints,
    detect_font_fallback,
    detect_text_overflow,
    detect_z_order_drift,
    extract_signals,
    match_by_bbox,
    parse_svg,
)


# ---------- helpers ----------

def _make_svg(body: str) -> str:
    """Wrap a body fragment in a minimal LibreOffice-style SVG envelope."""
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<svg version="1.2" viewBox="0 0 33867 19050"
     xmlns="http://www.w3.org/2000/svg">
  <g class="SlideGroup">
    <g id="dummy-slide" class="Slide"/>
    <g id="id1" class="Slide">
      {body}
    </g>
  </g>
</svg>"""


def _custom_shape(svg_id: str, x: int, y: int, w: int, h: int, text_block: str = "") -> str:
    return f"""
      <g class="com.sun.star.drawing.CustomShape">
        <g id="{svg_id}">
          <rect x="{x}" y="{y}" width="{w}" height="{h}"/>
          {text_block}
        </g>
      </g>"""


def _text_block(x: int, y: int, family: str, size_px: int, text: str, weight: str = "400") -> str:
    return f"""
          <text class="SVGTextShape">
            <tspan class="TextParagraph">
              <tspan class="TextPosition" x="{x}" y="{y}">
                <tspan font-family="{family}" font-size="{size_px}px"
                       font-weight="{weight}" textLength="100">{text}</tspan>
              </tspan>
            </tspan>
          </text>"""


def _multiline_text_block(x: int, base_y: int, family: str, size_px: int, lines: list[str]) -> str:
    """Produce N TextPosition lines, simulating wrapped text."""
    out = ['<text class="SVGTextShape"><tspan class="TextParagraph">']
    line_step = int(size_px * 1.2)
    for i, ln in enumerate(lines):
        y = base_y + i * line_step
        out.append(
            f'<tspan class="TextPosition" x="{x}" y="{y}">'
            f'<tspan font-family="{family}" font-size="{size_px}px" font-weight="400" textLength="100">{ln}</tspan>'
            f'</tspan>'
        )
    out.append('</tspan></text>')
    return "".join(out)


# ---------- low-level utilities ----------

def test_iou_no_overlap():
    assert _iou((0, 0, 100, 100), (200, 200, 100, 100)) == 0.0


def test_iou_full_overlap():
    assert _iou((0, 0, 100, 100), (0, 0, 100, 100)) == 1.0


def test_iou_partial():
    # 50x50 overlap area; union = 100*100 + 100*100 - 50*50 = 17500
    assert abs(_iou((0, 0, 100, 100), (50, 50, 100, 100)) - (2500 / 17500)) < 1e-6


def test_normalize_font_strips_embedded_and_fallback():
    assert _normalize_font("Calibri embedded") == "calibri"
    assert _normalize_font("Calibri, sans-serif") == "calibri"
    assert _normalize_font("Microsoft YaHei") == "microsoft yahei"


def test_path_endpoints_extracts_M_and_L():
    pts = list(_path_endpoints("M 100 200 L 300 400 L 500 600 Z"))
    assert pts == [(100.0, 200.0), (300.0, 400.0), (500.0, 600.0)]


# ---------- parse_svg ----------

def test_parse_svg_skips_dummy_slide(tmp_path):
    svg = _make_svg(_custom_shape("id3", 100, 200, 300, 400))
    f = tmp_path / "t.svg"
    f.write_text(svg, encoding="utf-8")
    slides = parse_svg(f)
    assert len(slides) == 1               # dummy filtered out
    assert slides[0].svg_id == "id1"
    assert len(slides[0].shapes) == 1
    sh = slides[0].shapes[0]
    assert sh.svg_id == "id3"
    assert sh.rendered_bbox_svg == (100.0, 200.0, 300.0, 400.0)
    assert sh.rendered_bbox_emu == (
        100 * SVG_UNIT_TO_EMU, 200 * SVG_UNIT_TO_EMU,
        300 * SVG_UNIT_TO_EMU, 400 * SVG_UNIT_TO_EMU,
    )


def test_parse_svg_extracts_text_runs(tmp_path):
    svg = _make_svg(_custom_shape(
        "id3", 0, 0, 1000, 500,
        text_block=_text_block(50, 250, "Microsoft YaHei", 24, "Hello", weight="700"),
    ))
    f = tmp_path / "t.svg"; f.write_text(svg, encoding="utf-8")
    slides = parse_svg(f)
    runs = slides[0].shapes[0].text_runs
    assert len(runs) == 1
    assert runs[0].font_family == "Microsoft YaHei"
    assert runs[0].font_size_px == 24
    assert runs[0].font_weight == "700"
    assert runs[0].text == "Hello"


# ---------- match_by_bbox ----------

def test_match_by_bbox_pairs_closest_iou():
    pptx = [
        {"shape_id": 5, "left": 0, "top": 0, "width": 36000, "height": 36000, "kind": "container"},
        {"shape_id": 6, "left": 100000, "top": 0, "width": 36000, "height": 36000, "kind": "container"},
    ]
    svg = [
        SvgShape("id3", 0, (0, 0, 100, 100), (0, 0, 36000, 36000)),
        SvgShape("id4", 1, (278, 0, 100, 100), (100080, 0, 36000, 36000)),
    ]
    m = match_by_bbox(pptx, svg)
    assert m[5].svg_id == "id3"
    assert m[6].svg_id == "id4"


def test_match_by_bbox_skips_anomalous():
    pptx = [{"shape_id": 7, "left": 0, "top": 0, "width": 1000, "height": 1000,
             "kind": "container", "anomalous": True}]
    svg = [SvgShape("id3", 0, (0, 0, 1000, 1000), (0, 0, 360000, 360000))]
    assert match_by_bbox(pptx, svg) == {}


# ---------- detect_text_overflow ----------

def test_text_overflow_fires_when_wrapped_lines_exceed_frame(tmp_path):
    # 8 wrapped lines × 60px × 1.2 = 576 SVG units = 207360 EMU.
    # Frame is only 30000 EMU tall → 177360 EMU overflow, well past threshold.
    pptx = {"shape_id": 9, "kind": "text", "text": "wrap me",
            "left": 0, "top": 0, "width": 100000, "height": 30000,
            "name": "wrap-shape"}
    svg = SvgShape("id3", 0, (0, 0, 200, 200), (0, 0, 72000, 72000))
    svg.text_runs = [
        SvgTextRun(x=0, y=10 + i * 72, font_family="Calibri", font_size_px=60,
                   font_weight="400", text=f"line{i}", text_length=100)
        for i in range(8)
    ]
    fix = detect_text_overflow(pptx, svg)
    assert fix is not None
    assert fix["wrap_lines"] == 8
    assert fix["overflow_emu"] > 100000


def test_text_overflow_silent_when_fits():
    pptx = {"shape_id": 9, "kind": "text", "text": "fits", "left": 0, "top": 0,
            "width": 100000, "height": 200000, "name": "ok"}
    svg = SvgShape("id3", 0, (0, 0, 200, 100), (0, 0, 72000, 36000))
    svg.text_runs = [
        SvgTextRun(x=0, y=10, font_family="Calibri", font_size_px=20, font_weight="400", text="ok", text_length=100),
    ]
    assert detect_text_overflow(pptx, svg) is None


def test_text_overflow_skips_non_text_shapes():
    pptx = {"shape_id": 9, "kind": "container", "text": "", "left": 0, "top": 0,
            "width": 100000, "height": 30000, "name": "box"}
    svg = SvgShape("id3", 0, (0, 0, 200, 200), (0, 0, 72000, 72000))
    assert detect_text_overflow(pptx, svg) is None


# ---------- detect_font_fallback ----------

def test_font_fallback_fires_when_renderer_swaps_family():
    pptx = {"shape_id": 9, "kind": "text", "text": "hi", "name": "t",
            "font_families": ["Microsoft YaHei"]}
    svg = SvgShape("id3", 0, (0, 0, 100, 100), (0, 0, 36000, 36000))
    svg.text_runs = [SvgTextRun(0, 0, "Calibri, sans-serif", 20, "400", "hi", 100)]
    fix = detect_font_fallback(pptx, svg)
    assert fix is not None
    assert "microsoft yahei" in fix["declared_fonts"]
    assert "calibri" in fix["rendered_fonts"]


def test_font_fallback_silent_when_one_declared_match_present():
    pptx = {"shape_id": 9, "kind": "text", "text": "hi", "name": "t",
            "font_families": ["Microsoft YaHei", "Arial"]}
    svg = SvgShape("id3", 0, (0, 0, 100, 100), (0, 0, 36000, 36000))
    # Renderer used Arial, declared list includes Arial → no fallback
    svg.text_runs = [SvgTextRun(0, 0, "Arial", 20, "400", "hi", 100)]
    assert detect_font_fallback(pptx, svg) is None


def test_font_fallback_silent_when_no_declared_family():
    pptx = {"shape_id": 9, "kind": "text", "text": "hi", "name": "t", "font_families": []}
    svg = SvgShape("id3", 0, (0, 0, 100, 100), (0, 0, 36000, 36000))
    svg.text_runs = [SvgTextRun(0, 0, "Calibri", 20, "400", "hi", 100)]
    assert detect_font_fallback(pptx, svg) is None


# ---------- detect_z_order_drift ----------

def test_z_order_drift_skips_parent_child_stack():
    """Panel(8) contains strip(7) which contains text(6). All three sit at
    the same top-left. PPTX list order is panel→strip→text (text on top).
    SVG renders the same order. NO drift; bbox-contains filters them out."""
    pptx_shapes = [
        {"shape_id": 6, "left": 0, "top": 0, "width": 1000, "height": 1000,
         "kind": "container", "fill_hex": "111827"},
        {"shape_id": 7, "left": 0, "top": 0, "width": 1000, "height": 200,
         "kind": "container", "fill_hex": "3B82F6"},
        {"shape_id": 8, "left": 100, "top": 50, "width": 800, "height": 100,
         "kind": "text", "text": "header"},
    ]
    match = {
        6: SvgShape("id3", 0, (0,0,1000,1000), (0,0,1000,1000)),
        7: SvgShape("id4", 1, (0,0,1000,200),  (0,0,1000,200)),
        8: SvgShape("id5", 2, (100,50,800,100),(100,50,800,100)),
    }
    assert detect_z_order_drift(pptx_shapes, match) == []


def test_z_order_drift_fires_on_real_swap():
    """Two heavily-overlapping cards (neither fully contains the other)
    with declared-vs-rendered z-order swapped. Uses realistic EMU values
    so the parent-child slack tolerance doesn't auto-collapse them."""
    pptx_shapes = [
        # sid=1: text content at (3M, 1M) sized 5M×2M
        {"shape_id": 1, "left": 3000000, "top": 1000000,
         "width": 5000000, "height": 2000000,
         "kind": "text", "text": "important note", "anomalous": False},
        # sid=2: opaque red container, offset enough to NOT contain sid=1
        {"shape_id": 2, "left": 3200000, "top": 1200000,
         "width": 5000000, "height": 2000000,
         "kind": "container", "fill_hex": "FF0000", "anomalous": False},
    ]
    # Declared list order: sid=1 (idx 0) then sid=2 (idx 1) → sid=2 on top.
    # We force SVG ordinals to swap: sid=1 has higher ordinal → SVG says
    # sid=1 on top → real drift.
    match = {
        1: SvgShape("id4", 1, (0,0,1,1), (3000000, 1000000, 5000000, 2000000)),
        2: SvgShape("id3", 0, (0,0,1,1), (3200000, 1200000, 5000000, 2000000)),
    }
    flagged = detect_z_order_drift(pptx_shapes, match)
    assert len(flagged) == 1
    assert flagged[0]["covers"] == 1
    assert flagged[0]["hides"] == 2


def test_z_order_drift_skips_non_overlapping_pairs():
    pptx_shapes = [
        {"shape_id": 1, "left": 0, "top": 0, "width": 100, "height": 100,
         "kind": "container", "fill_hex": "FF0000"},
        {"shape_id": 2, "left": 1000, "top": 1000, "width": 100, "height": 100,
         "kind": "container", "fill_hex": "00FF00"},
    ]
    match = {
        1: SvgShape("id4", 1, (0,0,100,100), (0,0,100,100)),
        2: SvgShape("id3", 0, (1000,1000,100,100), (1000,1000,100,100)),
    }
    # Different bboxes, no overlap → no z-drift regardless of ordering.
    assert detect_z_order_drift(pptx_shapes, match) == []


# ---------- extract_signals (top-level) ----------

def test_extract_signals_end_to_end(tmp_path):
    """One slide, one text shape that overflows, with a fallback font."""
    body = _custom_shape(
        "id3", 0, 0, 100, 30,
        text_block=_multiline_text_block(0, 10, "Calibri, sans-serif", 60, ["a"] * 8),
    )
    svg = _make_svg(body)
    f = tmp_path / "t.svg"; f.write_text(svg, encoding="utf-8")

    inspection = {"slides": [{"objects": [{
        "shape_id": 99,
        "name": "t",
        "kind": "text",
        "text": "wrap",
        "left": 0,
        "top": 0,
        "width": 36000,                   # 100 SVG units
        "height": 10800,                  # 30 SVG units (way too short)
        "font_families": ["Microsoft YaHei"],
        "anomalous": False,
    }]}]}
    signals = extract_signals(f, inspection)
    assert len(signals["text_overflow"]) == 1
    assert signals["text_overflow"][0]["shape_id"] == 99
    assert len(signals["font_fallback"]) == 1
    assert "calibri" in signals["font_fallback"][0]["rendered_fonts"]
    assert signals["match_stats"][0]["matched"] == 1
