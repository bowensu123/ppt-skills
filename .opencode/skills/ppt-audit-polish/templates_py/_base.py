"""Shared helpers used by every template renderer.

Templates are pure Python modules with a ``render(prs, content, theme)``
function. They build a fresh slide using python-pptx primitives. These
helpers wrap the verbose pptx API.
"""
from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt


def _hex(s: str) -> RGBColor:
    s = s.lstrip("#")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def add_rect(slide, left, top, width, height, fill=None, line=None, line_pt=None,
             corner_ratio: float | None = None, shape=MSO_SHAPE.RECTANGLE) -> object:
    sp = slide.shapes.add_shape(shape, Emu(left), Emu(top), Emu(width), Emu(height))
    if fill is not None:
        sp.fill.solid(); sp.fill.fore_color.rgb = _hex(fill)
    elif fill is False:
        sp.fill.background()
    if line == "none":
        try:
            sp.line.fill.background()
        except (AttributeError, ValueError):
            pass
    elif line is not None:
        sp.line.color.rgb = _hex(line)
        if line_pt is not None:
            sp.line.width = Pt(line_pt)
    if corner_ratio is not None:
        try:
            sp.adjustments[0] = corner_ratio
        except (IndexError, AttributeError):
            pass
    return sp


def add_text(slide, left, top, width, height, text, *,
             size_pt: float = 11, bold: bool = False, italic: bool = False,
             color: str = "393939", family: str | None = None,
             align: str = "left", v_align: str = "top",
             fill: str | None = None, padding_emu: int | None = None) -> object:
    box = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(width), Emu(height))
    tf = box.text_frame
    tf.text = ""
    tf.word_wrap = True
    if padding_emu is not None:
        tf.margin_left = Emu(padding_emu); tf.margin_right = Emu(padding_emu)
        tf.margin_top = Emu(padding_emu); tf.margin_bottom = Emu(padding_emu)
    if v_align == "middle":
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    elif v_align == "bottom":
        tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    elif v_align == "top":
        tf.vertical_anchor = MSO_ANCHOR.TOP

    if fill is not None:
        box.fill.solid(); box.fill.fore_color.rgb = _hex(fill)
    try:
        box.line.fill.background()
    except (AttributeError, ValueError):
        pass

    para = tf.paragraphs[0]
    align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}
    para.alignment = align_map.get(align, PP_ALIGN.LEFT)

    run = para.add_run()
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = _hex(color)
    if family:
        run.font.name = family
    return box


def add_circle(slide, cx, cy, radius, fill, line="none") -> object:
    return add_rect(slide, cx - radius, cy - radius, radius * 2, radius * 2,
                    fill=fill, line=line, shape=MSO_SHAPE.OVAL)


def add_rounded_rect(slide, left, top, width, height, *, fill=None, line=None, line_pt=None, corner_ratio=0.06):
    return add_rect(slide, left, top, width, height,
                    fill=fill, line=line, line_pt=line_pt,
                    shape=MSO_SHAPE.ROUNDED_RECTANGLE, corner_ratio=corner_ratio)


def horizontal_distribute(slide_w: int, n: int, margin: int) -> list[int]:
    """Return left-edge x positions for n items evenly spaced across the slide.

    Each item's slot width = (slide_w - 2*margin) / n.
    Returns the left edge of each slot (caller decides item internal offset).
    """
    if n <= 0:
        return []
    available = slide_w - 2 * margin
    slot_w = available // n
    return [margin + i * slot_w for i in range(n)]


def truncate(text: str, max_chars: int) -> str:
    return text if len(text) <= max_chars else text[: max_chars - 1] + "…"


def wipe_slides(prs) -> None:
    """Remove every existing slide so a template can write a fresh first slide."""
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        sld_id_lst.remove(sld_id)


def add_blank_slide(prs):
    """Add a single blank slide using the layout-6 (blank) of the master."""
    blank_layout = None
    for layout in prs.slide_layouts:
        if (layout.name or "").lower().startswith("blank"):
            blank_layout = layout
            break
    if blank_layout is None:
        blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    return prs.slides.add_slide(blank_layout)


def add_image_from_path(slide, image_path, left, top, width, height) -> object | None:
    """Drop a binary image into the slide at the given EMU coordinates.

    Returns the picture shape on success, None on failure (missing file,
    unsupported format). Used by templates to render `items[i].image`
    attributions the agent set on the content JSON.

    The image is fit-stretched into the box; preserve-aspect should be
    handled by the caller (compute width/height to match the original
    aspect ratio if you care about it).
    """
    from pathlib import Path
    p = Path(image_path)
    if not p.exists():
        return None
    try:
        return slide.shapes.add_picture(
            str(p), Emu(left), Emu(top), Emu(width), Emu(height),
        )
    except (ValueError, OSError, KeyError):
        return None


def aspect_fit_box(content_w: int, content_h: int,
                    box_w: int, box_h: int) -> tuple[int, int, int, int]:
    """Center an image of (content_w, content_h) inside (box_w, box_h),
    preserving aspect. Returns (offset_x, offset_y, fit_w, fit_h) in
    the same units as inputs.
    """
    if content_w <= 0 or content_h <= 0 or box_w <= 0 or box_h <= 0:
        return (0, 0, box_w, box_h)
    src_ratio = content_w / content_h
    box_ratio = box_w / box_h
    if src_ratio > box_ratio:
        # Width-bound
        fit_w = box_w
        fit_h = int(box_w / src_ratio)
    else:
        # Height-bound
        fit_h = box_h
        fit_w = int(box_h * src_ratio)
    return ((box_w - fit_w) // 2, (box_h - fit_h) // 2, fit_w, fit_h)


def add_rich_text(slide, left, top, width, height, runs: list[dict], *,
                  default_size_pt: float = 11,
                  default_color: str = "393939",
                  default_family: str | None = None,
                  align: str = "left", v_align: str = "top",
                  fill: str | None = None) -> object:
    """Render a text frame with mixed-format runs.

    Each run is a dict with keys:
      text, font_family, size_pt, bold, italic, color_hex,
      paragraph_index (optional), run_index (optional)

    Runs sharing the same paragraph_index land in the same paragraph.
    Missing format fields fall back to the `default_*` parameters. This
    preserves bold/italic/colored words in titles or descriptions that
    Path B would otherwise flatten to plain text.
    """
    box = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(width), Emu(height))
    tf = box.text_frame
    tf.text = ""
    tf.word_wrap = True
    if v_align == "middle":
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    elif v_align == "bottom":
        tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    else:
        tf.vertical_anchor = MSO_ANCHOR.TOP

    if fill is not None:
        box.fill.solid(); box.fill.fore_color.rgb = _hex(fill)
    try:
        box.line.fill.background()
    except (AttributeError, ValueError):
        pass

    align_map = {
        "left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY,
    }
    pp_align = align_map.get(align, PP_ALIGN.LEFT)

    # Group runs by paragraph_index. Default to a single paragraph.
    by_para: dict[int, list[dict]] = {}
    for run in runs:
        p_idx = int(run.get("paragraph_index", 0) or 0)
        by_para.setdefault(p_idx, []).append(run)
    if not by_para:
        return box

    para_indices = sorted(by_para.keys())
    for i, p_idx in enumerate(para_indices):
        if i == 0:
            para = tf.paragraphs[0]
        else:
            para = tf.add_paragraph()
        para.alignment = pp_align
        for run_data in by_para[p_idx]:
            r = para.add_run()
            r.text = run_data.get("text", "")
            r.font.size = Pt(float(run_data.get("size_pt") or default_size_pt))
            if run_data.get("bold") is not None:
                r.font.bold = bool(run_data["bold"])
            if run_data.get("italic") is not None:
                r.font.italic = bool(run_data["italic"])
            color = run_data.get("color_hex") or default_color
            if color:
                r.font.color.rgb = _hex(color)
            family = run_data.get("font_family") or default_family
            if family:
                r.font.name = family
    return box
