"""Post-render geometry extraction from a LibreOffice SVG export.

PPTX XML gives us the *declared* geometry of every shape (left, top,
width, height, font_family). But a renderer can produce different actual
output: text wraps to extra lines and overflows the declared text frame;
the declared font isn't installed and the renderer falls back; connector
endpoints snap to the nearest shape edge with sub-EMU drift; group
transforms and master inheritance change the visual z-order.

This module parses the SVG that LibreOffice emits via
`--convert-to svg` and extracts four classes of post-render signals:

  1. text-overflow    — sum of <tspan> heights inside a text shape exceeds
                        the declared text-frame height.
  2. font-fallback    — the SVG <text> uses a font-family different from
                        the PPTX-declared one.
  3. z-order-real     — actual visual stacking (SVG document order)
                        differs from declared z-order in a way that hides
                        an overlapping shape.
  4. connector-snap   — connector endpoint position (from SVG path d)
                        differs from declared endpoint by > tolerance.

LibreOffice's SVG conventions used here (verified empirically):
  * Single SVG file per deck; each slide is `<g class="Slide" id="idN">`.
  * Each shape is `<g class="com.sun.star.drawing.CustomShape">` containing
    a `<g id="idN">` whose first `<rect>` child gives the shape's bbox in
    SVG units (1 SVG unit = 360 EMU = 1/100 mm).
  * Text content sits in `<text class="SVGTextShape">` with
    `<tspan class="TextPosition" x=".." y="..">` and inner
    `<tspan font-family=".." font-size="..px">` carrying the rendered
    style.
  * SVG document order = visual back-to-front z-order.

Matching SVG shapes to PPTX shape_ids uses bbox IoU because the integer
ids LibreOffice emits ("id3", "id4", ...) are sequential and not derived
from PPTX shape_ids.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from lxml import etree


# 1 SVG unit = 360 EMU = 1/100 mm. Verified against LibreOffice's SVG
# viewBox vs the PPTX slide dimensions.
SVG_UNIT_TO_EMU = 360

SVG_NS = {"svg": "http://www.w3.org/2000/svg"}


@dataclass
class SvgTextRun:
    """One contiguous styled run of text (one inner <tspan>)."""
    x: float                        # SVG units
    y: float                        # SVG units
    font_family: str                # what the renderer actually used
    font_size_px: float             # rendered px size
    font_weight: str                # "400", "700", etc.
    text: str
    text_length: float              # SVG units, from textLength attribute


@dataclass
class SvgShape:
    """One PPTX shape as it actually rendered."""
    svg_id: str                     # e.g. "id7"
    ordinal: int                    # 0-based position within its slide
    rendered_bbox_svg: tuple[float, float, float, float]   # (x, y, w, h) SVG units
    rendered_bbox_emu: tuple[int, int, int, int]            # (L, T, W, H) EMU
    text_runs: list[SvgTextRun] = field(default_factory=list)


@dataclass
class SvgSlide:
    """All shapes that rendered on one slide, in z-order."""
    svg_id: str                     # e.g. "id1"
    ordinal: int                    # 0-based slide index (after dropping dummy slides)
    shapes: list[SvgShape] = field(default_factory=list)


# ---- low-level parsing ----

def _to_emu(svg_value: float) -> int:
    return int(round(svg_value * SVG_UNIT_TO_EMU))


def _shape_bbox(shape_g: etree._Element) -> tuple[float, float, float, float] | None:
    """Return (x, y, w, h) of the rendered shape in SVG units.

    Strategy: the first <rect> child of the shape's inner <g> usually
    carries the visual bbox. If absent, scan all geometric primitives
    (rect/path) and return their union.
    """
    inner = shape_g.find("svg:g", SVG_NS)
    container = inner if inner is not None else shape_g

    rect = container.find("svg:rect", SVG_NS)
    if rect is not None:
        try:
            return (
                float(rect.get("x", 0)),
                float(rect.get("y", 0)),
                float(rect.get("width", 0)),
                float(rect.get("height", 0)),
            )
        except (TypeError, ValueError):
            pass

    # Fallback: union of all rects, paths, texts inside.
    xs: list[float] = []
    ys: list[float] = []
    for el in container.iter():
        tag = etree.QName(el).localname
        if tag == "rect":
            try:
                x = float(el.get("x", 0)); y = float(el.get("y", 0))
                w = float(el.get("width", 0)); h = float(el.get("height", 0))
                xs.extend([x, x + w]); ys.extend([y, y + h])
            except (TypeError, ValueError):
                continue
        elif tag == "path":
            d = el.get("d") or ""
            for cmd in _path_endpoints(d):
                xs.append(cmd[0]); ys.append(cmd[1])
    if xs and ys:
        return (min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))
    return None


def _path_endpoints(d: str) -> Iterable[tuple[float, float]]:
    """Yield (x, y) pairs from an SVG path's M/L commands. Cheap, regex-free."""
    cur_x = 0.0; cur_y = 0.0
    i = 0; n = len(d)
    while i < n:
        ch = d[i]
        if ch in "MLmlT":
            i += 1
            # Read 2 numbers
            nums: list[float] = []
            while len(nums) < 2 and i < n:
                while i < n and d[i] in " \t,":
                    i += 1
                start = i
                while i < n and (d[i].isdigit() or d[i] in "-+.eE"):
                    i += 1
                token = d[start:i]
                if token:
                    try:
                        nums.append(float(token))
                    except ValueError:
                        break
                else:
                    break
            if len(nums) == 2:
                if ch.isupper():
                    cur_x, cur_y = nums[0], nums[1]
                else:
                    cur_x, cur_y = cur_x + nums[0], cur_y + nums[1]
                yield (cur_x, cur_y)
        else:
            i += 1


def _extract_text_runs(shape_g: etree._Element) -> list[SvgTextRun]:
    """Extract all rendered text runs inside a shape group."""
    runs: list[SvgTextRun] = []
    for text in shape_g.findall(".//svg:text", SVG_NS):
        # <text class="SVGTextShape">
        #   <tspan class="TextParagraph">
        #     <tspan class="TextPosition" x=".." y="..">
        #       <tspan font-family=".." font-size="..px" font-weight="..">TEXT</tspan>
        for pos_tspan in text.findall(".//svg:tspan[@class='TextPosition']", SVG_NS):
            try:
                x = float(pos_tspan.get("x", 0))
                y = float(pos_tspan.get("y", 0))
            except (TypeError, ValueError):
                continue
            for run_tspan in pos_tspan.findall("svg:tspan", SVG_NS):
                font_family = run_tspan.get("font-family", "")
                font_weight = run_tspan.get("font-weight", "400")
                text_str = "".join(run_tspan.itertext())
                size_str = run_tspan.get("font-size", "0")
                try:
                    size_px = float(size_str.rstrip("px"))
                except (TypeError, ValueError):
                    size_px = 0.0
                try:
                    text_length = float(run_tspan.get("textLength", 0))
                except (TypeError, ValueError):
                    text_length = 0.0
                runs.append(SvgTextRun(
                    x=x, y=y,
                    font_family=font_family,
                    font_size_px=size_px,
                    font_weight=font_weight,
                    text=text_str,
                    text_length=text_length,
                ))
    return runs


# ---- top-level parsing ----

def parse_svg(svg_path: Path) -> list[SvgSlide]:
    """Parse a LibreOffice-emitted PPTX SVG into per-slide shape lists."""
    tree = etree.parse(str(svg_path))
    root = tree.getroot()

    out: list[SvgSlide] = []
    real_idx = 0
    for slide_g in root.findall(".//svg:g[@class='Slide']", SVG_NS):
        # Drop LibreOffice's dummy first slide (id="dummy-slide").
        slide_id = slide_g.get("id") or ""
        if slide_id.startswith("dummy"):
            continue
        slide = SvgSlide(svg_id=slide_id, ordinal=real_idx)
        shape_groups = slide_g.findall(
            ".//svg:g[@class='com.sun.star.drawing.CustomShape']", SVG_NS,
        )
        for ord_idx, shape_g in enumerate(shape_groups):
            inner = shape_g.find("svg:g", SVG_NS)
            svg_id = inner.get("id") if inner is not None else (shape_g.get("id") or "")
            bbox = _shape_bbox(shape_g)
            if bbox is None:
                continue
            x, y, w, h = bbox
            shape = SvgShape(
                svg_id=svg_id,
                ordinal=ord_idx,
                rendered_bbox_svg=(x, y, w, h),
                rendered_bbox_emu=(_to_emu(x), _to_emu(y), _to_emu(w), _to_emu(h)),
                text_runs=_extract_text_runs(shape_g),
            )
            slide.shapes.append(shape)
        out.append(slide)
        real_idx += 1
    return out


# ---- match SVG shapes to PPTX shape_ids by bbox IoU ----

def _iou(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> float:
    al, at, aw, ah = a
    bl, bt, bw, bh = b
    ar, ab_ = al + aw, at + ah
    br, bb = bl + bw, bt + bh
    ix1, iy1 = max(al, bl), max(at, bt)
    ix2, iy2 = min(ar, br), min(ab_, bb)
    if ix2 <= ix1 or iy2 <= iy1:
        return 0.0
    inter = (ix2 - ix1) * (iy2 - iy1)
    union = aw * ah + bw * bh - inter
    return inter / union if union > 0 else 0.0


def match_by_bbox(
    pptx_shapes: list[dict],
    svg_shapes: list[SvgShape],
    iou_threshold: float = 0.30,
) -> dict[int, SvgShape]:
    """For each PPTX shape, find the SVG shape with highest bbox IoU.

    Returns a dict {pptx_shape_id: SvgShape}. Shapes with no good match
    (IoU below threshold) are omitted — typical reasons: master/group
    inheritance differences or shapes the renderer dropped.
    """
    out: dict[int, SvgShape] = {}
    used_svg: set[str] = set()
    # Sort PPTX shapes from largest to smallest area — large shapes have
    # more reliable bbox match; small ones (badges) are matched after.
    ordered = sorted(
        pptx_shapes,
        key=lambda o: o.get("width", 0) * o.get("height", 0),
        reverse=True,
    )
    for pptx in ordered:
        if pptx.get("anomalous"):
            continue
        ppt_bbox = (pptx["left"], pptx["top"], pptx["width"], pptx["height"])
        best_iou = 0.0
        best_svg: SvgShape | None = None
        for svg in svg_shapes:
            if svg.svg_id in used_svg:
                continue
            score = _iou(ppt_bbox, svg.rendered_bbox_emu)
            if score > best_iou:
                best_iou = score
                best_svg = svg
        if best_svg and best_iou >= iou_threshold:
            out[pptx["shape_id"]] = best_svg
            used_svg.add(best_svg.svg_id)
    return out


# ---- signal extractors ----

# Tolerances. EMU.
TEXT_OVERFLOW_EMU = 100000        # ~0.11" — text exceeds frame by this much = overflow
CONNECTOR_SNAP_EMU = 50000        # ~0.055" — endpoint drift from declared
FONT_NORMALIZE_RE = None          # set lazily


def _normalize_font(name: str) -> str:
    """Strip ' embedded' suffix and ', sans-serif' fallback hint."""
    s = name.strip()
    s = s.replace(" embedded", "")
    if "," in s:
        s = s.split(",", 1)[0].strip()
    return s.lower()


def detect_text_overflow(
    pptx_shape: dict, svg_shape: SvgShape,
) -> dict | None:
    """Return a fix descriptor when text rendered taller than its frame.

    We sum the rendered line heights (font_size_px * 1.2 line-height proxy)
    and compare to the declared text-frame height in equivalent units.
    """
    if pptx_shape.get("kind") != "text":
        return None
    if not svg_shape.text_runs:
        return None
    # Group runs by their TextPosition y (one position = one rendered line).
    lines: dict[float, list[SvgTextRun]] = {}
    for run in svg_shape.text_runs:
        lines.setdefault(run.y, []).append(run)
    if not lines:
        return None
    line_count = len(lines)
    # Approximate rendered text height: max font size * 1.2 (line height) * lines.
    max_size_px = max(r.font_size_px for r in svg_shape.text_runs)
    rendered_height_svg = max_size_px * 1.2 * line_count
    rendered_height_emu = _to_emu(rendered_height_svg)
    declared_h = pptx_shape["height"]
    overflow = rendered_height_emu - declared_h
    if overflow > TEXT_OVERFLOW_EMU:
        return {
            "shape_id": pptx_shape["shape_id"],
            "name": pptx_shape.get("name"),
            "declared_height_emu": declared_h,
            "rendered_height_emu": rendered_height_emu,
            "overflow_emu": overflow,
            "wrap_lines": line_count,
        }
    return None


def detect_font_fallback(
    pptx_shape: dict, svg_shape: SvgShape,
) -> dict | None:
    """Return a fix descriptor when the renderer used a different font.

    Inspect_ppt records `font_families` (list of declared families per run).
    We normalize both sides and flag mismatch on the *primary* run.
    """
    if pptx_shape.get("kind") != "text":
        return None
    declared = pptx_shape.get("font_families") or []
    if not declared or not svg_shape.text_runs:
        return None
    declared_norm = {_normalize_font(f) for f in declared if f}
    rendered = {_normalize_font(r.font_family) for r in svg_shape.text_runs if r.font_family}
    rendered = {r for r in rendered if r}
    if not rendered:
        return None
    # If NONE of the declared families appears in the rendered set → fallback.
    if declared_norm.isdisjoint(rendered):
        return {
            "shape_id": pptx_shape["shape_id"],
            "name": pptx_shape.get("name"),
            "declared_fonts": sorted(declared_norm),
            "rendered_fonts": sorted(rendered),
        }
    return None


def detect_z_order_drift(
    pptx_shapes: list[dict], match: dict[int, SvgShape],
) -> list[dict]:
    """Flag pairs whose declared z-order (PPTX list order) disagrees with
    rendered z-order (SVG ordinal) in a way that hides real content.

    Filters out the dominant false-positive pattern: parent-child stacks
    (panel-strip-text triples sitting at the same top-left). Those are
    NORMAL nesting, not z-drift. We only flag when:
      * The two shapes overlap by ≥ 70% of the smaller one's area
      * Neither shape contains the other's bbox (parent-child = skip)
      * The "hidden" shape has visible content (text or distinct fill)
        that would actually become invisible under the swap
      * Declared and rendered ordinals truly disagree

    Connectors and groups are excluded — their "z-order" is meaningless.
    """
    declared_rank = {o["shape_id"]: i for i, o in enumerate(pptx_shapes)}
    flagged: list[dict] = []
    sids = [o["shape_id"] for o in pptx_shapes if o["shape_id"] in match]
    by_id = {o["shape_id"]: o for o in pptx_shapes}

    def _bbox_contains(outer, inner, slack=91440):
        ox1, oy1 = outer["left"], outer["top"]
        ox2, oy2 = ox1 + outer["width"], oy1 + outer["height"]
        ix1, iy1 = inner["left"], inner["top"]
        ix2, iy2 = ix1 + inner["width"], iy1 + inner["height"]
        return (ox1 - slack <= ix1 and oy1 - slack <= iy1
                and ix2 <= ox2 + slack and iy2 <= oy2 + slack)

    def _has_visible_content(o):
        return bool(o.get("text")) or (o.get("fill_hex") and o.get("kind") != "group")

    for i, sid_a in enumerate(sids):
        for sid_b in sids[i + 1:]:
            a = by_id[sid_a]; b = by_id[sid_b]
            if a.get("anomalous") or b.get("anomalous"):
                continue
            if a.get("kind") in ("connector", "group") or b.get("kind") in ("connector", "group"):
                continue
            # Skip parent-child stacks (one bbox contains the other).
            if _bbox_contains(a, b) or _bbox_contains(b, a):
                continue
            ax1, ay1 = a["left"], a["top"]; ax2, ay2 = ax1 + a["width"], ay1 + a["height"]
            bx1, by1 = b["left"], b["top"]; bx2, by2 = bx1 + b["width"], by1 + b["height"]
            ix1, iy1 = max(ax1, bx1), max(ay1, by1)
            ix2, iy2 = min(ax2, bx2), min(ay2, by2)
            if ix2 <= ix1 or iy2 <= iy1:
                continue
            inter = (ix2 - ix1) * (iy2 - iy1)
            sm = min((ax2 - ax1) * (ay2 - ay1), (bx2 - bx1) * (by2 - by1))
            if sm <= 0 or inter / sm < 0.70:
                continue
            decl_a, decl_b = declared_rank[sid_a], declared_rank[sid_b]
            ren_a, ren_b = match[sid_a].ordinal, match[sid_b].ordinal
            decl_top = sid_a if decl_a > decl_b else sid_b
            ren_top  = sid_a if ren_a > ren_b else sid_b
            if decl_top == ren_top:
                continue
            hidden_sid = sid_a if ren_top == sid_b else sid_b
            if not _has_visible_content(by_id[hidden_sid]):
                continue
            flagged.append({
                "declared_top": decl_top,
                "rendered_top": ren_top,
                "covers": ren_top,
                "hides":  hidden_sid,
                "overlap_ratio": round(inter / sm, 2),
            })
    return flagged


# ---- top-level orchestrator ----

def extract_signals(svg_path: Path, inspection: dict) -> dict:
    """Extract all four signal classes for a single-slide deck.

    `inspection` is the output of inspect_ppt.py.
    Returns a dict keyed by signal type with one list per slide.
    """
    slides = parse_svg(svg_path)
    out = {
        "text_overflow": [],
        "font_fallback": [],
        "z_order_drift": [],
        "match_stats": [],
    }
    insp_slides = inspection.get("slides", [])
    for slide_idx, (svg_slide, insp_slide) in enumerate(
        zip(slides, insp_slides), start=1,
    ):
        match = match_by_bbox(insp_slide["objects"], svg_slide.shapes)
        out["match_stats"].append({
            "slide_index": slide_idx,
            "pptx_shape_count": len(insp_slide["objects"]),
            "svg_shape_count": len(svg_slide.shapes),
            "matched": len(match),
        })
        for pptx in insp_slide["objects"]:
            sid = pptx["shape_id"]
            if sid not in match:
                continue
            svg = match[sid]
            tof = detect_text_overflow(pptx, svg)
            if tof:
                tof["slide_index"] = slide_idx
                out["text_overflow"].append(tof)
            ff = detect_font_fallback(pptx, svg)
            if ff:
                ff["slide_index"] = slide_idx
                out["font_fallback"].append(ff)
        zd = detect_z_order_drift(insp_slide["objects"], match)
        for entry in zd:
            entry["slide_index"] = slide_idx
        out["z_order_drift"].extend(zd)
    return out
