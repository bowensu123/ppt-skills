from __future__ import annotations

import argparse
import json
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE


_TEXT_KIND_CONTAINER = "container"
_TEXT_KIND_TEXT = "text"
_TEXT_KIND_PICTURE = "picture"
_TEXT_KIND_CHART = "chart"
_TEXT_KIND_TABLE = "table"
_TEXT_KIND_CONNECTOR = "connector"
_TEXT_KIND_GROUP = "group"
_TEXT_KIND_OTHER = "shape"


def _shape_kind(shape) -> str:
    shape_type = getattr(shape, "shape_type", None)
    if shape_type == MSO_SHAPE_TYPE.LINE:
        return _TEXT_KIND_CONNECTOR
    if shape_type == MSO_SHAPE_TYPE.GROUP:
        return _TEXT_KIND_GROUP
    if shape_type == MSO_SHAPE_TYPE.PICTURE:
        return _TEXT_KIND_PICTURE
    if shape_type == MSO_SHAPE_TYPE.CHART:
        return _TEXT_KIND_CHART
    if shape_type == MSO_SHAPE_TYPE.TABLE:
        return _TEXT_KIND_TABLE
    has_tf = getattr(shape, "has_text_frame", False)
    if has_tf:
        text = "".join(p.text for p in shape.text_frame.paragraphs).strip()
        return _TEXT_KIND_TEXT if text else _TEXT_KIND_CONTAINER
    return _TEXT_KIND_OTHER


def _shape_text(shape) -> str:
    if not getattr(shape, "has_text_frame", False):
        return ""
    return "\n".join(paragraph.text for paragraph in shape.text_frame.paragraphs).strip()


def _font_sizes(shape) -> list[int]:
    sizes: list[int] = []
    if not getattr(shape, "has_text_frame", False):
        return sizes
    for paragraph in shape.text_frame.paragraphs:
        paragraph_sizes: list[int] = []
        for run in paragraph.runs:
            if run.font.size is not None:
                paragraph_sizes.append(int(run.font.size))
        if paragraph_sizes:
            sizes.extend(paragraph_sizes)
        elif paragraph.font.size is not None:
            sizes.append(int(paragraph.font.size))
    return sizes


def _fill_hex(shape) -> str | None:
    try:
        fill = shape.fill
        if fill.type is None:
            return None
        fore = fill.fore_color
        if fore.type is None:
            return None
        rgb = fore.rgb
        if rgb is None:
            return None
        return str(rgb)
    except (AttributeError, ValueError, KeyError):
        return None


def _line_hex(shape) -> str | None:
    try:
        line = shape.line
        color = line.color
        if color.type is None:
            return None
        rgb = color.rgb
        if rgb is None:
            return None
        return str(rgb)
    except (AttributeError, ValueError, KeyError):
        return None


def _normalize_geometry(left: int, top: int, width: int, height: int) -> tuple[int, int, int, int, bool]:
    flipped = width < 0 or height < 0
    if width < 0:
        left = left + width
        width = -width
    if height < 0:
        top = top + height
        height = -height
    return left, top, width, height, flipped


def _emit_object(slide_index: int, object_index: int, shape) -> dict:
    raw_left = int(shape.left) if shape.left is not None else 0
    raw_top = int(shape.top) if shape.top is not None else 0
    raw_width = int(shape.width) if shape.width is not None else 0
    raw_height = int(shape.height) if shape.height is not None else 0
    left, top, width, height, flipped = _normalize_geometry(raw_left, raw_top, raw_width, raw_height)

    kind = _shape_kind(shape)
    text = _shape_text(shape)
    font_sizes = _font_sizes(shape)
    is_anomalous = width <= 0 or height <= 0

    return {
        "object_index": object_index,
        "shape_id": getattr(shape, "shape_id", object_index),
        "name": getattr(shape, "name", f"Shape {object_index}"),
        "kind": kind,
        "shape_type": int(shape.shape_type) if getattr(shape, "shape_type", None) is not None else None,
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "raw_geometry": {"left": raw_left, "top": raw_top, "width": raw_width, "height": raw_height},
        "geometry_flipped": flipped,
        "anomalous": is_anomalous,
        "rotation": float(getattr(shape, "rotation", 0.0) or 0.0),
        "text": text,
        "font_sizes": font_sizes,
        "fill_hex": _fill_hex(shape),
        "line_hex": _line_hex(shape),
    }


def inspect_presentation(input_path: Path) -> dict:
    prs = Presentation(str(input_path))
    slides: list[dict] = []

    for slide_index, slide in enumerate(prs.slides, start=1):
        objects: list[dict] = []
        for object_index, shape in enumerate(slide.shapes, start=1):
            objects.append(_emit_object(slide_index, object_index, shape))

        slides.append(
            {
                "slide_index": slide_index,
                "width_emu": int(prs.slide_width),
                "height_emu": int(prs.slide_height),
                "objects": objects,
            }
        )

    return {"input": str(input_path), "slides": slides}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    payload = inspect_presentation(Path(args.input))
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
