from __future__ import annotations

import argparse
import json
from pathlib import Path

from pptx import Presentation


def _shape_kind(shape) -> str:
    if getattr(shape, "has_text_frame", False):
        return "text"
    if shape.shape_type == 13:
        return "picture"
    if shape.shape_type == 3:
        return "chart"
    return "shape"


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


def inspect_presentation(input_path: Path) -> dict:
    prs = Presentation(str(input_path))
    slides: list[dict] = []

    for slide_index, slide in enumerate(prs.slides, start=1):
        objects: list[dict] = []
        for object_index, shape in enumerate(slide.shapes, start=1):
            objects.append(
                {
                    "object_index": object_index,
                    "shape_id": getattr(shape, "shape_id", object_index),
                    "name": getattr(shape, "name", f"Shape {object_index}"),
                    "kind": _shape_kind(shape),
                    "left": int(shape.left),
                    "top": int(shape.top),
                    "width": int(shape.width),
                    "height": int(shape.height),
                    "rotation": float(getattr(shape, "rotation", 0.0) or 0.0),
                    "text": _shape_text(shape),
                    "font_sizes": _font_sizes(shape),
                }
            )

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
