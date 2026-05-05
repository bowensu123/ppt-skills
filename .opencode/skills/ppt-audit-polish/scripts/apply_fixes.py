from __future__ import annotations

import argparse
import json
from pathlib import Path
from statistics import median

from pptx import Presentation


def _record(action_log: list[dict], slide_index: int, action: str, target: str) -> None:
    action_log.append({"slide_index": slide_index, "action": action, "target": target})


def _shape_font_sizes(shape) -> list[int]:
    if not getattr(shape, "has_text_frame", False):
        return []

    sizes: list[int] = []
    for paragraph in shape.text_frame.paragraphs:
        paragraph_sizes = [int(run.font.size) for run in paragraph.runs if run.font.size is not None]
        if paragraph_sizes:
            sizes.extend(paragraph_sizes)
        elif paragraph.font.size is not None:
            sizes.append(int(paragraph.font.size))
    return sizes


def _fix_shape(
    shape,
    slide_index: int,
    slide_width: int,
    target_left: int | None,
    target_font_size: int | None,
    action_log: list[dict],
    skipped_items: list[dict],
) -> bool:
    if shape.shape_type == 3:
        skipped_items.append(
            {
                "slide_index": slide_index,
                "target": getattr(shape, "name", "Chart"),
                "reason": "chart requires manual review",
            }
        )
        return False

    changed = False
    right = int(shape.left) + int(shape.width)
    overflow = right - slide_width
    if overflow > 0:
        shape.left = max(0, int(shape.left) - overflow)
        changed = True
        _record(action_log, slide_index, "move-within-slide-bounds", getattr(shape, "name", "Shape"))

    if (
        target_left is not None
        and getattr(shape, "has_text_frame", False)
        and shape.text_frame.text.strip()
        and abs(int(shape.left) - target_left) > 120000
    ):
        shape.left = target_left
        changed = True
        _record(action_log, slide_index, "align-left-anchors", getattr(shape, "name", "Text"))

    if getattr(shape, "has_text_frame", False):
        for paragraph in shape.text_frame.paragraphs:
            paragraph_changed = False
            for run in paragraph.runs:
                if (
                    run.font.size is not None
                    and target_font_size is not None
                    and abs(int(run.font.size) - target_font_size) > 30000
                ):
                    run.font.size = target_font_size
                    changed = True
                    paragraph_changed = True
            if (
                not paragraph_changed
                and target_font_size is not None
                and paragraph.font.size is not None
                and abs(int(paragraph.font.size) - target_font_size) > 30000
            ):
                paragraph.font.size = target_font_size
                changed = True
                paragraph_changed = True
            if paragraph_changed:
                _record(action_log, slide_index, "normalize-font-hierarchy", getattr(shape, "name", "Text"))

    return changed


def apply_fixes(input_path: Path, output_path: Path, actions_output: Path | None) -> None:
    prs = Presentation(str(input_path))
    action_log: list[dict] = []
    skipped_items: list[dict] = []

    for slide_index, slide in enumerate(prs.slides, start=1):
        text_shapes = [
            shape
            for shape in slide.shapes
            if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip()
        ]
        target_left = min((int(shape.left) for shape in text_shapes), default=None)

        font_sizes = [size for shape in text_shapes for size in _shape_font_sizes(shape)]
        target_font_size = int(median(font_sizes)) if font_sizes else None

        for shape in slide.shapes:
            _fix_shape(shape, slide_index, int(prs.slide_width), target_left, target_font_size, action_log, skipped_items)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    if actions_output is not None:
        actions_output.parent.mkdir(parents=True, exist_ok=True)
        actions_output.write_text(
            json.dumps({"applied_actions": action_log, "skipped_items": skipped_items}, indent=2),
            encoding="utf-8",
        )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--actions-output")
    args = parser.parse_args()

    actions_output = Path(args.actions_output) if args.actions_output else None
    apply_fixes(Path(args.input), Path(args.output), actions_output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
