"""Render a slide with each shape's shape_id labeled on its bbox.

The agent's eyes are excellent at "this icon is in the wrong card" but
weak at "and the misplaced shape's shape_id is 38". Annotated renders
close that gap: the agent reads the labeled PNG, identifies the
problematic shape by its visible label, and calls mutate.py with the
correct shape_id.

Output:
  <output-dir>/slide-001.png       (vanilla render, same as render_slides)
  <output-dir>/slide-001.annotated.png   (labeled render)

Each shape's bbox is outlined with a thin colored stroke (color by kind:
text=red, container=blue, picture=green, others=gray) and its shape_id
appears in the top-left corner of the bbox in a small label tag.

CLI:
  python annotated_render.py --in deck.pptx --output-dir out/
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


SCRIPT_DIR = Path(__file__).resolve().parent


KIND_COLOR = {
    "text": (220, 38, 38),         # red
    "container": (29, 78, 216),    # blue
    "picture": (22, 163, 74),      # green
    "chart": (124, 58, 237),       # purple
    "table": (217, 119, 6),        # orange
    "connector": (107, 114, 128),  # gray
    "group": (16, 185, 129),       # teal
    "shape": (107, 114, 128),
}


def _run(script: str, *args: str) -> None:
    subprocess.run([sys.executable, str(SCRIPT_DIR / script), *args], check=True)


def _emu_to_px(emu: int, total_emu: int, total_px: int) -> int:
    if total_emu <= 0:
        return 0
    return int(emu * total_px / total_emu)


def annotate(slide_inspection: dict, base_png: Path, output_png: Path) -> dict:
    img = Image.open(base_png).convert("RGBA")
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    try:
        font = ImageFont.truetype("arial.ttf", 14)
        font_small = ImageFont.truetype("arial.ttf", 11)
    except OSError:
        font = ImageFont.load_default()
        font_small = ImageFont.load_default()

    sw_emu = slide_inspection["width_emu"]
    sh_emu = slide_inspection["height_emu"]
    img_w, img_h = img.size

    drawn = 0
    for obj in slide_inspection["objects"]:
        if obj.get("anomalous"):
            continue  # tiny / zero-dim shapes; skip
        kind = obj.get("kind", "shape")
        color = KIND_COLOR.get(kind, KIND_COLOR["shape"])
        x = _emu_to_px(obj["left"], sw_emu, img_w)
        y = _emu_to_px(obj["top"], sh_emu, img_h)
        w = _emu_to_px(obj["width"], sw_emu, img_w)
        h = _emu_to_px(obj["height"], sh_emu, img_h)
        if w <= 0 or h <= 0:
            continue

        # Outline the bbox.
        draw.rectangle([x, y, x + w, y + h], outline=color + (180,), width=2)

        # Label tag in top-left: "ID:42 text"
        label = f"#{obj['shape_id']} {kind[0]}"
        text_w = max(len(label) * 7, 36)
        text_h = 18
        draw.rectangle([x, y, x + text_w, y + text_h], fill=color + (220,))
        draw.text((x + 3, y + 1), label, fill=(255, 255, 255, 255), font=font_small)
        drawn += 1

    annotated = Image.alpha_composite(img, overlay).convert("RGB")
    output_png.parent.mkdir(parents=True, exist_ok=True)
    annotated.save(output_png)
    return {"shapes_labeled": drawn, "output": str(output_png)}


def render_with_annotations(input_path: Path, output_dir: Path) -> dict:
    output_dir.mkdir(parents=True, exist_ok=True)
    inspection = output_dir / "inspection.json"
    base_render = output_dir / "render"
    manifest = output_dir / "render.json"

    _run("inspect_ppt.py", "--input", str(input_path), "--output", str(inspection))
    _run("render_slides.py",
         "--input", str(input_path),
         "--output-dir", str(base_render),
         "--manifest", str(manifest))

    insp_data = json.loads(inspection.read_text(encoding="utf-8"))
    manifest_data = json.loads(manifest.read_text(encoding="utf-8"))
    if manifest_data.get("status") != "rendered":
        return {"status": "render-failed", "reason": manifest_data.get("reason")}

    annotated_paths = []
    for slide_idx, slide in enumerate(insp_data["slides"], start=1):
        png_name = f"slide-{slide_idx:03d}.png"
        base_png = base_render / png_name
        annotated_png = output_dir / f"slide-{slide_idx:03d}.annotated.png"
        if not base_png.exists():
            continue
        result = annotate(slide, base_png, annotated_png)
        annotated_paths.append({
            "slide_index": slide_idx,
            "base": str(base_png),
            "annotated": str(annotated_png),
            "shapes_labeled": result["shapes_labeled"],
        })

    summary = {
        "input": str(input_path),
        "status": "ok",
        "annotated": annotated_paths,
        "color_legend": {
            "text": "#DC2626 (red)",
            "container": "#1D4ED8 (blue)",
            "picture": "#16A34A (green)",
            "chart": "#7C3AED (purple)",
            "table": "#D97706 (orange)",
            "connector": "#6B7280 (gray)",
            "group": "#10B981 (teal)",
        },
    }
    (output_dir / "annotated-summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--in", dest="in_path", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    args = parser.parse_args()

    summary = render_with_annotations(args.in_path, args.output_dir)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
