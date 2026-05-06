"""End-to-end regenerate: extract content from a deck, apply a template,
render and score.

CLI:
  python regenerate.py --in deck.pptx --work-dir out/ --template horizontal-timeline
  python regenerate.py --in deck.pptx --work-dir out/ --auto      # pick best by item count

Outputs in <work-dir>:
  content.json                     extracted content (from extract_content.py)
  <stem>.regen-<template>.pptx     fresh deck
  state-summary.json               score of the regenerated deck
  render/slide-001.png             rendered preview

Mode 5 entry point. Mode 4 agent decides whether to call this or stay in
polish based on the structural diagnosis.
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent


def _run(script: str, *args: str) -> None:
    subprocess.run([sys.executable, str(SCRIPT_DIR / script), *args], check=True)


def _pick_template(item_count: int) -> str:
    """Choose a sensible default template from the item count."""
    if item_count <= 0:
        return "feature-list"   # falls back gracefully when no items
    if item_count <= 7:
        return "horizontal-timeline"
    return "grid-2x3"


def regenerate(
    input_path: Path,
    work_dir: Path,
    template_name: str | None = None,
    theme_path: Path | None = None,
    auto: bool = False,
) -> dict:
    work_dir.mkdir(parents=True, exist_ok=True)

    # 1. Extract content.
    _run("extract_content.py", "--in", str(input_path), "--work-dir", str(work_dir))
    content = json.loads((work_dir / "content.json").read_text(encoding="utf-8"))

    item_count = len(content.get("items") or [])
    if auto or template_name is None:
        template_name = _pick_template(item_count)

    out_pptx = work_dir / f"{input_path.stem}.regen-{template_name}.pptx"
    apply_args = [
        "--content", str(work_dir / "content.json"),
        "--template", template_name,
        "--out", str(out_pptx),
    ]
    if theme_path is not None:
        apply_args.extend(["--theme", str(theme_path)])
    _run("apply_template.py", *apply_args)

    # 2. Score the regenerated deck.
    _run(
        "state_summary.py",
        "--in", str(out_pptx),
        "--work-dir", str(work_dir / "state"),
        "--diff-from", str(input_path),
    )

    state = json.loads((work_dir / "state" / "state-summary.json").read_text(encoding="utf-8"))

    # Copy final deck to a stable location.
    final = work_dir / f"{input_path.stem}.polished.pptx"
    shutil.copy(out_pptx, final)

    summary = {
        "input": str(input_path),
        "template": template_name,
        "regenerated_pptx": str(out_pptx),
        "final_pptx": str(final),
        "item_count": item_count,
        "score": state.get("score"),
        "verdict": state.get("verdict"),
        "metrics": state.get("metrics"),
        "render": state.get("render"),
    }
    (work_dir / "regenerate-summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8",
    )
    return summary


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--in", dest="in_path", required=True, type=Path)
    parser.add_argument("--work-dir", required=True, type=Path)
    parser.add_argument("--template", help="template name; omit with --auto")
    parser.add_argument("--auto", action="store_true",
                        help="pick template automatically from item count")
    parser.add_argument("--theme", type=Path)
    args = parser.parse_args()

    if not args.template and not args.auto:
        parser.error("specify --template <name> or --auto")

    summary = regenerate(
        input_path=args.in_path,
        work_dir=args.work_dir,
        template_name=args.template,
        theme_path=args.theme,
        auto=args.auto,
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
