"""End-to-end regenerate: extract content from a deck, apply a template,
render and score.

CLI:
  python regenerate.py --in deck.pptx --work-dir out/ --template horizontal-timeline

The template MUST be picked by the calling agent. The recommended flow:

  1. python extract_content.py --in deck.pptx --work-dir out/
       (or run state_summary.py first to also see the original render)
  2. Read out/content.json AND out/render/slide-001.png with your own eyes.
     Decide which template fits: is the content sequential (steps, stages,
     timeline) → horizontal-timeline; independent parallel features →
     grid-2x3; one hero topic with secondary list → feature-list.
  3. python regenerate.py --in deck.pptx --work-dir out/ --template <chosen>

Item-count heuristics were removed deliberately — a 4-step process and
4 independent benefits both have item_count=4 but want different layouts.
The agent's multimodal judgment is the source of truth.

Outputs in <work-dir>:
  content.json                     extracted content (from extract_content.py)
  <stem>.regen-<template>.pptx     fresh deck
  state-summary.json               score of the regenerated deck
  render/slide-001.png             rendered preview

Mode 5 entry point. Mode 4 agent decides whether to call this or stay in
polish based on the structural diagnosis AND content-fit judgment.
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


def regenerate(
    input_path: Path,
    work_dir: Path,
    template_name: str,
    theme_path: Path | None = None,
    skip_repair: bool = False,
) -> dict:
    work_dir.mkdir(parents=True, exist_ok=True)

    # 0. Pre-repair: structural bugs (oversized peer-card boxes, misplaced
    # children) confuse content extraction because shapes that visually
    # belong to card N may sit inside card N-1's bloated bbox. Run
    # repair-peer-cards into a temp copy first; the user's input is never
    # modified.
    extract_source = input_path
    if not skip_repair:
        repaired = work_dir / "_pre-repair.pptx"
        try:
            subprocess.run(
                [sys.executable, str(SCRIPT_DIR / "mutate.py"),
                 "repair-peer-cards", "--in", str(input_path), "--out", str(repaired)],
                check=True, capture_output=True, text=True,
            )
            extract_source = repaired
        except subprocess.CalledProcessError as exc:
            # Repair failure is non-fatal — fall back to raw input.
            (work_dir / "_pre-repair-error.txt").write_text(
                exc.stderr or str(exc), encoding="utf-8")

    # 1. Extract content (from repaired copy so card boundaries are correct).
    _run("extract_content.py", "--in", str(extract_source), "--work-dir", str(work_dir))
    content = json.loads((work_dir / "content.json").read_text(encoding="utf-8"))

    item_count = len(content.get("items") or [])
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
    parser = argparse.ArgumentParser(
        description=(
            "Regenerate a deck with a chosen template. The template is "
            "agent-decided; run extract_content.py first, look at the "
            "content.json + original render, then pick the template."
        ),
    )
    parser.add_argument("--in", dest="in_path", required=True, type=Path)
    parser.add_argument("--work-dir", required=True, type=Path)
    parser.add_argument(
        "--template", required=True,
        help="template name (e.g., horizontal-timeline / grid-2x3 / "
             "feature-list). Pick based on content judgment, not item count.",
    )
    parser.add_argument("--theme", type=Path)
    parser.add_argument("--skip-repair", action="store_true",
                        help="skip the pre-repair pass (faster but may miss content "
                             "if the input has oversized peer-card outliers)")
    # Friendly error if a caller passes the removed --auto flag.
    parser.add_argument(
        "--auto", action="store_true", help=argparse.SUPPRESS,
    )
    args = parser.parse_args()

    if args.auto:
        parser.error(
            "--auto was removed. Template selection is now agent-driven: "
            "run extract_content.py to produce content.json, look at it + "
            "the original render PNG, decide which template fits the content "
            "(timeline / grid / feature-list), then pass --template <name>."
        )

    summary = regenerate(
        input_path=args.in_path,
        work_dir=args.work_dir,
        template_name=args.template,
        theme_path=args.theme,
        skip_repair=args.skip_repair,
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
