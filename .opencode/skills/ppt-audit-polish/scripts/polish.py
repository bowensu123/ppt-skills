"""One-click polish pipeline for batch use.

The default agent flow (Path A) iterates one mutate at a time with
visual verification — appropriate for high-stakes decks but overkill
for batch jobs. This script chains the four most common ops in one
shot:

  1. repair-grid --nested        - fix 2D grid layout outliers
  2. repair-peer-cards --scope safe
                                  - fix 1D row peer outliers
  3. unify-font                  - enforce Microsoft YaHei (Latin + EA)
  4. polish-business --level N   - smart typography + decoration

Each step writes to a fresh temp file so partial failures don't
corrupt the input. State-summary runs at the start and end so the
final report shows baseline-vs-polished delta.

Constraint preserved across the chain: NEVER changes text content.

CLI:
  python polish.py --in deck.pptx --out polished.pptx
  python polish.py --in deck.pptx --out polished.pptx --level 3
  python polish.py --in deck.pptx --out polished.pptx --skip-repair
  python polish.py --in deck.pptx --out polished.pptx --theme \
      themes/business-warm.json
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent


import os


def _utf8_env() -> dict:
    """Force UTF-8 in subprocesses so non-ASCII slide content (Chinese
    titles, etc.) doesn't trip Windows' default GBK codec when the
    child's stdout/stderr is captured."""
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def _run_op(op_name: str, *args: str, capture: bool = True) -> dict:
    """Invoke `python mutate.py <op> ...` and parse the JSON line emitted."""
    cmd = [sys.executable, str(SCRIPT_DIR / "mutate.py"), op_name, *args]
    result = subprocess.run(
        cmd, check=True, capture_output=capture, text=True,
        encoding="utf-8", errors="replace", env=_utf8_env(),
    )
    if not capture or not result.stdout.strip():
        return {}
    # mutate.py emits one JSON record per op invocation; parse the LAST line
    # so any preceding warnings are ignored.
    last = [ln for ln in result.stdout.strip().splitlines() if ln.startswith("{")]
    if not last:
        return {}
    try:
        return json.loads(last[-1])
    except json.JSONDecodeError:
        return {}


def _run_summary(input_path: Path, work_dir: Path,
                 diff_from: Path | None = None,
                 skip_svg: bool = True) -> dict:
    """Invoke state_summary.py and return the summary dict."""
    args = [
        sys.executable, str(SCRIPT_DIR / "state_summary.py"),
        "--in", str(input_path),
        "--work-dir", str(work_dir),
    ]
    if skip_svg:
        args.append("--skip-svg")
    if diff_from is not None:
        args.extend(["--diff-from", str(diff_from)])
    subprocess.run(
        args, check=True, capture_output=True, text=True,
        encoding="utf-8", errors="replace", env=_utf8_env(),
    )
    summary_path = work_dir / "state-summary.json"
    if not summary_path.exists():
        return {}
    return json.loads(summary_path.read_text(encoding="utf-8"))


def polish_one_click(
    input_path: Path,
    output_path: Path,
    level: int = 2,
    theme: Path | None = None,
    skip_repair: bool = False,
    skip_font: bool = False,
    work_dir: Path | None = None,
    peer_groups: Path | None = None,
    skip_asset_extract: bool = False,
) -> dict:
    """Run the chained pipeline. Returns a summary dict.

    `peer_groups` (optional): path to a JSON the agent wrote categorizing
    shapes into semantic peer groups. When provided, `repair-peers-smart`
    runs INSTEAD of the geometric `repair-grid` + `repair-peer-cards` ops.
    The agent's semantic categorization tends to be more accurate than
    pure geometric clustering for decks where size-similarity ≠ same role.

    `skip_asset_extract`: by default we run `_asset_extract.py` at the
    start so assets are preserved in `<work_dir>/assets/` regardless of
    whether the rest of the pipeline references them. The agent can read
    assets-manifest.json to know what icons/pictures exist on the deck.
    """
    cleanup_temp = work_dir is None
    if work_dir is None:
        work_dir = Path(tempfile.mkdtemp(prefix="polish-"))
    else:
        work_dir.mkdir(parents=True, exist_ok=True)

    artifacts: list[dict] = []
    try:
        # 0. Asset extraction — preserve every picture/icon binary so
        #    the agent has full visibility and downstream tools can
        #    relocate them without losing fidelity.
        if not skip_asset_extract:
            try:
                subprocess.run(
                    [sys.executable, str(SCRIPT_DIR / "_asset_extract.py"),
                     "--in", str(input_path),
                     "--work-dir", str(work_dir)],
                    check=True, capture_output=True, text=True,
                    encoding="utf-8", errors="replace", env=_utf8_env(),
                )
                manifest_path = work_dir / "assets-manifest.json"
                if manifest_path.exists():
                    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
                    artifacts.append({
                        "stage": "asset-extract",
                        "extracted": manifest.get("extracted_count", 0),
                    })
            except subprocess.CalledProcessError:
                # Asset extraction is best-effort; pipeline continues.
                artifacts.append({"stage": "asset-extract", "extracted": 0,
                                  "error": "extraction-failed"})

        # 1. Baseline summary (one render so user can compare).
        baseline = _run_summary(input_path, work_dir / "state-baseline")
        baseline_score = baseline.get("score", 0.0)
        baseline_render = baseline.get("render", {}).get("first_slide_png")

        current = input_path
        # 2. Structural repair: agent-defined peer groups when provided,
        #    geometric clustering as fallback.
        if not skip_repair:
            if peer_groups is not None:
                stage = work_dir / "stage-1-peers-smart.pptx"
                r = _run_op("repair-peers-smart",
                            "--in", str(current), "--out", str(stage),
                            "--groups", str(peer_groups))
                artifacts.append({
                    "stage": "repair-peers-smart",
                    "groups": r.get("groups_processed", 0),
                    "actions": r.get("actions_applied", 0),
                })
                current = stage
            else:
                stage = work_dir / "stage-1-grid.pptx"
                r = _run_op("repair-grid",
                            "--in", str(current), "--out", str(stage),
                            "--nested")
                artifacts.append({"stage": "repair-grid",
                                   "actions": r.get("actions_applied", 0)})
                current = stage

                stage = work_dir / "stage-2-peers.pptx"
                r = _run_op("repair-peer-cards",
                            "--in", str(current), "--out", str(stage),
                            "--scope", "safe")
                artifacts.append({"stage": "repair-peer-cards",
                                   "actions": r.get("actions_applied", 0)})
                current = stage

        # 3. unify-font
        if not skip_font:
            stage = work_dir / "stage-3-font.pptx"
            r = _run_op("unify-font",
                        "--in", str(current), "--out", str(stage))
            artifacts.append({"stage": "unify-font", "actions": r.get("changed", 0)})
            current = stage

        # 4. polish-business
        stage = work_dir / "stage-4-business.pptx"
        polish_args = ["--in", str(current), "--out", str(stage), "--level", str(level)]
        if theme is not None:
            polish_args.extend(["--theme", str(theme)])
        r = _run_op("polish-business", *polish_args)
        artifacts.append({
            "stage": "polish-business",
            "actions": r.get("actions_applied", 0),
            "theme": r.get("theme_source"),
        })
        current = stage

        # 5. refine-contrast — background-aware text color fix. Always
        #    runs after polish-business since polish applies theme colors
        #    that may not contrast against per-shape backgrounds.
        stage = work_dir / "stage-5-contrast.pptx"
        refine_args = ["--in", str(current), "--out", str(stage)]
        if theme is not None:
            refine_args.extend(["--theme", str(theme)])
        try:
            r = _run_op("refine-contrast", *refine_args)
            artifacts.append({
                "stage": "refine-contrast",
                "applied": r.get("applied", 0),
                "skipped": r.get("skipped", 0),
            })
            current = stage
        except subprocess.CalledProcessError:
            artifacts.append({"stage": "refine-contrast", "error": "failed"})

        # 6. refine-proportions — visual composition auditor that catches
        #    icon-too-small / icon-too-big / card-too-empty / icon-text
        #    size mismatch, then auto-applies safe resize fixes.
        stage = work_dir / "stage-6-proportions.pptx"
        try:
            r = _run_op("refine-proportions",
                        "--in", str(current), "--out", str(stage))
            artifacts.append({
                "stage": "refine-proportions",
                "detected": r.get("detected", 0),
                "applied": r.get("applied", 0),
            })
            current = stage
        except subprocess.CalledProcessError:
            artifacts.append({"stage": "refine-proportions", "error": "failed"})

        # 7. Copy final to output and produce final summary with diff.
        shutil.copy(current, output_path)
        final = _run_summary(
            output_path, work_dir / "state-final",
            diff_from=input_path,
        )
        final_score = final.get("score", 0.0)
        final_render = final.get("render", {}).get("first_slide_png")

        return {
            "input": str(input_path),
            "output": str(output_path),
            "level": level,
            "baseline_score": baseline_score,
            "final_score": final_score,
            "delta": round(final_score - baseline_score, 2),
            "baseline_render": baseline_render,
            "final_render": final_render,
            "stages": artifacts,
            "agent_report": final.get("agent_report"),
        }
    finally:
        if cleanup_temp:
            shutil.rmtree(work_dir, ignore_errors=True)


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "One-click polish pipeline: repair-grid + repair-peer-cards + "
            "unify-font + polish-business. Doesn't change content."
        ),
    )
    parser.add_argument("--in", dest="in_path", required=True, type=Path)
    parser.add_argument("--out", dest="out_path", required=True, type=Path)
    parser.add_argument("--level", type=int, choices=[1, 2, 3], default=2,
                        help="polish-business level: 1 subtle, 2 standard (default), 3 rich")
    parser.add_argument("--theme", type=Path, default=None,
                        help="optional theme JSON; auto-pick from content if omitted")
    parser.add_argument("--skip-repair", action="store_true",
                        help="skip repair-grid + repair-peer-cards")
    parser.add_argument("--skip-font", action="store_true",
                        help="skip unify-font")
    parser.add_argument("--peer-groups", type=Path, default=None,
                        help="path to peer-groups.json the agent wrote; "
                             "if provided, repair-peers-smart runs INSTEAD "
                             "of the geometric repair-grid + repair-peer-cards")
    parser.add_argument("--skip-asset-extract", action="store_true",
                        help="skip asset extraction (faster but agent loses "
                             "visibility into which icons/pictures exist)")
    parser.add_argument("--work-dir", type=Path, default=None,
                        help="keep intermediate artifacts here (default: temp + cleanup)")
    args = parser.parse_args()

    summary = polish_one_click(
        input_path=args.in_path,
        output_path=args.out_path,
        level=args.level,
        theme=args.theme,
        skip_repair=args.skip_repair,
        skip_font=args.skip_font,
        work_dir=args.work_dir,
        peer_groups=args.peer_groups,
        skip_asset_extract=args.skip_asset_extract,
    )
    sys.stdout.write(json.dumps(summary, ensure_ascii=False, indent=2) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
