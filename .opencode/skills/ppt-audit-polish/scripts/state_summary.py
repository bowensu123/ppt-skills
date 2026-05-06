"""One-shot state snapshot for the agent-driven Mode 4 loop.

Replaces a 5-step probe chain (inspect → detect → render → score → critique)
with a single CLI call that returns everything an agent needs to decide its
next mutation:

  * overall self-critique score and verdict
  * top-K issues ranked by severity / impact (with suggested mutate ops)
  * key shape ids by role (title, subtitle, badge, cards)
  * deterministic "what to try next" suggestions (already-bound mutate argv)
  * render image path so the agent can Read it visually
  * (optional) diff vs a previous iteration's deck

Output is compact JSON — agent reads it in one tool call instead of 4-5.

CLI:
  python state_summary.py --in deck.pptx --work-dir state/
  python state_summary.py --in iter-2.pptx --work-dir state/ --diff-from iter-1.pptx
  python state_summary.py --in deck.pptx --work-dir state/ --skip-render   # for speed
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

from _common import ensure_parent, write_json


SCRIPT_DIR = Path(__file__).resolve().parent


SEVERITY_RANK = {"error": 3, "warning": 2, "info": 1}


def _run(script_name: str, *args: str) -> None:
    subprocess.run([sys.executable, str(SCRIPT_DIR / script_name), *args], check=True)


def _suggest_mutate_argv(issue: dict) -> list[str] | None:
    """Map an issue to a ready-to-run mutate argv."""
    fix = issue.get("suggested_fix")
    sid = issue.get("shape_id")
    sids = issue.get("shape_ids")

    # Some detectors pre-bake their argv (e.g., shape-overlap nudges).
    if issue.get("suggested_argv"):
        return list(issue["suggested_argv"])

    if fix == "move-within-slide-bounds" and sid:
        return ["fit-to-slide", "--shape-id", str(sid)]
    if fix == "increase-margin" and sid:
        return ["fit-to-slide", "--shape-id", str(sid), "--pad-emu", "457200"]
    if fix == "align-row-tops" and sids:
        return ["align", "--shape-ids", ",".join(map(str, sids)), "--edge", "top"]
    if fix == "equalize-row-gaps" and sids:
        return ["distribute", "--shape-ids", ",".join(map(str, sids)), "--axis", "horizontal"]
    if fix == "align-column-lefts" and sids:
        return ["align", "--shape-ids", ",".join(map(str, sids)), "--edge", "left"]
    if fix == "repair-peer-cards":
        return ["repair-peer-cards"]
    if fix == "resolve-overlap":
        return None  # already handled by suggested_argv above
    if fix == "set-font-color" and sid:
        return None  # detector pre-bakes the color choice via suggested_argv
    if fix in ("normalize-row-font", "manual-review"):
        return None
    return None


def _key_shapes(roles: dict) -> dict:
    """For each canonical role, return its shape_id (one per role)."""
    out: dict[str, list[int]] = {}
    for slide in roles.get("slides", []):
        for entry in slide.get("shapes", []):
            role = entry.get("role")
            if not role:
                continue
            out.setdefault(role, []).append(entry["shape_id"])
    return out


def _diff_summary(before_inspection: dict, after_inspection: dict) -> dict:
    """Light-weight per-shape diff: what moved, what changed fill, etc."""
    a = {o["shape_id"]: o for o in before_inspection["slides"][0]["objects"]}
    b = {o["shape_id"]: o for o in after_inspection["slides"][0]["objects"]}
    moved = []
    resized = []
    recolored = []
    for sid in a.keys() & b.keys():
        if (a[sid]["left"], a[sid]["top"]) != (b[sid]["left"], b[sid]["top"]):
            moved.append({"shape_id": sid, "name": a[sid]["name"],
                          "from": [a[sid]["left"], a[sid]["top"]],
                          "to": [b[sid]["left"], b[sid]["top"]]})
        if (a[sid]["width"], a[sid]["height"]) != (b[sid]["width"], b[sid]["height"]):
            resized.append({"shape_id": sid, "name": a[sid]["name"],
                            "from": [a[sid]["width"], a[sid]["height"]],
                            "to": [b[sid]["width"], b[sid]["height"]]})
        if a[sid].get("fill_hex") != b[sid].get("fill_hex"):
            recolored.append({"shape_id": sid, "name": a[sid]["name"],
                              "from": a[sid].get("fill_hex"),
                              "to": b[sid].get("fill_hex")})
    return {
        "added": [int(s) for s in b.keys() - a.keys()],
        "removed": [int(s) for s in a.keys() - b.keys()],
        "moved": moved[:8],
        "resized": resized[:8],
        "recolored": recolored[:8],
    }


def summarize(
    input_path: Path,
    work_dir: Path,
    skip_render: bool = False,
    diff_from: Path | None = None,
    top_k: int = 8,
) -> dict:
    work_dir.mkdir(parents=True, exist_ok=True)
    inspection = work_dir / "inspection.json"
    roles = work_dir / "roles.json"
    findings = work_dir / "findings.json"
    critique = work_dir / "critique.json"

    _run("inspect_ppt.py", "--input", str(input_path), "--output", str(inspection))
    _run("detect_roles.py", "--inspection", str(inspection), "--output", str(roles))
    _run("score_layout.py", "--inspection", str(inspection), "--roles", str(roles), "--output", str(findings))
    _run("self_critique.py", "--findings", str(findings), "--output", str(critique))

    insp_data = json.loads(inspection.read_text(encoding="utf-8"))
    roles_data = json.loads(roles.read_text(encoding="utf-8"))
    findings_data = json.loads(findings.read_text(encoding="utf-8"))
    crit_data = json.loads(critique.read_text(encoding="utf-8"))

    # Collect issues, rank, and attach mutate suggestions.
    all_issues: list[dict] = []
    for slide in findings_data.get("slides", []):
        for issue in slide.get("issues", []):
            argv = _suggest_mutate_argv(issue)
            entry = {
                "category": issue["category"],
                "severity": issue["severity"],
                "shape_id": issue.get("shape_id"),
                "shape_ids": issue.get("shape_ids"),
                "message": issue["message"],
                "suggested_fix": issue.get("suggested_fix"),
                "mutate_argv": argv,
                "slide_index": slide["slide_index"],
            }
            all_issues.append(entry)
    all_issues.sort(key=lambda i: (-SEVERITY_RANK.get(i["severity"], 0), i["category"]))

    # Render the slide for the agent to look at.
    render_info: dict = {"status": "skipped" if skip_render else "pending"}
    if not skip_render:
        images_dir = work_dir / "render"
        manifest = work_dir / "render.json"
        try:
            _run("render_slides.py", "--input", str(input_path),
                 "--output-dir", str(images_dir), "--manifest", str(manifest))
            mf = json.loads(manifest.read_text(encoding="utf-8"))
            if mf.get("status") == "rendered" and mf.get("images"):
                render_info = {
                    "status": "ok",
                    "images": mf["images"],
                    "first_slide_png": mf["images"][0],
                }
            else:
                render_info = {"status": "failed", "reason": mf.get("reason", "unknown")}
        except subprocess.CalledProcessError as exc:
            render_info = {"status": "failed", "reason": str(exc)[:200]}

    # Optional: diff vs previous iteration (geometry + scores).
    diff_info: dict | None = None
    score_delta: dict | None = None
    if diff_from is not None and Path(diff_from).exists():
        prev_dir = work_dir / "prev"
        prev_dir.mkdir(parents=True, exist_ok=True)
        prev_inspection_path = prev_dir / "inspection.json"
        prev_roles_path = prev_dir / "roles.json"
        prev_findings_path = prev_dir / "findings.json"
        prev_critique_path = prev_dir / "critique.json"
        try:
            _run("inspect_ppt.py", "--input", str(diff_from), "--output", str(prev_inspection_path))
            _run("detect_roles.py", "--inspection", str(prev_inspection_path), "--output", str(prev_roles_path))
            _run("score_layout.py", "--inspection", str(prev_inspection_path),
                 "--roles", str(prev_roles_path), "--output", str(prev_findings_path))
            _run("self_critique.py", "--findings", str(prev_findings_path), "--output", str(prev_critique_path))

            prev_inspection = json.loads(prev_inspection_path.read_text(encoding="utf-8"))
            prev_critique = json.loads(prev_critique_path.read_text(encoding="utf-8"))
            prev_findings = json.loads(prev_findings_path.read_text(encoding="utf-8"))

            diff_info = _diff_summary(prev_inspection, insp_data)
            diff_info["from"] = str(diff_from)

            score_delta = _compute_score_delta(prev_critique, crit_data, prev_findings, findings_data)
        except subprocess.CalledProcessError as exc:
            diff_info = {"status": "failed-to-inspect-previous", "reason": str(exc)[:200]}

    summary = {
        "input": str(input_path),
        "score": crit_data.get("score"),
        "verdict": crit_data.get("verdict"),
        "metrics": crit_data.get("metrics"),
        "issue_count": len(all_issues),
        "top_issues": all_issues[:top_k],
        "key_shapes": _key_shapes(roles_data),
        "shape_count": len(insp_data["slides"][0]["objects"]) if insp_data.get("slides") else 0,
        "render": render_info,
        "diff": diff_info,
        "score_delta": score_delta,
        "next_step_hints": _next_step_hints(all_issues, crit_data),
        "agent_report": _build_agent_report(crit_data, score_delta, diff_info, render_info),
    }
    write_json(work_dir / "state-summary.json", summary)
    return summary


def _compute_score_delta(prev_crit: dict, cur_crit: dict, prev_find: dict, cur_find: dict) -> dict:
    """Build a per-component score delta between two iterations."""
    prev_m = prev_crit.get("metrics") or {}
    cur_m = cur_crit.get("metrics") or {}
    components = {}
    for key in ("alignment_score", "density_score", "hierarchy_score", "palette_score", "balance_score"):
        before = float(prev_m.get(key, 50.0 if key == "balance_score" else 0.0))
        after = float(cur_m.get(key, 50.0 if key == "balance_score" else 0.0))
        components[key] = {
            "before": round(before, 2),
            "after": round(after, 2),
            "delta": round(after - before, 2),
        }
    prev_score = float(prev_crit.get("score", 0.0))
    cur_score = float(cur_crit.get("score", 0.0))
    delta = cur_score - prev_score
    rec = "adopt" if delta >= 0.5 else ("reject" if delta <= -0.5 else "neutral")
    return {
        "previous_score": round(prev_score, 2),
        "current_score": round(cur_score, 2),
        "delta": round(delta, 2),
        "components": components,
        "issue_count_change": cur_m.get("issue_count", 0) - prev_m.get("issue_count", 0),
        "overflow_count_change": cur_m.get("overflow_count", 0) - prev_m.get("overflow_count", 0),
        "verdict_change": f"{prev_crit.get('verdict')} → {cur_crit.get('verdict')}",
        "decision_recommendation": rec,
        "rationale": (
            "score improved by ≥0.5 → recommend ADOPT"
            if rec == "adopt"
            else "score worsened by ≥0.5 → recommend REJECT"
            if rec == "reject"
            else "score change in noise band → neutral; rely on visual judgment"
        ),
    }


def _build_agent_report(crit: dict, score_delta: dict | None, diff: dict | None, render: dict) -> dict:
    """A ready-to-paste 3-part report for the agent to use in STEP D.

    The agent can quote the `numeric_block` and `geometric_changes` directly,
    then add its own `visual_observation` and `final_decision`.
    """
    metrics = crit.get("metrics") or {}
    if score_delta:
        sd = score_delta["components"]
        def _fmt(key):
            if key not in sd:
                return None
            return f"  {key.replace('_score',''):10s}: {sd[key]['before']} → {sd[key]['after']} (Δ {sd[key]['delta']:+})"
        numeric_lines = [
            f"score: {score_delta['previous_score']} → {score_delta['current_score']} (Δ {score_delta['delta']:+})",
        ]
        for key in ("alignment_score", "density_score", "hierarchy_score", "palette_score", "balance_score"):
            line = _fmt(key)
            if line:
                numeric_lines.append(line)
        numeric_lines.append(f"  issues:    Δ {score_delta['issue_count_change']:+}")
        numeric_lines.append(f"  verdict:   {score_delta['verdict_change']}")
        recommendation = score_delta["decision_recommendation"]
    else:
        numeric_lines = [
            f"score: {crit.get('score', 0.0)} ({crit.get('verdict', 'unknown')})",
            f"  alignment: {metrics.get('alignment_score', 0.0)}",
            f"  density:   {metrics.get('density_score', 0.0)}",
            f"  hierarchy: {metrics.get('hierarchy_score', 0.0)}",
            f"  palette:   {metrics.get('palette_score', 0.0)}",
            f"  balance:   {metrics.get('balance_score', 50.0)}",
            f"  issues:    {metrics.get('issue_count', 0)}",
        ]
        recommendation = "baseline"

    geometric_summary = ""
    if diff:
        moved_n = len(diff.get("moved", []))
        resized_n = len(diff.get("resized", []))
        recolored_n = len(diff.get("recolored", []))
        added_n = len(diff.get("added", []))
        removed_n = len(diff.get("removed", []))
        parts = []
        if moved_n: parts.append(f"{moved_n} moved")
        if resized_n: parts.append(f"{resized_n} resized")
        if recolored_n: parts.append(f"{recolored_n} recolored")
        if added_n: parts.append(f"{added_n} added")
        if removed_n: parts.append(f"{removed_n} removed")
        geometric_summary = ", ".join(parts) if parts else "no shape geometry change"

    return {
        "numeric_block": "\n".join(numeric_lines),
        "geometric_changes": geometric_summary,
        "render_path": render.get("first_slide_png"),
        "decision_recommendation": recommendation,
        "instructions_for_agent": (
            "Quote `numeric_block` verbatim, then read `render_path` with the Read tool, "
            "describe what you SEE changed in one sentence, and finally state ADOPT or REJECT "
            "with a one-sentence justification."
        ),
    }


def _next_step_hints(issues: list[dict], crit: dict) -> list[dict]:
    """Surface 3-5 concrete next-step suggestions the agent can pick from.

    Each hint includes a one-sentence reason and the exact argv to run.
    The agent is free to ignore these and use vision-based judgment.
    """
    hints: list[dict] = []
    seen_argv: set[str] = set()

    # 1. Issue-driven hints: errors first.
    for issue in issues:
        if not issue["mutate_argv"]:
            continue
        key = " ".join(issue["mutate_argv"])
        if key in seen_argv:
            continue
        seen_argv.add(key)
        hints.append({
            "reason": f"{issue['category']}: {issue['message']}",
            "mutate_argv": issue["mutate_argv"],
            "expected_metric": issue["category"].split("-")[0],
        })
        if len(hints) >= 3:
            break

    # 2. Score-targeted hints based on weakest dimension.
    metrics = crit.get("metrics", {})
    if metrics:
        weakest = min(
            ("alignment_score", "density_score", "hierarchy_score", "palette_score"),
            key=lambda k: metrics.get(k, 100),
        )
        weakest_v = metrics.get(weakest, 100)
        if weakest_v < 70:
            recipes = {
                "alignment_score": ("structured", "row_equalize"),
                "density_score":   ("light-touch", "card_borders"),
                "hierarchy_score": ("balanced",   "row_force_font"),
                "palette_score":   ("tinted",     "card_fills"),
            }
            recipe, opt = recipes.get(weakest, (None, None))
            if recipe:
                hints.append({
                    "reason": f"weak {weakest}={weakest_v}; consider running variant '{recipe}' or toggling '{opt}'",
                    "mutate_argv": None,
                    "expected_metric": weakest,
                    "recipe": recipe,
                    "polish_option": opt,
                })

    return hints


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--in", dest="in_path", required=True, type=Path)
    parser.add_argument("--work-dir", required=True, type=Path)
    parser.add_argument("--skip-render", action="store_true")
    parser.add_argument("--diff-from", type=Path, help="previous iteration's pptx for diff")
    parser.add_argument("--top-k", type=int, default=8)
    args = parser.parse_args()

    summary = summarize(
        args.in_path,
        args.work_dir,
        skip_render=args.skip_render,
        diff_from=args.diff_from,
        top_k=args.top_k,
    )
    # Print compact JSON to stdout for the agent.
    sys.stdout.write(json.dumps(summary, ensure_ascii=False, indent=2) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
