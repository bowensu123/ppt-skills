"""Weighted quality score 0-100 from L1 findings + metrics.

Reads a findings.json (output of score_layout.py) and emits a single
deck-level score plus per-slide breakdown. The orchestrator uses this as a
fitness function for greedy iteration; the variant runner uses it to rank.

Default weights are tuned on the AutoDS architecture deck. Override by
passing --weights /path/to/weights.json with any subset of:
  alignment, density, hierarchy, palette, issue_penalty, overflow_penalty
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path


DEFAULT_WEIGHTS = {
    "alignment": 0.25,
    "density": 0.20,
    "hierarchy": 0.20,
    "palette": 0.15,
    "balance": 0.20,         # NEW: visual mass-center balance
    "issue_penalty": 1.5,
    "overflow_penalty": 5.0,
}


def critique(findings: dict, weights: dict | None = None) -> dict:
    w = {**DEFAULT_WEIGHTS, **(weights or {})}
    deck_metrics = findings.get("metrics") or {}
    slides = findings.get("slides", [])

    base = (
        w["alignment"] * deck_metrics.get("alignment_score", 0.0)
        + w["density"] * deck_metrics.get("density_score", 0.0)
        + w["hierarchy"] * deck_metrics.get("hierarchy_score", 0.0)
        + w["palette"] * deck_metrics.get("palette_score", 0.0)
        + w["balance"] * deck_metrics.get("balance_score", 50.0)
    )

    # Penalize unresolved issues and especially overflows.
    issue_count = deck_metrics.get("issue_count", 0)
    overflow_count = deck_metrics.get("overflow_count", 0)
    penalty = w["issue_penalty"] * issue_count + w["overflow_penalty"] * overflow_count

    score = max(0.0, min(100.0, base - penalty))

    per_slide = []
    for slide in slides:
        m = slide.get("metrics", {})
        sb = (
            w["alignment"] * m.get("alignment_score", 0.0)
            + w["density"] * m.get("density_score", 0.0)
            + w["hierarchy"] * m.get("hierarchy_score", 0.0)
            + w["palette"] * m.get("palette_score", 0.0)
            + w["balance"] * m.get("balance_score", 50.0)
        )
        sp = w["issue_penalty"] * m.get("issue_count", 0) + w["overflow_penalty"] * m.get("overflow_count", 0)
        per_slide.append({
            "slide_index": slide["slide_index"],
            "score": round(max(0.0, min(100.0, sb - sp)), 2),
            "metrics": m,
        })

    return {
        "score": round(score, 2),
        "metrics": deck_metrics,
        "weights": w,
        "per_slide": per_slide,
        "verdict": _verdict(score),
    }


def _verdict(score: float) -> str:
    if score >= 85:
        return "excellent"
    if score >= 70:
        return "good"
    if score >= 55:
        return "fair"
    if score >= 40:
        return "needs-work"
    return "poor"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--findings", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--weights", help="optional JSON file with weight overrides")
    args = parser.parse_args()

    findings = json.loads(Path(args.findings).read_text(encoding="utf-8"))
    weights = json.loads(Path(args.weights).read_text(encoding="utf-8")) if args.weights else None
    payload = critique(findings, weights)
    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({"score": payload["score"], "verdict": payload["verdict"]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
