"""Tests for self_critique scoring."""
from __future__ import annotations

from self_critique import _verdict, critique


def _make_findings(score_kwargs):
    """Build a findings dict that score_layout would have produced."""
    metrics = {
        "alignment_score": 100.0,
        "density_score": 100.0,
        "hierarchy_score": 100.0,
        "palette_score": 100.0,
        "balance_score": 100.0,
        "issue_count": 0,
        "overflow_count": 0,
        "slide_count": 1,
        **score_kwargs,
    }
    return {
        "input": "test.pptx",
        "summary": {"issue_count": metrics["issue_count"]},
        "metrics": metrics,
        "slides": [{"slide_index": 1, "issues": [], "metrics": metrics}],
    }


def test_perfect_deck_scores_max():
    res = critique(_make_findings({}))
    assert res["score"] == 100.0
    assert res["verdict"] == "excellent"


def test_overflow_heavily_penalized():
    res = critique(_make_findings({"overflow_count": 3}))
    assert res["score"] < 95.0
    assert res["verdict"] in ("good", "fair", "needs-work", "excellent")


def test_issues_penalize_score():
    perfect = critique(_make_findings({})) ["score"]
    issued = critique(_make_findings({"issue_count": 5})) ["score"]
    assert issued < perfect


def test_low_alignment_drops_score():
    high = critique(_make_findings({"alignment_score": 100.0})) ["score"]
    low = critique(_make_findings({"alignment_score": 0.0})) ["score"]
    assert high - low >= 20.0


def test_verdict_thresholds():
    assert _verdict(90) == "excellent"
    assert _verdict(75) == "good"
    assert _verdict(60) == "fair"
    assert _verdict(45) == "needs-work"
    assert _verdict(20) == "poor"


def test_custom_weights():
    base = critique(_make_findings({"issue_count": 4}))["score"]
    boosted = critique(_make_findings({"issue_count": 4}), weights={"issue_penalty": 5.0})["score"]
    assert boosted < base
