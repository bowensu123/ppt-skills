from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
FIXTURE_SCRIPT = ROOT / "tests" / "fixtures" / "build_fixtures.py"
APPLY_SCRIPT = ROOT / "scripts" / "apply_fixes.py"
INSPECT_SCRIPT = ROOT / "scripts" / "inspect_ppt.py"
SCORE_SCRIPT = ROOT / "scripts" / "score_layout.py"


def test_apply_fixes_reduces_issue_count_for_overflow_case(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    subprocess.run([sys.executable, str(FIXTURE_SCRIPT), "--output-dir", str(fixture_dir)], check=True)

    polished_path = tmp_path / "overflow-case.polished.pptx"
    actions_json = tmp_path / "actions.json"
    result = subprocess.run(
        [
            sys.executable,
            str(APPLY_SCRIPT),
            "--input",
            str(fixture_dir / "overflow-case.pptx"),
            "--output",
            str(polished_path),
            "--actions-output",
            str(actions_json),
        ],
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    assert polished_path.exists()
    assert actions_json.exists()

    original_inspect = tmp_path / "original.json"
    polished_inspect = tmp_path / "polished.json"
    original_findings = tmp_path / "original-findings.json"
    polished_findings = tmp_path / "polished-findings.json"

    subprocess.run([sys.executable, str(INSPECT_SCRIPT), "--input", str(fixture_dir / "overflow-case.pptx"), "--output", str(original_inspect)], check=True)
    subprocess.run([sys.executable, str(INSPECT_SCRIPT), "--input", str(polished_path), "--output", str(polished_inspect)], check=True)
    subprocess.run([sys.executable, str(SCORE_SCRIPT), "--inspection", str(original_inspect), "--output", str(original_findings)], check=True)
    subprocess.run([sys.executable, str(SCORE_SCRIPT), "--inspection", str(polished_inspect), "--output", str(polished_findings)], check=True)

    before = json.loads(original_findings.read_text(encoding="utf-8"))["summary"]["issue_count"]
    after = json.loads(polished_findings.read_text(encoding="utf-8"))["summary"]["issue_count"]
    assert after < before


def test_apply_fixes_normalizes_peer_card_alignment_spacing_and_hierarchy(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    subprocess.run([sys.executable, str(FIXTURE_SCRIPT), "--output-dir", str(fixture_dir)], check=True)

    polished_path = tmp_path / "peer-gap-case.polished.pptx"
    actions_json = tmp_path / "actions.json"
    result = subprocess.run(
        [
            sys.executable,
            str(APPLY_SCRIPT),
            "--input",
            str(fixture_dir / "peer-gap-case.pptx"),
            "--output",
            str(polished_path),
            "--actions-output",
            str(actions_json),
        ],
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr

    polished_inspect = tmp_path / "polished.json"
    polished_findings = tmp_path / "polished-findings.json"
    subprocess.run([sys.executable, str(INSPECT_SCRIPT), "--input", str(polished_path), "--output", str(polished_inspect)], check=True)
    subprocess.run([sys.executable, str(SCORE_SCRIPT), "--inspection", str(polished_inspect), "--output", str(polished_findings)], check=True)

    findings = json.loads(polished_findings.read_text(encoding="utf-8"))
    remaining_categories = {issue["category"] for issue in findings["slides"][0]["issues"]}
    assert "alignment-inconsistency" not in remaining_categories
    assert "spacing-inconsistency" not in remaining_categories
    assert "font-hierarchy-inconsistency" not in remaining_categories

    action_names = {action["action"] for action in json.loads(actions_json.read_text(encoding="utf-8"))["applied_actions"]}
    assert "align-peer-row" in action_names
    assert "normalize-peer-gaps" in action_names
    assert "normalize-font-hierarchy" in action_names
