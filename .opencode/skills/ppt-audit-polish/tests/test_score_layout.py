from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
FIXTURE_SCRIPT = ROOT / "tests" / "fixtures" / "build_fixtures.py"
INSPECT_SCRIPT = ROOT / "scripts" / "inspect_ppt.py"
SCORE_SCRIPT = ROOT / "scripts" / "score_layout.py"


def test_score_layout_flags_overflow_and_style_inconsistency(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    subprocess.run([sys.executable, str(FIXTURE_SCRIPT), "--output-dir", str(fixture_dir)], check=True)

    inspect_json = tmp_path / "slides.json"
    subprocess.run(
        [
            sys.executable,
            str(INSPECT_SCRIPT),
            "--input",
            str(fixture_dir / "alignment-case.pptx"),
            "--output",
            str(inspect_json),
        ],
        check=True,
    )

    findings_json = tmp_path / "findings.json"
    result = subprocess.run(
        [
            sys.executable,
            str(SCORE_SCRIPT),
            "--inspection",
            str(inspect_json),
            "--output",
            str(findings_json),
        ],
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    findings = json.loads(findings_json.read_text(encoding="utf-8"))
    categories = {issue["category"] for issue in findings["slides"][0]["issues"]}
    assert "alignment-inconsistency" in categories
    assert "font-hierarchy-inconsistency" in categories
