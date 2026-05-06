from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
REPORT_SCRIPT = ROOT / "scripts" / "build_report.py"


def test_build_report_writes_slide_grouped_markdown(tmp_path: Path) -> None:
    findings_json = tmp_path / "findings.json"
    findings_json.write_text(
        json.dumps(
            {
                "input": "deck.pptx",
                "summary": {"issue_count": 2},
                "applied_actions": [{"slide_index": 1, "action": "move-within-slide-bounds", "target": "Title 1"}],
                "skipped_items": [{"slide_index": 1, "reason": "chart requires manual review", "target": "Revenue chart"}],
                "dependency_notes": ["Rendered slide validation skipped: soffice-not-found"],
                "slides": [
                    {
                        "slide_index": 1,
                        "issues": [
                            {"category": "boundary-overflow", "severity": "error", "message": "Title exceeds boundary"},
                            {"category": "alignment-inconsistency", "severity": "warning", "message": "Cards are misaligned"},
                        ],
                    }
                ],
            }
        ),
        encoding="utf-8",
    )

    report_path = tmp_path / "deck.audit-report.md"
    result = subprocess.run(
        [sys.executable, str(REPORT_SCRIPT), "--findings", str(findings_json), "--output", str(report_path)],
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    text = report_path.read_text(encoding="utf-8")
    assert "# PPT Audit Report" in text
    assert "## Slide 1" in text
    assert "boundary-overflow" in text
    assert "## Applied Actions" in text
    assert "## Skipped Or High-Risk Items" in text
    assert "## Dependency Notes" in text
