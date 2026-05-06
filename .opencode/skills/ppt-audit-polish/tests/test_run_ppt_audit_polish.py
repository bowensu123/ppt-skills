from __future__ import annotations

import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
FIXTURE_SCRIPT = ROOT / "tests" / "fixtures" / "build_fixtures.py"
RUN_SCRIPT = ROOT / "scripts" / "run_ppt_audit_polish.py"


def test_audit_mode_only_writes_report(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    subprocess.run([sys.executable, str(FIXTURE_SCRIPT), "--output-dir", str(fixture_dir)], check=True)

    result = subprocess.run(
        [
            sys.executable,
            str(RUN_SCRIPT),
            "--input",
            str(fixture_dir / "overflow-case.pptx"),
            "--mode",
            "audit",
            "--work-dir",
            str(tmp_path / "audit-work"),
        ],
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    assert (tmp_path / "audit-work" / "overflow-case.audit-report.md").exists()
    assert (tmp_path / "audit-work" / "overflow-case.audit-artifacts").exists()
    assert not (tmp_path / "audit-work" / "overflow-case.polished.pptx").exists()


def test_direct_fix_mode_writes_report_and_polished_copy(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    subprocess.run([sys.executable, str(FIXTURE_SCRIPT), "--output-dir", str(fixture_dir)], check=True)

    result = subprocess.run(
        [
            sys.executable,
            str(RUN_SCRIPT),
            "--input",
            str(fixture_dir / "overflow-case.pptx"),
            "--mode",
            "direct-fix",
            "--work-dir",
            str(tmp_path / "fix-work"),
        ],
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    assert (tmp_path / "fix-work" / "overflow-case.audit-report.md").exists()
    assert (tmp_path / "fix-work" / "overflow-case.audit-artifacts").exists()
    assert (tmp_path / "fix-work" / "overflow-case.polished.pptx").exists()
