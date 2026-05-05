from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
FIXTURE_SCRIPT = ROOT / "tests" / "fixtures" / "build_fixtures.py"
RENDER_SCRIPT = ROOT / "scripts" / "render_slides.py"


def test_render_slides_reports_skipped_when_soffice_missing(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    subprocess.run([sys.executable, str(FIXTURE_SCRIPT), "--output-dir", str(fixture_dir)], check=True)

    manifest_path = tmp_path / "render.json"
    result = subprocess.run(
        [
            sys.executable,
            str(RENDER_SCRIPT),
            "--input",
            str(fixture_dir / "overflow-case.pptx"),
            "--output-dir",
            str(tmp_path / "images"),
            "--manifest",
            str(manifest_path),
        ],
        env={"PATH": ""},
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert payload["status"] == "skipped"
    assert payload["reason"] == "soffice-not-found"


def test_render_slides_outputs_pngs_when_soffice_exists(tmp_path: Path) -> None:
    if shutil.which("soffice") is None:
        pytest.skip("soffice not installed")

    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    subprocess.run([sys.executable, str(FIXTURE_SCRIPT), "--output-dir", str(fixture_dir)], check=True)

    manifest_path = tmp_path / "render.json"
    output_dir = tmp_path / "images"
    result = subprocess.run(
        [
            sys.executable,
            str(RENDER_SCRIPT),
            "--input",
            str(fixture_dir / "overflow-case.pptx"),
            "--output-dir",
            str(output_dir),
            "--manifest",
            str(manifest_path),
        ],
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert payload["status"] == "rendered"
    assert len(list(output_dir.glob("slide-*.png"))) >= 1
