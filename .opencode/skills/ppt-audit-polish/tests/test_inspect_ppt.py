from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
FIXTURE_SCRIPT = ROOT / "tests" / "fixtures" / "build_fixtures.py"
INSPECT_SCRIPT = ROOT / "scripts" / "inspect_ppt.py"


def test_inspect_ppt_extracts_slide_geometry(tmp_path: Path) -> None:
    fixture_dir = tmp_path / "fixtures"
    fixture_dir.mkdir()
    subprocess.run([sys.executable, str(FIXTURE_SCRIPT), "--output-dir", str(fixture_dir)], check=True)

    output_json = tmp_path / "slides.json"
    result = subprocess.run(
        [
            sys.executable,
            str(INSPECT_SCRIPT),
            "--input",
            str(fixture_dir / "overflow-case.pptx"),
            "--output",
            str(output_json),
        ],
        capture_output=True,
        text=True,
    )

    assert result.returncode == 0, result.stderr
    payload = json.loads(output_json.read_text(encoding="utf-8"))
    assert payload["input"].endswith("overflow-case.pptx")
    assert payload["slides"][0]["width_emu"] > 0
    overflow_headline = next(obj for obj in payload["slides"][0]["objects"] if obj["text"] == "Overflow headline")
    assert overflow_headline["font_sizes"] == [355600]
