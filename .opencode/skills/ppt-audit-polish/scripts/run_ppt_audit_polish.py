from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent


def _run(script_name: str, *args: str) -> None:
    subprocess.run([sys.executable, str(SCRIPT_DIR / script_name), *args], check=True)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--mode", choices=("audit", "direct-fix"), default="audit")
    parser.add_argument("--work-dir", required=True)
    args = parser.parse_args()

    input_path = Path(args.input)
    work_dir = Path(args.work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)
    stem = input_path.stem

    artifacts_dir = work_dir / f"{stem}.audit-artifacts"
    inspection_path = artifacts_dir / "inspection.json"
    render_manifest = artifacts_dir / "render.json"
    findings_path = artifacts_dir / "findings.json"
    report_payload = artifacts_dir / "report-payload.json"
    report_path = work_dir / f"{stem}.audit-report.md"
    images_dir = artifacts_dir / "images"

    _run("inspect_ppt.py", "--input", str(input_path), "--output", str(inspection_path))
    _run("render_slides.py", "--input", str(input_path), "--output-dir", str(images_dir), "--manifest", str(render_manifest))
    _run("score_layout.py", "--inspection", str(inspection_path), "--output", str(findings_path))
    report_data = json.loads(findings_path.read_text(encoding="utf-8"))
    render_data = json.loads(render_manifest.read_text(encoding="utf-8"))
    report_data["dependency_notes"] = []
    if render_data["status"] != "rendered":
        report_data["dependency_notes"].append("Rendered slide validation skipped: soffice-not-found")
    report_data["applied_actions"] = []
    report_data["skipped_items"] = []

    if args.mode == "audit":
        report_payload.write_text(json.dumps(report_data, indent=2), encoding="utf-8")
        _run("build_report.py", "--findings", str(report_payload), "--output", str(report_path))
        print(json.dumps({"mode": "audit", "report": str(report_path)}))
        return 0

    polished_path = work_dir / f"{stem}.polished.pptx"
    actions_path = artifacts_dir / "actions.json"
    _run("apply_fixes.py", "--input", str(input_path), "--output", str(polished_path), "--actions-output", str(actions_path))
    action_data = json.loads(actions_path.read_text(encoding="utf-8"))
    report_data["applied_actions"] = action_data["applied_actions"]
    report_data["skipped_items"] = action_data["skipped_items"]

    polished_inspection = artifacts_dir / "polished-inspection.json"
    polished_findings = artifacts_dir / "polished-findings.json"
    _run("inspect_ppt.py", "--input", str(polished_path), "--output", str(polished_inspection))
    _run("score_layout.py", "--inspection", str(polished_inspection), "--output", str(polished_findings))
    report_data["post_fix_summary"] = json.loads(polished_findings.read_text(encoding="utf-8"))["summary"]
    report_payload.write_text(json.dumps(report_data, indent=2), encoding="utf-8")
    _run("build_report.py", "--findings", str(report_payload), "--output", str(report_path))

    print(json.dumps({"mode": "direct-fix", "report": str(report_path), "polished": str(polished_path)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
