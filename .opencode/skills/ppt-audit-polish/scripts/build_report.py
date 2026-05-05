from __future__ import annotations

import argparse
import json
from pathlib import Path


def build_report(findings: dict) -> str:
    lines = [
        "# PPT Audit Report",
        "",
        f"- Input: `{findings['input']}`",
        f"- Total issues: `{findings['summary']['issue_count']}`",
        "",
    ]

    for slide in findings["slides"]:
        lines.append(f"## Slide {slide['slide_index']}")
        lines.append("")
        if not slide["issues"]:
            lines.append("- No issues detected.")
            lines.append("")
            continue
        for issue in slide["issues"]:
            suggestion = issue.get("suggested_fix", "manual-review")
            lines.append(
                f"- `{issue['severity']}` `{issue['category']}`: {issue['message']} (suggested fix: `{suggestion}`)"
            )
        lines.append("")

    if "post_fix_summary" in findings:
        lines.append("## Post-Fix Summary")
        lines.append("")
        lines.append(f"- Remaining issues after repair: `{findings['post_fix_summary']['issue_count']}`")
        lines.append("")

    lines.append("## Applied Actions")
    lines.append("")
    for action in findings.get("applied_actions", []):
        lines.append(f"- Slide {action['slide_index']}: `{action['action']}` on `{action['target']}`")
    if not findings.get("applied_actions"):
        lines.append("- No automatic fixes were applied.")
    lines.append("")

    lines.append("## Skipped Or High-Risk Items")
    lines.append("")
    for item in findings.get("skipped_items", []):
        lines.append(f"- Slide {item['slide_index']}: `{item['target']}` skipped because {item['reason']}")
    if not findings.get("skipped_items"):
        lines.append("- No skipped or high-risk items recorded.")
    lines.append("")

    lines.append("## Dependency Notes")
    lines.append("")
    for note in findings.get("dependency_notes", []):
        lines.append(f"- {note}")
    if not findings.get("dependency_notes"):
        lines.append("- No dependency limitations detected.")
    lines.append("")

    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--findings", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    findings = json.loads(Path(args.findings).read_text(encoding="utf-8"))
    Path(args.output).write_text(build_report(findings), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
