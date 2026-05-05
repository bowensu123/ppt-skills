from __future__ import annotations

import argparse
import json
from pathlib import Path


ALIGN_TOLERANCE = 120000
EDGE_TOLERANCE = 80000
FONT_TOLERANCE = 30000
SPACING_TOLERANCE = 140000
DENSITY_THRESHOLD = 0.55


def _gap_values(text_objects: list[dict]) -> list[int]:
    ordered = sorted(text_objects, key=lambda obj: obj["left"])
    gaps: list[int] = []
    for previous, current in zip(ordered, ordered[1:]):
        gaps.append(current["left"] - (previous["left"] + previous["width"]))
    return gaps


def _collect_issues(slide: dict) -> list[dict]:
    issues: list[dict] = []
    width = slide["width_emu"]
    height = slide["height_emu"]

    text_objects = [obj for obj in slide["objects"] if obj["kind"] == "text" and obj["text"]]
    for obj in slide["objects"]:
        if obj["kind"] == "chart":
            issues.append(
                {
                    "category": "high-risk-object",
                    "severity": "info",
                    "shape_id": obj["shape_id"],
                    "message": f"{obj['name']} is a chart and should be treated cautiously",
                    "suggested_fix": "manual-review",
                }
            )

        right = obj["left"] + obj["width"]
        bottom = obj["top"] + obj["height"]
        if right > width + EDGE_TOLERANCE:
            issues.append(
                {
                    "category": "boundary-overflow",
                    "severity": "error",
                    "shape_id": obj["shape_id"],
                    "message": f"{obj['name']} exceeds the right boundary",
                    "suggested_fix": "move-within-slide-bounds",
                }
            )
        elif (
            obj["left"] < EDGE_TOLERANCE
            or obj["top"] < EDGE_TOLERANCE
            or width - right < EDGE_TOLERANCE
            or height - bottom < EDGE_TOLERANCE
        ):
            issues.append(
                {
                    "category": "near-edge-crowding",
                    "severity": "warning",
                    "shape_id": obj["shape_id"],
                    "message": f"{obj['name']} is too close to a slide edge",
                    "suggested_fix": "increase-margin",
                }
            )

    if len(text_objects) >= 2:
        lefts = [obj["left"] for obj in text_objects]
        if max(lefts) - min(lefts) > ALIGN_TOLERANCE:
            issues.append(
                {
                    "category": "alignment-inconsistency",
                    "severity": "warning",
                    "shape_id": text_objects[0]["shape_id"],
                    "message": "Peer text objects do not share a consistent left edge",
                    "suggested_fix": "align-left-anchors",
                }
            )

        first_sizes = [obj["font_sizes"][0] for obj in text_objects if obj["font_sizes"]]
        if first_sizes and max(first_sizes) - min(first_sizes) > FONT_TOLERANCE:
            issues.append(
                {
                    "category": "font-hierarchy-inconsistency",
                    "severity": "warning",
                    "shape_id": text_objects[0]["shape_id"],
                    "message": "Comparable text boxes use inconsistent headline sizes",
                    "suggested_fix": "normalize-font-hierarchy",
                }
            )

    if len(text_objects) >= 3:
        gaps = _gap_values(text_objects)
        if gaps and max(gaps) - min(gaps) > SPACING_TOLERANCE:
            issues.append(
                {
                    "category": "spacing-inconsistency",
                    "severity": "warning",
                    "shape_id": text_objects[0]["shape_id"],
                    "message": "Peer text boxes use uneven horizontal spacing",
                    "suggested_fix": "normalize-peer-gaps",
                }
            )

    slide_area = width * height
    object_area = sum(obj["width"] * obj["height"] for obj in slide["objects"])
    if slide_area and (object_area / slide_area) > DENSITY_THRESHOLD:
        issues.append(
            {
                "category": "density-high",
                "severity": "warning",
                "shape_id": slide["objects"][0]["shape_id"] if slide["objects"] else 0,
                "message": "The slide uses too much of the available canvas and needs more whitespace",
                "suggested_fix": "rebalance-whitespace",
            }
        )

    return issues


def score_layout(inspection: dict) -> dict:
    slides = []
    for slide in inspection["slides"]:
        issues = _collect_issues(slide)
        slides.append({"slide_index": slide["slide_index"], "issues": issues})

    return {
        "input": inspection["input"],
        "summary": {"issue_count": sum(len(slide["issues"]) for slide in slides)},
        "slides": slides,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--inspection", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    inspection = json.loads(Path(args.inspection).read_text(encoding="utf-8"))
    findings = score_layout(inspection)
    Path(args.output).write_text(json.dumps(findings, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
