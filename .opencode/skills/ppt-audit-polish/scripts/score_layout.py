from __future__ import annotations

import argparse
import json
from pathlib import Path


ALIGN_TOLERANCE = 120000
EDGE_TOLERANCE = 80000
FONT_TOLERANCE = 30000
SPACING_TOLERANCE = 140000
DENSITY_THRESHOLD = 0.55
ROW_TOLERANCE = 220000
MIN_DIMENSION_TOLERANCE = 180000
DIMENSION_RATIO_TOLERANCE = 0.35


def _gap_values(text_objects: list[dict]) -> list[int]:
    ordered = sorted(text_objects, key=lambda obj: obj["left"])
    gaps: list[int] = []
    for previous, current in zip(ordered, ordered[1:]):
        gaps.append(current["left"] - (previous["left"] + previous["width"]))
    return gaps


def _is_layout_text(obj: dict) -> bool:
    return obj["kind"] == "text" and bool(obj["text"]) and obj["width"] > 0 and obj["height"] > 0


def _first_font_size(obj: dict) -> int | None:
    return obj["font_sizes"][0] if obj["font_sizes"] else None


def _dimension_close(value: int, target: int) -> bool:
    tolerance = max(MIN_DIMENSION_TOLERANCE, int(abs(target) * DIMENSION_RATIO_TOLERANCE))
    return abs(value - target) <= tolerance


def _row_groups(text_objects: list[dict]) -> list[list[dict]]:
    rows: list[list[dict]] = []
    for obj in sorted(text_objects, key=lambda item: item["top"]):
        for row in rows:
            row_top = sum(item["top"] for item in row) / len(row)
            if abs(obj["top"] - row_top) <= ROW_TOLERANCE:
                row.append(obj)
                break
        else:
            rows.append([obj])
    return rows


def _peer_groups(text_objects: list[dict]) -> list[list[dict]]:
    groups: list[list[dict]] = []
    for row in _row_groups(text_objects):
        row_groups: list[list[dict]] = []
        for obj in sorted(row, key=lambda item: (item["width"], item["height"])):
            for group in row_groups:
                avg_width = int(sum(item["width"] for item in group) / len(group))
                avg_height = int(sum(item["height"] for item in group) / len(group))
                if _dimension_close(obj["width"], avg_width) and _dimension_close(obj["height"], avg_height):
                    group.append(obj)
                    break
            else:
                row_groups.append([obj])
        groups.extend(group for group in row_groups if len(group) >= 2)
    return groups


def _collect_issues(slide: dict) -> list[dict]:
    issues: list[dict] = []
    width = slide["width_emu"]
    height = slide["height_emu"]

    text_objects = [obj for obj in slide["objects"] if _is_layout_text(obj)]
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

    for peer_group in _peer_groups(text_objects):
        tops = [obj["top"] for obj in peer_group]
        if max(tops) - min(tops) > ALIGN_TOLERANCE:
            issues.append(
                {
                    "category": "alignment-inconsistency",
                    "severity": "warning",
                    "shape_id": peer_group[0]["shape_id"],
                    "message": "Peer text objects in a row do not share a consistent top edge",
                    "suggested_fix": "align-peer-row",
                }
            )

        first_sizes = [size for size in (_first_font_size(obj) for obj in peer_group) if size is not None]
        if first_sizes and max(first_sizes) - min(first_sizes) > FONT_TOLERANCE:
            issues.append(
                {
                    "category": "font-hierarchy-inconsistency",
                    "severity": "warning",
                    "shape_id": peer_group[0]["shape_id"],
                    "message": "Comparable text boxes use inconsistent headline sizes",
                    "suggested_fix": "normalize-font-hierarchy",
                }
            )

        gaps = _gap_values(peer_group)
        if gaps and max(gaps) - min(gaps) > SPACING_TOLERANCE:
            issues.append(
                {
                    "category": "spacing-inconsistency",
                    "severity": "warning",
                    "shape_id": peer_group[0]["shape_id"],
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
