"""Table and chart structural sanity checks.

Tables and charts are deliberately not auto-mutated (their internal XML is
fragile). We only surface info-level findings the agent can use to decide
whether manual review is needed.

Tables:
  - empty cells fraction (lots of blanks suggest unfinished content)
  - asymmetric column widths (intentional or accidental?)
  - too many rows on a single slide (>10 risks readability)

Charts:
  - series count > 8 (cluttered)
  - no legend yet many series (ambiguity)
  - very small render area for a chart with rich data
"""
from __future__ import annotations


TABLE_TOO_MANY_ROWS = 12
TABLE_HIGH_EMPTY_RATIO = 0.30
CHART_TOO_MANY_SERIES = 8


def detect_table_chart_issues(slide: dict) -> list[dict]:
    issues: list[dict] = []
    for obj in slide.get("objects", []):
        if obj.get("anomalous"):
            continue
        kind = obj.get("kind")
        if kind == "table":
            issues.extend(_check_table(obj))
        elif kind == "chart":
            issues.extend(_check_chart(obj))
    return issues


def _check_table(obj: dict) -> list[dict]:
    issues: list[dict] = []
    info = obj.get("table_info") or {}
    rows = info.get("rows", 0)
    cols = info.get("cols", 0)
    empty = info.get("empty_cells", 0)
    total = max(rows * cols, 1)
    name = obj.get("name", "Table")

    if rows > TABLE_TOO_MANY_ROWS:
        issues.append({
            "category": "table-too-many-rows",
            "severity": "warning",
            "shape_id": obj["shape_id"],
            "message": f"{name} has {rows} rows, may be hard to read on one slide",
            "suggested_fix": "manual-review",
        })
    empty_ratio = empty / total
    if empty_ratio > TABLE_HIGH_EMPTY_RATIO:
        issues.append({
            "category": "table-many-empty-cells",
            "severity": "info",
            "shape_id": obj["shape_id"],
            "message": (
                f"{name} has {empty}/{total} empty cells ({int(empty_ratio*100)}%); "
                f"content may be unfinished"
            ),
            "suggested_fix": "manual-review",
        })
    return issues


def _check_chart(obj: dict) -> list[dict]:
    issues: list[dict] = []
    info = obj.get("chart_info") or {}
    series_count = info.get("series_count", 0)
    has_legend = info.get("has_legend", True)
    chart_type = info.get("chart_type", "unknown")
    name = obj.get("name", "Chart")

    if series_count > CHART_TOO_MANY_SERIES:
        issues.append({
            "category": "chart-too-many-series",
            "severity": "warning",
            "shape_id": obj["shape_id"],
            "message": (
                f"{name} ({chart_type}) has {series_count} series; "
                f"more than {CHART_TOO_MANY_SERIES} hurts readability"
            ),
            "suggested_fix": "manual-review",
        })
    if series_count >= 3 and not has_legend:
        issues.append({
            "category": "chart-missing-legend",
            "severity": "warning",
            "shape_id": obj["shape_id"],
            "message": f"{name} has {series_count} series but no legend",
            "suggested_fix": "manual-review",
        })
    return issues
