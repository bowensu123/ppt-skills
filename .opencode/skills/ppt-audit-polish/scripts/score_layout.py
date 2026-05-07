"""Role-aware layout scorer.

Emits two complementary outputs per slide:
  * ``issues``  - qualitative findings (each maps to a suggested mutate op)
  * ``metrics`` - quantitative scores in [0, 100] per dimension

Quantitative metrics let the L4 self-critique compute a single fitness score
and let the orchestrator run a greedy hill-climb without re-implementing each
heuristic.
"""
from __future__ import annotations

import argparse
import json
import math
from collections import Counter
from pathlib import Path
from statistics import median, pstdev


EDGE_TOLERANCE = 80000             # how far past slide bounds counts as "overflow"
NEAR_EDGE_THRESHOLD = 100000       # inside-but-this-close-to-edge counts as "crowding"  (~0.11")
ROW_ALIGN_TOLERANCE = 80000
ROW_GAP_VARIANCE_TOLERANCE = 137160
ROW_FONT_SIZE_TOLERANCE = 30000
DENSITY_TARGET_LOW = 0.35
DENSITY_TARGET_HIGH = 0.60
PALETTE_TARGET_MAX = 6
OVERLAP_AREA_RATIO_THRESHOLD = 0.10   # report when overlap > 10% of smaller shape's area
OVERLAP_BBOX_SLACK = 91440            # 0.1" — when checking parent-child containment


def _bbox_intersect_area(a: dict, b: dict) -> int:
    x1 = max(a["left"], b["left"])
    y1 = max(a["top"], b["top"])
    x2 = min(a["left"] + a["width"], b["left"] + b["width"])
    y2 = min(a["top"] + a["height"], b["top"] + b["height"])
    if x2 <= x1 or y2 <= y1:
        return 0
    return (x2 - x1) * (y2 - y1)


def _bbox_contains(outer: dict, inner: dict, slack: int = OVERLAP_BBOX_SLACK) -> bool:
    return (
        inner["left"] >= outer["left"] - slack
        and inner["top"] >= outer["top"] - slack
        and inner["left"] + inner["width"] <= outer["left"] + outer["width"] + slack
        and inner["top"] + inner["height"] <= outer["top"] + outer["height"] + slack
    )


def _detect_overlaps(slide: dict) -> list[dict]:
    """Pairwise bbox intersection, excluding parent-child cases.

    A 'parent-child' is when one shape fully contains the other (with a small
    slack to allow visual tucks like a numbered badge in a card corner).
    Connectors and groups are excluded — connectors legitimately cross cards
    by design, groups are containers.
    """
    from itertools import combinations

    candidates = [
        o for o in slide["objects"]
        if not o.get("anomalous")
        and o.get("kind") not in ("group", "connector")
        and o.get("width", 0) > 0
        and o.get("height", 0) > 0
    ]
    issues: list[dict] = []
    for a, b in combinations(candidates, 2):
        if _bbox_contains(a, b) or _bbox_contains(b, a):
            continue  # nested layout — fine
        area = _bbox_intersect_area(a, b)
        if area == 0:
            continue
        a_area = max(a["width"] * a["height"], 1)
        b_area = max(b["width"] * b["height"], 1)
        smaller_area = min(a_area, b_area)
        ratio = area / smaller_area
        if ratio < OVERLAP_AREA_RATIO_THRESHOLD:
            continue  # tiny overlap — likely intentional decorative touch
        # Move the SMALLER shape; pick smallest movement direction.
        smaller = a if a_area <= b_area else b
        larger = b if a_area <= b_area else a
        suggested_argv = _suggest_overlap_resolution(smaller, larger)
        sev = "error" if ratio > 0.5 else "warning"
        issues.append({
            "category": "shape-overlap",
            "severity": sev,
            "shape_id": smaller["shape_id"],
            "shape_ids": [smaller["shape_id"], larger["shape_id"]],
            "message": (
                f"{smaller['name']} overlaps {larger['name']} "
                f"({int(ratio * 100)}% of {smaller['name']}'s area)"
            ),
            "suggested_fix": "resolve-overlap",
            "suggested_argv": suggested_argv,
            "overlap_ratio": round(ratio, 3),
        })
    return issues


def _suggest_overlap_resolution(smaller: dict, larger: dict) -> list[str]:
    """Pick the cheapest direction to nudge `smaller` out of `larger`.

    Returns argv for `mutate.py nudge` that moves smaller to no longer overlap
    larger. Strategy: compute required dx/dy to clear in each direction and
    pick the smallest absolute displacement.
    """
    a = smaller; b = larger
    a_right = a["left"] + a["width"]; b_right = b["left"] + b["width"]
    a_bottom = a["top"] + a["height"]; b_bottom = b["top"] + b["height"]
    options = [
        ("left",  -(a_right - b["left"]) - 50000),       # push left until right-of-a < left-of-b
        ("right",  b_right - a["left"] + 50000),          # push right until left-of-a > right-of-b
        ("up",    -(a_bottom - b["top"]) - 50000),        # push up until bottom-of-a < top-of-b
        ("down",   b_bottom - a["top"] + 50000),          # push down until top-of-a > bottom-of-b
    ]
    options.sort(key=lambda o: abs(o[1]))
    direction, displacement = options[0]
    if direction in ("left", "right"):
        return ["nudge", "--shape-id", str(a["shape_id"]), "--dx", str(displacement)]
    return ["nudge", "--shape-id", str(a["shape_id"]), "--dy", str(displacement)]


def _detect_boundary_issues(obj: dict, width: int, height: int) -> list[dict]:
    """Detect overflow on all 4 edges + 'near-edge crowding' inside bounds."""
    issues: list[dict] = []
    name = obj["name"]
    sid = obj["shape_id"]
    left = obj["left"]; top = obj["top"]
    right = left + obj["width"]; bottom = top + obj["height"]

    # Hard overflow (any side past slide bounds + tolerance)
    if right > width + EDGE_TOLERANCE:
        issues.append({
            "category": "boundary-overflow",
            "severity": "error",
            "shape_id": sid,
            "message": f"{name} exceeds the right boundary by {right - width} EMU",
            "suggested_fix": "move-within-slide-bounds",
        })
    if bottom > height + EDGE_TOLERANCE:
        issues.append({
            "category": "boundary-overflow",
            "severity": "error",
            "shape_id": sid,
            "message": f"{name} exceeds the bottom boundary by {bottom - height} EMU",
            "suggested_fix": "move-within-slide-bounds",
        })
    if left < -EDGE_TOLERANCE:
        issues.append({
            "category": "boundary-overflow",
            "severity": "error",
            "shape_id": sid,
            "message": f"{name} extends past the left boundary by {-left} EMU",
            "suggested_fix": "move-within-slide-bounds",
        })
    if top < -EDGE_TOLERANCE:
        issues.append({
            "category": "boundary-overflow",
            "severity": "error",
            "shape_id": sid,
            "message": f"{name} extends past the top boundary by {-top} EMU",
            "suggested_fix": "move-within-slide-bounds",
        })

    # Near-edge crowding (inside bounds but uncomfortably close to an edge)
    no_overflow = (
        right <= width + EDGE_TOLERANCE
        and bottom <= height + EDGE_TOLERANCE
        and left >= -EDGE_TOLERANCE
        and top >= -EDGE_TOLERANCE
    )
    if no_overflow:
        if 0 <= left < NEAR_EDGE_THRESHOLD:
            issues.append({
                "category": "near-edge-crowding",
                "severity": "warning",
                "shape_id": sid,
                "message": f"{name} sits {left} EMU from the left edge — uncomfortable margin",
                "suggested_fix": "increase-margin",
            })
        if 0 <= top < NEAR_EDGE_THRESHOLD:
            issues.append({
                "category": "near-edge-crowding",
                "severity": "warning",
                "shape_id": sid,
                "message": f"{name} sits {top} EMU from the top edge — uncomfortable margin",
                "suggested_fix": "increase-margin",
            })
        if 0 <= width - right < NEAR_EDGE_THRESHOLD:
            issues.append({
                "category": "near-edge-crowding",
                "severity": "warning",
                "shape_id": sid,
                "message": f"{name} sits {width - right} EMU from the right edge",
                "suggested_fix": "increase-margin",
            })
        if 0 <= height - bottom < NEAR_EDGE_THRESHOLD:
            issues.append({
                "category": "near-edge-crowding",
                "severity": "warning",
                "shape_id": sid,
                "message": f"{name} sits {height - bottom} EMU from the bottom edge",
                "suggested_fix": "increase-margin",
            })

    return issues


def _gap_variance(items: list[dict]) -> int:
    if len(items) < 3:
        return 0
    ordered = sorted(items, key=lambda obj: obj["left"])
    gaps = [
        ordered[i + 1]["left"] - (ordered[i]["left"] + ordered[i]["width"])
        for i in range(len(ordered) - 1)
    ]
    return max(gaps) - min(gaps)


def _svg_signals_to_issues(svg_signals: dict | None, slide_index: int) -> list[dict]:
    """Convert SVG-derived post-render signals into issue dicts.

    Three signal types map to issues; connector-snap-drift is intentionally
    excluded from issue surface (low severity, usually self-resolves).
    """
    if not svg_signals:
        return []
    out: list[dict] = []
    for sig in svg_signals.get("text_overflow", []):
        if sig.get("slide_index") != slide_index:
            continue
        out.append({
            "category": "text-overflow",
            "severity": "warning",
            "shape_id": sig["shape_id"],
            "message": (
                f"Text in '{sig.get('name')}' overflows its frame: "
                f"rendered ~{sig['rendered_height_emu']} EMU vs declared "
                f"{sig['declared_height_emu']} EMU ({sig['wrap_lines']} wrapped lines)"
            ),
            "suggested_fix": "resize-or-shrink-font",
            "svg_signal": sig,
        })
    for sig in svg_signals.get("font_fallback", []):
        if sig.get("slide_index") != slide_index:
            continue
        out.append({
            "category": "font-fallback",
            "severity": "info",
            "shape_id": sig["shape_id"],
            "message": (
                f"Renderer used {sig['rendered_fonts']} for '{sig.get('name')}' "
                f"although the deck declares {sig['declared_fonts']} (font not "
                f"installed on this system)"
            ),
            "suggested_fix": "unify-font",
            "svg_signal": sig,
        })
    for sig in svg_signals.get("z_order_drift", []):
        if sig.get("slide_index") != slide_index:
            continue
        out.append({
            "category": "z-order-real-drift",
            "severity": "warning",
            "shape_id": sig["hides"],
            "message": (
                f"Shape #{sig['hides']} is hidden by #{sig['covers']} in the "
                f"actual render even though declared z-order says otherwise "
                f"(overlap {sig['overlap_ratio']})"
            ),
            "suggested_fix": "z-order-bring-to-front",
            "svg_signal": sig,
        })
    return out


def _collect_issues(slide: dict, role_data: dict, svg_signals: dict | None = None) -> list[dict]:
    issues: list[dict] = []
    width = slide["width_emu"]
    height = slide["height_emu"]
    obj_by_id = {obj["shape_id"]: obj for obj in slide["objects"]}

    for obj in slide["objects"]:
        if obj.get("anomalous") or obj.get("kind") in ("connector", "group"):
            continue
        issues.extend(_detect_boundary_issues(obj, width, height))

    issues.extend(_detect_overlaps(slide))

    for row in role_data.get("rows", []):
        members = [obj_by_id[i] for i in row["shape_ids"] if i in obj_by_id]
        if len(members) < 2:
            continue
        tops = [m["top"] for m in members]
        if max(tops) - min(tops) > ROW_ALIGN_TOLERANCE:
            issues.append({
                "category": "row-alignment-drift",
                "severity": "warning",
                "shape_id": members[0]["shape_id"],
                "shape_ids": [m["shape_id"] for m in members],
                "message": f"Row {row['row_id']} peers are not vertically aligned",
                "suggested_fix": "align-row-tops",
            })
        if len(members) >= 3 and _gap_variance(members) > ROW_GAP_VARIANCE_TOLERANCE:
            issues.append({
                "category": "row-spacing-uneven",
                "severity": "warning",
                "shape_id": members[0]["shape_id"],
                "shape_ids": [m["shape_id"] for m in members],
                "message": f"Row {row['row_id']} peers have uneven horizontal gaps",
                "suggested_fix": "equalize-row-gaps",
            })
        primary_sizes = [max(m["font_sizes"]) for m in members if m["font_sizes"]]
        if primary_sizes and max(primary_sizes) - min(primary_sizes) > ROW_FONT_SIZE_TOLERANCE:
            issues.append({
                "category": "row-font-hierarchy-drift",
                "severity": "warning",
                "shape_id": members[0]["shape_id"],
                "shape_ids": [m["shape_id"] for m in members],
                "message": f"Row {row['row_id']} peers use inconsistent font sizes",
                "suggested_fix": "normalize-row-font",
            })

    for col in role_data.get("columns", []):
        members = [obj_by_id[i] for i in col["shape_ids"] if i in obj_by_id]
        if len(members) < 2:
            continue
        lefts = [m["left"] for m in members]
        if max(lefts) - min(lefts) > ROW_ALIGN_TOLERANCE:
            issues.append({
                "category": "column-alignment-drift",
                "severity": "warning",
                "shape_id": members[0]["shape_id"],
                "shape_ids": [m["shape_id"] for m in members],
                "message": f"Column {col['col_id']} peers are not left-aligned",
                "suggested_fix": "align-column-lefts",
            })

    has_title = any(entry.get("role") == "title" for entry in role_data.get("shapes", []))
    # Suppress missing-title when the slide layout/master provides a title
    # placeholder — that title IS rendered, just inherited.
    inheritance = slide.get("inheritance") or {}
    inherited_title = False
    try:
        from _master_inherit import has_inherited_title
        inherited_title = has_inherited_title(inheritance)
    except ImportError:
        pass
    if (not has_title) and (not inherited_title) and any(obj["kind"] == "text" and obj["text"] for obj in slide["objects"]):
        issues.append({
            "category": "missing-title",
            "severity": "info",
            "shape_id": 0,
            "message": "Slide has no clear title shape (no slide-level or inherited title)",
            "suggested_fix": "manual-review",
        })

    # Peer-card outlier detection — produces issues that the orchestrator
    # can dispatch to the `repair-peer-cards` mutate op.
    try:
        from _card_repair import diagnose_repair
        plan = diagnose_repair({"objects": slide["objects"], "width_emu": width, "height_emu": height})
        for row in plan.get("rows", []):
            for fix in row.get("card_box_fixes", []):
                issues.append({
                    "category": "peer-card-box-outlier",
                    "severity": "warning",
                    "shape_id": fix["shape_id"],
                    "message": f"Card '{fix['name']}' has outlier dimensions vs row peers",
                    "suggested_fix": "repair-peer-cards",
                })
            for fix in row.get("header_strip_fixes", []):
                issues.append({
                    "category": "peer-card-header-outlier",
                    "severity": "warning",
                    "shape_id": fix["shape_id"],
                    "message": f"Header '{fix['name']}' is much wider than peer header strips",
                    "suggested_fix": "repair-peer-cards",
                })
            for fix in row.get("displaced_relocations", []):
                issues.append({
                    "category": "peer-card-misplaced-child",
                    "severity": "warning",
                    "shape_id": fix["shape_id"],
                    "message": f"'{fix['name']}' sits inside a different card than its peer slot",
                    "suggested_fix": "repair-peer-cards",
                })
            for fix in row.get("orphan_relocations", []):
                issues.append({
                    "category": "peer-card-orphan-shape",
                    "severity": "warning",
                    "shape_id": fix["shape_id"],
                    "message": f"'{fix['name']}' floats in the row band but inside no card",
                    "suggested_fix": "repair-peer-cards",
                })
    except Exception:
        pass  # never let detection break the pipeline

    # New blind-spot detectors (each is best-effort; failures are silenced
    # so the pipeline always produces SOMETHING).
    for module_name, fn_name in (
        ("_color_harmony", "detect_color_issues"),
        ("_image_check", "detect_image_issues"),
        ("_table_chart", "detect_table_chart_issues"),
    ):
        try:
            mod = __import__(module_name)
            fn = getattr(mod, fn_name)
            issues.extend(fn(slide))
        except Exception:
            pass

    # Visual-balance is a metric AND a possible issue.
    try:
        from _visual_balance import compute_balance
        balance = compute_balance(slide)
        if balance.get("issue"):
            issues.append(balance["issue"])
    except Exception:
        pass

    # SVG-derived post-render signals (text-overflow, font-fallback,
    # z-order-real-drift). These come from _svg_geom.extract_signals()
    # which only runs when the caller passed --also-svg to render_slides.
    issues.extend(_svg_signals_to_issues(svg_signals, slide["slide_index"]))

    return issues


def _alignment_score(slide: dict, role_data: dict) -> float:
    """Aggregate row/column alignment quality. 100 = perfectly aligned."""
    obj_by_id = {obj["shape_id"]: obj for obj in slide["objects"]}
    drifts: list[float] = []

    for row in role_data.get("rows", []):
        members = [obj_by_id[i] for i in row["shape_ids"] if i in obj_by_id]
        if len(members) < 2:
            continue
        tops = [m["top"] for m in members]
        drift = (max(tops) - min(tops)) / ROW_ALIGN_TOLERANCE
        drifts.append(min(drift, 4.0))

    for col in role_data.get("columns", []):
        members = [obj_by_id[i] for i in col["shape_ids"] if i in obj_by_id]
        if len(members) < 2:
            continue
        lefts = [m["left"] for m in members]
        drift = (max(lefts) - min(lefts)) / ROW_ALIGN_TOLERANCE
        drifts.append(min(drift, 4.0))

    if not drifts:
        return 100.0
    avg_drift = sum(drifts) / len(drifts)
    return max(0.0, 100.0 - 25.0 * avg_drift)


def _density_score(slide: dict) -> float:
    """100 when usage is in [DENSITY_TARGET_LOW, DENSITY_TARGET_HIGH]."""
    width = slide["width_emu"]
    height = slide["height_emu"]
    used = sum(
        obj["width"] * obj["height"]
        for obj in slide["objects"]
        if not obj.get("anomalous") and obj.get("kind") not in ("group", "connector")
    )
    slide_area = width * height
    if slide_area == 0:
        return 0.0
    ratio = used / slide_area
    if DENSITY_TARGET_LOW <= ratio <= DENSITY_TARGET_HIGH:
        return 100.0
    if ratio < DENSITY_TARGET_LOW:
        return max(0.0, 100.0 - 200.0 * (DENSITY_TARGET_LOW - ratio))
    return max(0.0, 100.0 - 200.0 * (ratio - DENSITY_TARGET_HIGH))


def _hierarchy_score(slide: dict, role_data: dict) -> float:
    """Reward 3-5 distinct font sizes (clear hierarchy), penalize <2 or >7."""
    sizes = []
    for obj in slide["objects"]:
        if obj.get("kind") != "text":
            continue
        if obj["font_sizes"]:
            sizes.append(max(obj["font_sizes"]))
    if not sizes:
        return 50.0
    distinct = len({round(s / 12700) for s in sizes})
    if 3 <= distinct <= 5:
        return 100.0
    if distinct == 2:
        return 80.0
    if distinct == 1:
        return 50.0
    if distinct == 6:
        return 80.0
    if distinct == 7:
        return 60.0
    return 30.0


def _palette_score(slide: dict) -> float:
    """100 when distinct fill colors <= PALETTE_TARGET_MAX."""
    colors = []
    for obj in slide["objects"]:
        if obj.get("fill_hex"):
            colors.append(obj["fill_hex"])
    distinct = len(set(colors))
    if distinct == 0:
        return 100.0
    if distinct <= PALETTE_TARGET_MAX:
        return 100.0
    return max(0.0, 100.0 - 10.0 * (distinct - PALETTE_TARGET_MAX))


def _hierarchy_entropy(slide: dict) -> float:
    """Shannon entropy over rounded font sizes (debug-only signal)."""
    sizes = []
    for obj in slide["objects"]:
        if obj.get("kind") == "text" and obj["font_sizes"]:
            sizes.append(round(max(obj["font_sizes"]) / 12700))
    if not sizes:
        return 0.0
    counts = Counter(sizes)
    total = sum(counts.values())
    return -sum((c / total) * math.log2(c / total) for c in counts.values())


def _overflow_count(slide: dict) -> int:
    width = slide["width_emu"]
    height = slide["height_emu"]
    n = 0
    for obj in slide["objects"]:
        if obj.get("anomalous") or obj.get("kind") in ("group", "connector"):
            continue
        right = obj["left"] + obj["width"]
        bottom = obj["top"] + obj["height"]
        if right > width + EDGE_TOLERANCE or bottom > height + EDGE_TOLERANCE:
            n += 1
    return n


def _collect_metrics(slide: dict, role_data: dict, issues: list[dict]) -> dict:
    counts_by_severity = Counter(issue["severity"] for issue in issues)
    counts_by_category = Counter(issue["category"] for issue in issues)

    width = slide["width_emu"]
    height = slide["height_emu"]
    used = sum(
        obj["width"] * obj["height"]
        for obj in slide["objects"]
        if not obj.get("anomalous") and obj.get("kind") not in ("group", "connector")
    )
    density_ratio = used / (width * height) if width and height else 0.0

    text_objects = [o for o in slide["objects"] if o["kind"] == "text" and o["text"]]
    distinct_sizes = len({round(max(o["font_sizes"]) / 12700) for o in text_objects if o["font_sizes"]})
    distinct_fills = len({o["fill_hex"] for o in slide["objects"] if o.get("fill_hex")})

    # Visual balance score (mass-center distance from slide center).
    balance_score = 50.0
    try:
        from _visual_balance import compute_balance
        balance_score = compute_balance(slide).get("score", 50.0)
    except Exception:
        pass

    return {
        "alignment_score": round(_alignment_score(slide, role_data), 2),
        "density_score": round(_density_score(slide), 2),
        "hierarchy_score": round(_hierarchy_score(slide, role_data), 2),
        "palette_score": round(_palette_score(slide), 2),
        "balance_score": round(balance_score, 2),
        "density_ratio": round(density_ratio, 4),
        "hierarchy_entropy": round(_hierarchy_entropy(slide), 4),
        "distinct_font_sizes": distinct_sizes,
        "distinct_fill_colors": distinct_fills,
        "shape_count": len(slide["objects"]),
        "text_shape_count": len(text_objects),
        "overflow_count": _overflow_count(slide),
        "issue_count": len(issues),
        "issues_by_severity": dict(counts_by_severity),
        "issues_by_category": dict(counts_by_category),
    }


def score_layout(inspection: dict, roles_payload: dict, svg_signals: dict | None = None) -> dict:
    role_by_slide = {s["slide_index"]: s for s in roles_payload["slides"]}
    slides = []
    aggregate_metrics: list[dict] = []
    for slide in inspection["slides"]:
        role_data = role_by_slide.get(slide["slide_index"], {"shapes": [], "rows": [], "columns": []})
        issues = _collect_issues(slide, role_data, svg_signals=svg_signals)
        metrics = _collect_metrics(slide, role_data, issues)
        slides.append({
            "slide_index": slide["slide_index"],
            "issues": issues,
            "metrics": metrics,
        })
        aggregate_metrics.append(metrics)

    summary_score_keys = ("alignment_score", "density_score", "hierarchy_score", "palette_score", "balance_score")
    deck_metrics = {
        key: round(sum(m[key] for m in aggregate_metrics) / max(len(aggregate_metrics), 1), 2)
        for key in summary_score_keys
    }
    deck_metrics["issue_count"] = sum(m["issue_count"] for m in aggregate_metrics)
    deck_metrics["overflow_count"] = sum(m["overflow_count"] for m in aggregate_metrics)
    deck_metrics["slide_count"] = len(aggregate_metrics)

    return {
        "input": inspection["input"],
        "summary": {"issue_count": deck_metrics["issue_count"]},
        "metrics": deck_metrics,
        "slides": slides,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--inspection", required=True)
    parser.add_argument("--roles", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument(
        "--svg-signals",
        help="Optional path to a JSON dump of _svg_geom.extract_signals(). "
             "When passed, post-render SVG-derived issues "
             "(text-overflow / font-fallback / z-order-real-drift) are added.",
    )
    args = parser.parse_args()

    inspection = json.loads(Path(args.inspection).read_text(encoding="utf-8"))
    roles_payload = json.loads(Path(args.roles).read_text(encoding="utf-8"))
    svg_signals = None
    if args.svg_signals and Path(args.svg_signals).exists():
        svg_signals = json.loads(Path(args.svg_signals).read_text(encoding="utf-8"))
    findings = score_layout(inspection, roles_payload, svg_signals=svg_signals)
    Path(args.output).write_text(json.dumps(findings, indent=2, ensure_ascii=False), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
