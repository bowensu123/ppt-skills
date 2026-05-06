"""Color-harmony detection: WCAG text-on-background contrast + palette clash.

Two checks:
  1. low-contrast-text -- a text shape's font color vs its parent container's
     fill color must meet WCAG AA contrast (4.5:1 normal, 3.0:1 for >=18pt).
  2. palette-too-busy -- if the slide uses more than PALETTE_TARGET_MAX
     distinct fills (already counted by score_layout's palette_score, but
     surfaced as an issue so the orchestrator can target it).
"""
from __future__ import annotations

from typing import Iterable

from _common import contrast_ratio


WCAG_NORMAL_RATIO = 4.5
WCAG_LARGE_RATIO = 3.0
PALETTE_TARGET_MAX = 6
LARGE_TEXT_PT = 18.0


def _bbox_contains(outer: dict, inner: dict, slack: int = 91440) -> bool:
    return (
        inner["left"] >= outer["left"] - slack
        and inner["top"] >= outer["top"] - slack
        and inner["left"] + inner["width"] <= outer["left"] + outer["width"] + slack
        and inner["top"] + inner["height"] <= outer["top"] + outer["height"] + slack
    )


def _find_container(text_obj: dict, all_objects: Iterable[dict]) -> dict | None:
    """Smallest filled container that visually frames this text."""
    candidates = [
        o for o in all_objects
        if o["shape_id"] != text_obj["shape_id"]
        and not o.get("anomalous")
        and o.get("kind") == "container"
        and o.get("fill_hex")
        and _bbox_contains(o, text_obj)
    ]
    if not candidates:
        return None
    return min(candidates, key=lambda c: c["width"] * c["height"])


def detect_color_issues(slide: dict) -> list[dict]:
    issues: list[dict] = []
    objects = slide.get("objects", [])

    # 1. WCAG contrast for text on a clearly-identifiable background.
    for obj in objects:
        if obj.get("kind") != "text" or not obj.get("text"):
            continue
        if obj.get("anomalous"):
            continue
        text_color = obj.get("text_color")
        if not text_color:
            continue
        container = _find_container(obj, objects)
        bg_color = container["fill_hex"] if container else "FFFFFF"  # assume white slide bg
        ratio = contrast_ratio(text_color, bg_color)
        font_pt = max(obj.get("font_sizes") or [11 * 12700]) / 12700.0
        min_ratio = WCAG_LARGE_RATIO if font_pt >= LARGE_TEXT_PT else WCAG_NORMAL_RATIO
        if ratio < min_ratio:
            issues.append({
                "category": "low-contrast-text",
                "severity": "warning" if ratio >= min_ratio - 1.0 else "error",
                "shape_id": obj["shape_id"],
                "message": (
                    f"{obj['name']} text {text_color} on background {bg_color} "
                    f"has contrast ratio {ratio:.1f}:1 (WCAG AA needs {min_ratio}:1 at {font_pt:.0f}pt)"
                ),
                "suggested_fix": "set-font-color",
                "suggested_argv": [
                    "set-font-color",
                    "--shape-id", str(obj["shape_id"]),
                    # Pick highest-contrast safe choice based on bg luminance.
                    "--color", _pick_legible_text_color(bg_color),
                ],
                "contrast_ratio": round(ratio, 2),
                "min_required": min_ratio,
            })

    # 2. Palette-too-busy.
    distinct_fills = {o["fill_hex"] for o in objects if o.get("fill_hex")}
    if len(distinct_fills) > PALETTE_TARGET_MAX:
        issues.append({
            "category": "palette-too-busy",
            "severity": "warning",
            "shape_id": 0,
            "message": (
                f"Slide uses {len(distinct_fills)} distinct fill colors "
                f"(target ≤ {PALETTE_TARGET_MAX})"
            ),
            "suggested_fix": "manual-review",
            "fill_count": len(distinct_fills),
        })

    return issues


def _pick_legible_text_color(bg_hex: str) -> str:
    from _common import contrast_ratio
    if contrast_ratio("FFFFFF", bg_hex) >= contrast_ratio("161616", bg_hex):
        return "FFFFFF"
    return "161616"
