"""Design-quality refinement: background-aware text contrast + iterative polish.

Path A's `polish-business` applies theme typography (size/weight/family/
color) per-role naively — every "title" gets `text_strong`, every
"body" gets `text`, regardless of what background each text shape is
actually rendered against. Real decks often have:

  * A title sitting on a dark background → text_strong (#161616) gives
    1.07:1 contrast, fails AA.
  * A label inside a colored badge → muted gray on red → fails.
  * A footer over a slide background image → unpredictable.

This module fixes that: for each text shape, find its EFFECTIVE
background (parent container's fill OR slide background), compute
WCAG contrast against the current text color, and if < 4.5:1 (AA),
swap to whichever of {pure_white, pure_black, theme.text_strong,
theme.background} gives the best contrast on that specific
background.

Usage:
  refine_text_contrast(prs, theme, target_ratio=4.5)
    → returns list of (shape_id, old_color, new_color, ratio_before,
                       ratio_after) actions

Combine with other refinement passes (typography rhythm, spacing) for
a holistic design polish step.
"""
from __future__ import annotations

from statistics import median

from _common import contrast_ratio, hex_to_rgb, relative_luminance


# Default WCAG AA threshold for normal text. Use 3.0 for large text
# (>=18pt regular or >=14pt bold).
WCAG_AA_NORMAL = 4.5
WCAG_AA_LARGE = 3.0


# ---- background discovery ----

def _bbox_contains_center(outer: dict, inner: dict, slack: int = 91440) -> bool:
    cx = inner["left"] + inner["width"] // 2
    cy = inner["top"] + inner["height"] // 2
    return (
        outer["left"] - slack <= cx <= outer["left"] + outer["width"] + slack
        and outer["top"] - slack <= cy <= outer["top"] + outer["height"] + slack
    )


def _find_effective_background(text_obj: dict, all_objects: list[dict],
                                slide_background: str | None) -> str:
    """Return the hex color this text shape is most likely rendered ON TOP OF.

    Resolution order (most specific wins):
      1. Innermost container with fill_hex whose bbox contains the text
      2. Slide background (if known)
      3. Default white
    """
    candidates = []
    for o in all_objects:
        if o is text_obj or o.get("anomalous"):
            continue
        if o.get("kind") not in ("container", "shape"):
            continue
        if not o.get("fill_hex"):
            continue
        if _bbox_contains_center(o, text_obj):
            candidates.append(o)
    if candidates:
        # Innermost = smallest containing area
        candidates.sort(key=lambda o: o["width"] * o["height"])
        return candidates[0]["fill_hex"]
    return slide_background or "FFFFFF"


def _best_contrast_color(bg_hex: str, palette_options: list[str]) -> tuple[str, float]:
    """Pick the palette option that gives the highest contrast on bg_hex.

    Returns (hex_color, contrast_ratio).
    """
    best = ("000000", 0.0)
    for option in palette_options:
        try:
            ratio = contrast_ratio(option, bg_hex)
        except (ValueError, KeyError):
            continue
        if ratio > best[1]:
            best = (option, ratio)
    return best


def _is_large_text(text_obj: dict) -> bool:
    """WCAG large-text threshold: >= 18pt regular or >= 14pt bold."""
    sizes = text_obj.get("font_sizes") or []
    if not sizes:
        return False
    max_pt = max(sizes) / 12700  # EMU → pt
    return max_pt >= 18.0  # bold check would need run-level data


# ---- main refine function ----

def refine_text_contrast(prs, theme, slide_inspections: list[dict],
                          target_ratio: float = WCAG_AA_NORMAL) -> dict:
    """For every text shape across all slides, ensure its color contrasts
    enough with its effective background. Returns action log.

    `theme` is a _common.Theme dataclass; we use its palette to pick
    candidate colors.
    """
    from _shape_ops import set_font_color

    palette = (theme.raw if hasattr(theme, "raw") else theme).get("palette", {})
    candidates = list({
        palette.get("text_strong", "161616"),
        palette.get("text", "393939"),
        palette.get("background", "FFFFFF"),
        "FFFFFF",
        "000000",
    })

    actions: list[dict] = []
    skipped: list[dict] = []

    sid_index: dict[int, object] = {}
    for slide in prs.slides:
        for shape in slide.shapes:
            sid = getattr(shape, "shape_id", None)
            if sid is not None:
                sid_index[int(sid)] = shape

    for slide_idx, slide_data in enumerate(slide_inspections, start=1):
        objs = slide_data.get("objects", [])
        # Slide background unknown from inspection.json directly; fall back
        # to white. The orchestrator should pass an explicit override
        # via slide_data["__background_hex"] when known.
        slide_bg = slide_data.get("__background_hex")

        for obj in objs:
            if obj.get("kind") != "text":
                continue
            if not obj.get("text"):
                continue
            sid = obj["shape_id"]
            shape = sid_index.get(sid)
            if shape is None:
                continue

            current = obj.get("text_color")
            if not current:
                # No explicit color set — inherits from theme. Skip; theme
                # will decide. We focus on shapes with explicit colors that
                # we KNOW will render that way.
                continue

            bg_hex = _find_effective_background(obj, objs, slide_bg)
            try:
                ratio_before = contrast_ratio(current, bg_hex)
            except (ValueError, KeyError):
                continue

            threshold = WCAG_AA_LARGE if _is_large_text(obj) else target_ratio
            if ratio_before >= threshold:
                continue   # passes already

            new_color, ratio_after = _best_contrast_color(bg_hex, candidates)
            if ratio_after <= ratio_before + 0.5:
                # Could not meaningfully improve — log and skip.
                skipped.append({
                    "shape_id": sid,
                    "reason": "no-better-color",
                    "ratio_before": round(ratio_before, 2),
                    "best_available": round(ratio_after, 2),
                })
                continue

            try:
                set_font_color(shape, new_color)
                actions.append({
                    "action": "refine-contrast",
                    "shape_id": sid,
                    "slide_index": slide_idx,
                    "background_hex": bg_hex,
                    "old_color": current,
                    "new_color": new_color,
                    "ratio_before": round(ratio_before, 2),
                    "ratio_after": round(ratio_after, 2),
                    "is_large_text": _is_large_text(obj),
                })
            except (AttributeError, ValueError):
                continue

    return {
        "applied": len(actions),
        "skipped": len(skipped),
        "actions": actions,
        "skipped_details": skipped,
    }


# ---- background-luminance theme picker ----

def detect_background_luminance(slide_inspection: dict) -> float:
    """Average luminance of the largest container / slide-background fills.

    Returns 0.0 (dark) to 1.0 (light).
    """
    objs = slide_inspection.get("objects", [])
    fills = []
    for o in objs:
        if o.get("anomalous"):
            continue
        if not o.get("fill_hex"):
            continue
        # Weight by area
        area = o.get("width", 0) * o.get("height", 0)
        if area <= 0:
            continue
        try:
            lum = relative_luminance(o["fill_hex"])
        except (ValueError, KeyError):
            continue
        fills.append((area, lum))
    if not fills:
        return 1.0   # no info → assume light (safe default)
    total_area = sum(a for a, _ in fills)
    weighted = sum(a * l for a, l in fills) / max(total_area, 1)
    return weighted


def pick_theme_for_background(content_text: str, themes_dir,
                                background_lum: float | None = None) -> str:
    """Theme picker that BIAS for background luminance.

    Falls back to keyword-only when luminance unknown.
    """
    from _business_polish import pick_theme as keyword_pick
    base = keyword_pick(content_text, themes_dir)

    if background_lum is None:
        return base
    # Light/dark thresholds. WCAG-style luminance < 0.18 = dark.
    if background_lum < 0.18:
        # Slide is dark — prefer dark themes regardless of keywords.
        if base in ("clean-tech", "business-warm", "academic-soft"):
            # Map to closest dark equivalent.
            return "claude-code" if base == "clean-tech" else "editorial-dark"
        return base   # already a dark theme
    elif background_lum > 0.82:
        # Slide is bright — keep keyword pick (it's already light-ish).
        return base
    return base
