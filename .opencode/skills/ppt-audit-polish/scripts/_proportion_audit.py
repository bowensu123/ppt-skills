"""Proportion / composition audit: catch visual-balance issues that
contrast/typography fixes don't address.

Looking at real "polished" decks that still feel awkward, the most
common complaints are:

  * Icon obviously oversized relative to its card (eats half the
    card body)
  * Icon obviously undersized (a tiny emoji floating in a big empty
    box, like the iter1_huawei.pptx case)
  * Card has too much empty space (content occupies < 30% of card area)
  * Top-heavy composition (all content crammed in upper third)
  * Text-icon size mismatch (heading 9pt next to a 100pt emoji)

Rules detect these by measuring per-card ratios and emit issues with
concrete `resize` / `nudge` mutate_argv. The agent reads the report,
looks at the render, and decides what to apply.

Output schema (each issue):
  {
    "category": "icon-undersized" | "icon-oversized" | "card-too-empty"
              | "icon-text-size-mismatch" | "top-heavy-composition",
    "severity": "info" | "warning",
    "shape_id": int,
    "shape_ids": [int, ...],   # for card-level issues
    "message": "...",
    "suggested_fix": "resize" | "nudge" | ...,
    "suggested_argv": [...],
    "proportion_report": {
      "card_area_emu2": int,
      "icon_area_emu2": int,
      "icon_pct": float,
      "empty_pct": float,
      "icon_size_pt_estimate": float,
      "max_text_size_pt": float,
      ...
    },
  }
"""
from __future__ import annotations

from statistics import median


# ---- design heuristics ----

# Icon size as a fraction of card area:
ICON_AREA_RATIO_TARGET = 0.12     # sweet spot: ~12% of card
ICON_AREA_RATIO_MIN = 0.05        # below this → icon undersized
ICON_AREA_RATIO_MAX = 0.30        # above this → icon oversized

# Card emptiness:
CARD_EMPTY_RATIO_THRESHOLD = 0.55  # > 55% empty area = uncomfortably sparse

# Text-icon size mismatch:
ICON_VS_TITLE_RATIO_MAX = 8.0     # icon-pt / title-pt > 8 = mismatch
ICON_VS_TITLE_RATIO_MIN = 1.5     # icon-pt should be > 1.5x title-pt

# Top-heavy: content_area_in_top_half / content_area_in_bottom_half
TOP_HEAVY_RATIO = 4.0

# Card identification:
CARD_MIN_W_EMU = 1500000     # ~1.6"
CARD_MIN_H_EMU = 1000000     # ~1.1"

# Icon text characters (single emoji or short label like "01")
ICON_TEXT_MAX_CHARS = 4
ICON_TEXT_MIN_PT_FOR_ICON = 36   # text font ≥ 36pt + ≤ 4 chars = treat as icon


# ---- helpers ----

def _bbox_contains_center(outer: dict, inner: dict, slack: int = 91440) -> bool:
    cx = inner["left"] + inner["width"] // 2
    cy = inner["top"] + inner["height"] // 2
    return (
        outer["left"] - slack <= cx <= outer["left"] + outer["width"] + slack
        and outer["top"] - slack <= cy <= outer["top"] + outer["height"] + slack
    )


def _is_card(obj: dict) -> bool:
    return (
        not obj.get("anomalous")
        and obj.get("kind") == "container"
        and obj.get("width", 0) >= CARD_MIN_W_EMU
        and obj.get("height", 0) >= CARD_MIN_H_EMU
        and obj.get("fill_hex")
    )


def _max_font_pt(obj: dict) -> float:
    sizes = obj.get("font_sizes") or []
    if not sizes:
        return 0.0
    return max(sizes) / 12700.0   # EMU per point


def _is_icon_like(obj: dict) -> bool:
    """Picture or single-glyph text used as an icon.

    For pictures + non-text shapes: aspect ratio gate (≤ 2.5:1) — long
    horizontal/vertical bars are decorations, not icons.

    For TEXT shapes containing a single emoji / glyph: aspect ratio is
    NOT a gate, because emoji-text-frames often inherit a wide text-box
    bbox even though the actual glyph is square. We trust the
    "1-4 char + ≥36pt font" signal instead.
    """
    w = obj.get("width", 0); h = obj.get("height", 0)
    if w <= 0 or h <= 0:
        return False

    if obj.get("kind") == "picture":
        # Square-ish picture only.
        if max(w, h) / max(min(w, h), 1) > 2.5:
            return False
        return True

    if obj.get("kind") == "text":
        text = (obj.get("text") or "").strip()
        if 1 <= len(text) <= ICON_TEXT_MAX_CHARS:
            if _max_font_pt(obj) >= ICON_TEXT_MIN_PT_FOR_ICON:
                return True

    # Other shape kinds (containers etc.) — apply aspect gate.
    if max(w, h) / max(min(w, h), 1) > 2.5:
        return False
    return False   # default: not an icon


def _icon_visual_area(icon: dict) -> int:
    """Effective visible glyph / image area.

    For text-emoji icons, the bbox is misleading (text frames are
    often wider than the glyph). Estimate visible glyph size from
    font_pt × 1.2 line-height proxy. For pictures, use the bbox.
    """
    if icon.get("kind") == "text":
        font_pt = _max_font_pt(icon)
        glyph_emu = int(font_pt * 12700 * 1.2)   # 1.2x line-height proxy
        return max(glyph_emu * glyph_emu, 1)
    return icon["width"] * icon["height"]


def _children_inside(card: dict, all_objects: list[dict]) -> list[dict]:
    return [
        o for o in all_objects
        if o is not card
        and not o.get("anomalous")
        and _bbox_contains_center(card, o)
    ]


# ---- per-issue detectors ----

def _detect_icon_size_issues(card: dict, children: list[dict]) -> list[dict]:
    """Icon too small / too big relative to card area.

    For text-emoji icons, the fix is `set-font-size` (changes glyph,
    keeps frame). For pictures, the fix is `resize`.
    """
    out = []
    card_area = card["width"] * card["height"]
    if card_area <= 0:
        return out
    icons = [c for c in children if _is_icon_like(c)]
    for icon in icons:
        icon_area = _icon_visual_area(icon)
        if icon_area <= 0:
            continue
        ratio = icon_area / card_area
        is_text = icon.get("kind") == "text"

        def _build_fix(category: str, scale: float) -> dict:
            if is_text:
                old_pt = _max_font_pt(icon)
                new_pt = max(round(old_pt * scale, 1), 12.0)
                return {
                    "suggested_fix": "set-font-size",
                    "suggested_argv": [
                        "set-font-size",
                        "--shape-id", str(icon["shape_id"]),
                        "--size-pt", str(new_pt),
                    ],
                    "old_pt": old_pt, "new_pt": new_pt,
                }
            new_w = int(icon["width"] * scale)
            new_h = int(icon["height"] * scale)
            return {
                "suggested_fix": "resize",
                "suggested_argv": [
                    "resize",
                    "--shape-id", str(icon["shape_id"]),
                    "--width", str(new_w),
                    "--height", str(new_h),
                ],
                "new_w": new_w, "new_h": new_h,
            }

        target_area = card_area * ICON_AREA_RATIO_TARGET
        scale = (target_area / icon_area) ** 0.5

        if ratio < ICON_AREA_RATIO_MIN:
            fix = _build_fix("icon-undersized", scale)
            out.append({
                "category": "icon-undersized",
                "severity": "warning",
                "shape_id": icon["shape_id"],
                "shape_ids": [card["shape_id"], icon["shape_id"]],
                "message": (
                    f"Icon #{icon['shape_id']} (visual area "
                    f"{round(ratio*100, 1)}% of card #{card['shape_id']}) "
                    f"is too small — target {ICON_AREA_RATIO_TARGET*100:.0f}%."
                ),
                "suggested_fix": fix["suggested_fix"],
                "suggested_argv": fix["suggested_argv"],
                "proportion_report": {
                    "card_area_emu2": card_area,
                    "icon_visual_area_emu2": icon_area,
                    "current_ratio_pct": round(ratio * 100, 2),
                    "target_ratio_pct": ICON_AREA_RATIO_TARGET * 100,
                    "scale_factor": round(scale, 2),
                    "icon_kind": icon.get("kind"),
                    **{k: v for k, v in fix.items()
                        if k not in ("suggested_fix", "suggested_argv")},
                },
            })
        elif ratio > ICON_AREA_RATIO_MAX:
            fix = _build_fix("icon-oversized", scale)
            out.append({
                "category": "icon-oversized",
                "severity": "warning",
                "shape_id": icon["shape_id"],
                "shape_ids": [card["shape_id"], icon["shape_id"]],
                "message": (
                    f"Icon #{icon['shape_id']} (visual area "
                    f"{round(ratio*100, 1)}% of card #{card['shape_id']}) "
                    f"is too large — target {ICON_AREA_RATIO_TARGET*100:.0f}%."
                ),
                "suggested_fix": fix["suggested_fix"],
                "suggested_argv": fix["suggested_argv"],
                "proportion_report": {
                    "card_area_emu2": card_area,
                    "icon_visual_area_emu2": icon_area,
                    "current_ratio_pct": round(ratio * 100, 2),
                    "target_ratio_pct": ICON_AREA_RATIO_TARGET * 100,
                    "scale_factor": round(scale, 2),
                    "icon_kind": icon.get("kind"),
                    **{k: v for k, v in fix.items()
                        if k not in ("suggested_fix", "suggested_argv")},
                },
            })
    return out


def _detect_card_emptiness(card: dict, children: list[dict]) -> list[dict]:
    """Card with too much empty space relative to content."""
    card_area = card["width"] * card["height"]
    if card_area <= 0:
        return []
    used = sum(
        c["width"] * c["height"]
        for c in children
        if c.get("kind") in ("container", "text", "picture", "shape")
        and not c.get("anomalous")
    )
    empty_ratio = max(0.0, 1.0 - used / card_area)
    if empty_ratio < CARD_EMPTY_RATIO_THRESHOLD:
        return []
    return [{
        "category": "card-too-empty",
        "severity": "info",
        "shape_id": card["shape_id"],
        "shape_ids": [card["shape_id"]],
        "message": (
            f"Card #{card['shape_id']} is {round(empty_ratio*100, 1)}% empty — "
            f"icons may be undersized or content sparse. Consider growing "
            f"the icon, adding visual elements, or shrinking the card."
        ),
        "suggested_fix": "review",
        "suggested_argv": None,
        "proportion_report": {
            "card_area_emu2": card_area,
            "used_area_emu2": used,
            "empty_pct": round(empty_ratio * 100, 1),
        },
    }]


def _detect_icon_text_mismatch(card: dict, children: list[dict]) -> list[dict]:
    """Icon font_pt vs title font_pt should be in [1.5x, 8x]."""
    icons = [c for c in children if _is_icon_like(c) and c.get("kind") == "text"]
    titles = [c for c in children
              if c.get("kind") == "text" and not _is_icon_like(c)
              and len(c.get("text") or "") >= 5]
    if not icons or not titles:
        return []
    out = []
    icon_pt = max((_max_font_pt(i) for i in icons), default=0)
    title_pt = max((_max_font_pt(t) for t in titles), default=0)
    if icon_pt <= 0 or title_pt <= 0:
        return []
    ratio = icon_pt / title_pt
    if ratio < ICON_VS_TITLE_RATIO_MIN or ratio > ICON_VS_TITLE_RATIO_MAX:
        biggest_icon = max(icons, key=_max_font_pt)
        target_pt = title_pt * 3.0   # icon should be ~3x title
        out.append({
            "category": "icon-text-size-mismatch",
            "severity": "info",
            "shape_id": biggest_icon["shape_id"],
            "shape_ids": [card["shape_id"], biggest_icon["shape_id"]],
            "message": (
                f"Icon-pt {round(icon_pt, 1)} vs title-pt {round(title_pt, 1)} "
                f"= {round(ratio, 1)}x ratio — typical balance is 1.5-8x. "
                f"Suggest icon target {round(target_pt, 1)}pt."
            ),
            "suggested_fix": "set-font-size",
            "suggested_argv": [
                "set-font-size",
                "--shape-id", str(biggest_icon["shape_id"]),
                "--size-pt", str(round(target_pt, 1)),
            ],
            "proportion_report": {
                "icon_pt": round(icon_pt, 1),
                "title_pt": round(title_pt, 1),
                "ratio": round(ratio, 2),
                "target_icon_pt": round(target_pt, 1),
            },
        })
    return out


def _detect_top_heavy(card: dict, children: list[dict]) -> list[dict]:
    """All content crammed in top half of card."""
    if not children:
        return []
    mid_y = card["top"] + card["height"] // 2
    top_area = bot_area = 0
    for c in children:
        if c.get("anomalous"):
            continue
        cy = c["top"] + c["height"] // 2
        a = c["width"] * c["height"]
        if cy < mid_y:
            top_area += a
        else:
            bot_area += a
    if bot_area == 0 and top_area == 0:
        return []
    ratio = top_area / max(bot_area, 1)
    if bot_area == 0 or ratio > TOP_HEAVY_RATIO:
        return [{
            "category": "top-heavy-composition",
            "severity": "info",
            "shape_id": card["shape_id"],
            "shape_ids": [card["shape_id"]],
            "message": (
                f"Card #{card['shape_id']} is top-heavy: top:bottom area "
                f"ratio = {round(ratio, 1) if bot_area else 'inf'}. "
                f"Move some content (icon?) into the lower half."
            ),
            "suggested_fix": "review",
            "suggested_argv": None,
            "proportion_report": {
                "top_area_emu2": top_area,
                "bottom_area_emu2": bot_area,
                "ratio": round(ratio, 2) if bot_area else None,
            },
        }]
    return []


# ---- top-level orchestrator ----

def detect_proportion_issues(slide: dict) -> list[dict]:
    """Run all proportion checks across every card on a slide."""
    objects = slide.get("objects", [])
    cards = [o for o in objects if _is_card(o)]
    issues: list[dict] = []
    for card in cards:
        children = _children_inside(card, objects)
        if not children:
            continue
        issues.extend(_detect_icon_size_issues(card, children))
        issues.extend(_detect_card_emptiness(card, children))
        issues.extend(_detect_icon_text_mismatch(card, children))
        issues.extend(_detect_top_heavy(card, children))
    return issues


def auto_apply_proportion_fixes(prs, slide_inspections: list[dict]) -> dict:
    """Walk every issue with a concrete suggested_argv and apply via mutate.

    Currently auto-applies only `resize` (icon-undersized / oversized)
    and `set-font-size` (icon-text-size-mismatch). Skips review-only
    suggestions.
    """
    from pptx.util import Emu, Pt
    from _shape_ops import set_size, set_font_size

    sid_index: dict[int, object] = {}
    for slide in prs.slides:
        for shape in slide.shapes:
            sid = getattr(shape, "shape_id", None)
            if sid is not None:
                sid_index[int(sid)] = shape

    actions: list[dict] = []
    for slide_data in slide_inspections:
        for issue in detect_proportion_issues(slide_data):
            argv = issue.get("suggested_argv")
            if not argv:
                continue
            if argv[0] == "resize":
                sid = int(argv[argv.index("--shape-id") + 1])
                w = int(argv[argv.index("--width") + 1])
                h = int(argv[argv.index("--height") + 1])
                shape = sid_index.get(sid)
                if shape is None:
                    continue
                # Keep center fixed: shift left/top by half the size delta.
                old_w = int(shape.width or 0); old_h = int(shape.height or 0)
                old_l = int(shape.left or 0); old_t = int(shape.top or 0)
                shape.left = Emu(old_l + (old_w - w) // 2)
                shape.top = Emu(old_t + (old_h - h) // 2)
                set_size(shape, width=w, height=h)
                actions.append({
                    "action": "resize-icon",
                    "category": issue["category"],
                    "shape_id": sid,
                    "from": [old_w, old_h],
                    "to": [w, h],
                })
            elif argv[0] == "set-font-size":
                sid = int(argv[argv.index("--shape-id") + 1])
                size_pt = float(argv[argv.index("--size-pt") + 1])
                shape = sid_index.get(sid)
                if shape is None:
                    continue
                set_font_size(shape, size_pt)
                actions.append({
                    "action": "set-font-size",
                    "category": issue["category"],
                    "shape_id": sid,
                    "size_pt": size_pt,
                })

    return {"applied": len(actions), "actions": actions}
