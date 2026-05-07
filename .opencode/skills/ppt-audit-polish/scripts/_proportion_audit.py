"""Composition descriptor — purely descriptive, no judgments.

Earlier versions of this module hardcoded thresholds ("icon area must
be 8% of card", "card empty if < 45% filled") and auto-applied resizes.
That produced mechanical results — same emoji size on every deck, same
"correct" proportions regardless of the slide's communication goal.

This refactor flips the architecture: rules DESCRIBE the slide's
current composition (numbers + relationships); the agent reads
docs/DESIGN_PRINCIPLES.md, looks at the rendered slide, and JUDGES
what's wrong using its multimodal vision. Fixes are applied via
existing mutate ops (move / resize / set-font-size / align / ...).

Output schema (per slide):

  {
    "slide_dimensions": [width_emu, height_emu],
    "title_candidates": [
      {"shape_id": ..., "text": "...", "font_pt": 28, "bbox": [...]}
    ],
    "cards": [
      {
        "card_id": int,
        "bbox": [L, T, W, H],
        "fill_hex": "...",
        "area_pct_of_slide": 17.3,
        "children": [
          {
            "shape_id": ...,
            "kind": "text" | "picture" | "container",
            "role_hint": "icon" | "title" | "body" | "badge" | "decoration",
            "text_summary": "...",
            "bbox_in_card_pct": [x_pct, y_pct, w_pct, h_pct],
            "visible_area_pct_of_card": 12.3,   # for emoji uses glyph-pt²
            "font_pt": 18,
            "bold": true,
            "color_hex": "F5F5F5",
            "fill_hex": null,
          }
        ],
        "composition_summary": {
          "n_children": 8,
          "content_area_pct": 42.0,    # sum of children visible area / card area
          "empty_pct": 58.0,
          "vertical_balance": {        # how content distributes top vs bottom
            "top_half_area": 1234,
            "bottom_half_area": 567,
            "top_to_bottom_ratio": 2.18,
          },
          "horizontal_balance": {
            "left_half_area": ...,
            "right_half_area": ...,
            "left_to_right_ratio": ...,
          },
          "font_pt_levels": [9, 14, 28],   # distinct sizes used
          "fill_color_count": 4,
        },
      },
    ],
    "global_alignment": {
      "detected_column_lefts_emu": [457200, 6500000],
      "detected_row_tops_emu": [457200, 1500000, 4500000],
      "shapes_off_grid": [
        {"shape_id": ..., "off_by_emu": 50000, "axis": "top"}
      ],
    },
  }

The agent's job (per docs/DESIGN_PRINCIPLES.md):
  1. Read this descriptor + the rendered slide PNG
  2. Compare against the 5 design principles (clear conclusion,
     distinct partitions, strict alignment, visual hierarchy,
     information density)
  3. Decide what to change — using JUDGMENT, not thresholds
  4. Apply via mutate ops
"""
from __future__ import annotations

from statistics import median


# Card identification — keep "is this a content card vs slide-bg" logic
# (this is a structural fact, not an aesthetic threshold).
CARD_MIN_W_EMU = 1500000
CARD_MIN_H_EMU = 1000000
CARD_MAX_AREA_FRACTION = 0.80   # > 80% of slide → background, not card

# Role classification thresholds — also structural, not aesthetic
ICON_TEXT_MAX_CHARS = 4
ICON_TEXT_MIN_PT_FOR_ICON = 36
TITLE_TEXT_MIN_PT = 18
BADGE_MAX_PT = 16
BADGE_MAX_CHARS = 8


# ---- helpers ----

def _bbox_contains_center(outer: dict, inner: dict, slack: int = 91440) -> bool:
    cx = inner["left"] + inner["width"] // 2
    cy = inner["top"] + inner["height"] // 2
    return (
        outer["left"] - slack <= cx <= outer["left"] + outer["width"] + slack
        and outer["top"] - slack <= cy <= outer["top"] + outer["height"] + slack
    )


def _is_card(obj: dict, slide_w: int = 0, slide_h: int = 0) -> bool:
    if obj.get("anomalous"):
        return False
    if obj.get("kind") != "container":
        return False
    if not obj.get("fill_hex"):
        return False
    w = obj.get("width", 0); h = obj.get("height", 0)
    if w < CARD_MIN_W_EMU or h < CARD_MIN_H_EMU:
        return False
    if slide_w > 0 and slide_h > 0:
        slide_area = slide_w * slide_h
        if slide_area > 0 and (w * h) / slide_area > CARD_MAX_AREA_FRACTION:
            return False
    return True


def _max_font_pt(obj: dict) -> float:
    sizes = obj.get("font_sizes") or []
    if not sizes:
        return 0.0
    return max(sizes) / 12700.0


def _classify_role(child: dict, card: dict) -> str:
    """Best-effort role hint for the agent. NOT a final classification —
    the agent has visual judgment to override."""
    kind = child.get("kind")
    if kind == "picture":
        # Square-ish picture inside a card → likely icon
        w, h = child.get("width", 0), child.get("height", 0)
        if w > 0 and h > 0 and max(w, h) / max(min(w, h), 1) <= 2.5:
            return "icon"
        return "image"
    if kind == "text":
        text = (child.get("text") or "").strip()
        font_pt = _max_font_pt(child)
        if 1 <= len(text) <= ICON_TEXT_MAX_CHARS and font_pt >= ICON_TEXT_MIN_PT_FOR_ICON:
            return "icon"
        if 1 <= len(text) <= BADGE_MAX_CHARS and font_pt <= BADGE_MAX_PT:
            return "badge"
        if font_pt >= TITLE_TEXT_MIN_PT:
            return "title"
        return "body"
    if kind == "container":
        # Small filled container with text or no text = badge / decoration
        return "decoration"
    return "unknown"


def _icon_visible_area(child: dict) -> int:
    """For emoji-text icons, use glyph² (font_pt × 1.2)² instead of bbox.
    For other shapes, use bbox area."""
    if child.get("kind") == "text":
        font_pt = _max_font_pt(child)
        if font_pt > 0:
            glyph_emu = int(font_pt * 12700 * 1.2)
            return max(glyph_emu * glyph_emu, 1)
    return max(child.get("width", 0) * child.get("height", 0), 1)


def _children_inside(card: dict, all_objects: list[dict]) -> list[dict]:
    return [
        o for o in all_objects
        if o is not card
        and not o.get("anomalous")
        and _bbox_contains_center(card, o)
    ]


def _describe_child(child: dict, card: dict) -> dict:
    cw = max(card["width"], 1); ch = max(card["height"], 1)
    rel_x = (child["left"] - card["left"]) / cw
    rel_y = (child["top"] - card["top"]) / ch
    rel_w = child["width"] / cw
    rel_h = child["height"] / ch
    visible_area = _icon_visible_area(child)
    visible_pct = visible_area / max(cw * ch, 1) * 100
    text = (child.get("text") or "").strip()
    text_summary = text if len(text) <= 30 else text[:30] + "…"
    return {
        "shape_id": child["shape_id"],
        "kind": child.get("kind"),
        "role_hint": _classify_role(child, card),
        "text_summary": text_summary,
        "bbox_in_card_pct": [
            round(rel_x * 100, 1),
            round(rel_y * 100, 1),
            round(rel_w * 100, 1),
            round(rel_h * 100, 1),
        ],
        "visible_area_pct_of_card": round(visible_pct, 2),
        "font_pt": round(_max_font_pt(child), 1) if child.get("kind") == "text" else None,
        "color_hex": child.get("text_color"),
        "fill_hex": child.get("fill_hex"),
    }


def _card_summary(card: dict, children: list[dict]) -> dict:
    cw = max(card["width"], 1); ch = max(card["height"], 1)
    card_area = cw * ch
    mid_y = card["top"] + ch // 2
    mid_x = card["left"] + cw // 2
    top_a = bot_a = left_a = right_a = 0
    visible_total = 0
    for c in children:
        if c.get("anomalous"):
            continue
        a = _icon_visible_area(c) if _classify_role(c, card) == "icon" else c["width"] * c["height"]
        visible_total += a
        cy = c["top"] + c["height"] // 2
        cx = c["left"] + c["width"] // 2
        if cy < mid_y: top_a += a
        else: bot_a += a
        if cx < mid_x: left_a += a
        else: right_a += a
    content_pct = round(visible_total / card_area * 100, 1)

    pt_levels = sorted({
        round(_max_font_pt(c), 1)
        for c in children if c.get("kind") == "text" and _max_font_pt(c) > 0
    })
    fill_count = len({c.get("fill_hex") for c in children if c.get("fill_hex")})

    return {
        "n_children": len(children),
        "content_area_pct": content_pct,
        "empty_pct": round(max(0.0, 100.0 - content_pct), 1),
        "vertical_balance": {
            "top_half_area": top_a,
            "bottom_half_area": bot_a,
            "top_to_bottom_ratio": round(top_a / max(bot_a, 1), 2)
                                    if bot_a else None,
        },
        "horizontal_balance": {
            "left_half_area": left_a,
            "right_half_area": right_a,
            "left_to_right_ratio": round(left_a / max(right_a, 1), 2)
                                    if right_a else None,
        },
        "font_pt_levels": pt_levels,
        "fill_color_count": fill_count,
    }


def _detect_alignment_grid(objects: list[dict]) -> dict:
    """Cluster shape lefts/tops to find the slide's implicit grid lines."""
    lefts = [o["left"] for o in objects if not o.get("anomalous")]
    tops = [o["top"] for o in objects if not o.get("anomalous")]

    def cluster(values: list[int], eps: int = 100000) -> list[int]:
        if not values:
            return []
        ordered = sorted(values)
        clusters = [[ordered[0]]]
        for v in ordered[1:]:
            if v - clusters[-1][-1] <= eps:
                clusters[-1].append(v)
            else:
                clusters.append([v])
        # Only return clusters with ≥ 2 members (real grid lines)
        return [int(median(c)) for c in clusters if len(c) >= 2]

    cols = cluster(lefts)
    rows = cluster(tops)

    # Find shapes whose left/top is "close-but-not-on" any grid line
    EPS = 100000
    OFF_GRID_THRESHOLD = 30000   # nearer than this but not on = drift
    off_grid = []
    for o in objects:
        if o.get("anomalous"):
            continue
        sid = o.get("shape_id")
        if sid is None:
            continue
        for axis, val, anchors in (("left", o["left"], cols),
                                    ("top",  o["top"],  rows)):
            if not anchors:
                continue
            nearest = min(anchors, key=lambda a: abs(a - val))
            d = abs(nearest - val)
            if 0 < d < EPS:
                off_grid.append({
                    "shape_id": sid,
                    "axis": axis,
                    "actual_emu": val,
                    "nearest_grid_emu": nearest,
                    "off_by_emu": d,
                })
                break
    return {
        "detected_column_lefts_emu": cols,
        "detected_row_tops_emu": rows,
        "shapes_off_grid": off_grid[:20],   # cap for readability
    }


# ---- top-level ----

def describe_composition(slide: dict) -> dict:
    """Pure description of the slide's current composition.

    The agent reads this + the render PNG + DESIGN_PRINCIPLES.md and
    decides what (if anything) to change. NO hardcoded thresholds for
    "right/wrong" — only structural facts.
    """
    objects = slide.get("objects", [])
    sw = int(slide.get("width_emu", 0))
    sh = int(slide.get("height_emu", 0))
    cards = [o for o in objects if _is_card(o, sw, sh)]

    # Title candidates: large-pt text shapes in the upper third of the slide.
    # Exclude single-glyph emoji icons that happen to be in upper third.
    title_zone_bottom = sh // 3
    titles = []
    for o in objects:
        if o.get("kind") != "text" or o.get("anomalous"):
            continue
        text = (o.get("text") or "").strip()
        if not text:
            continue
        if o["top"] >= title_zone_bottom:
            continue
        if _max_font_pt(o) < TITLE_TEXT_MIN_PT:
            continue
        # Skip emoji / single-glyph icons
        if len(text) <= ICON_TEXT_MAX_CHARS \
                and _max_font_pt(o) >= ICON_TEXT_MIN_PT_FOR_ICON:
            continue
        titles.append({
            "shape_id": o["shape_id"],
            "text": text[:60],
            "font_pt": round(_max_font_pt(o), 1),
            "bbox_emu": [o["left"], o["top"], o["width"], o["height"]],
        })
    titles.sort(key=lambda t: -t["font_pt"])   # largest first

    card_descriptions = []
    for card in cards:
        children = _children_inside(card, objects)
        card_area_pct = round((card["width"] * card["height"]) / max(sw * sh, 1) * 100, 1)
        card_descriptions.append({
            "card_id": card["shape_id"],
            "bbox_emu": [card["left"], card["top"], card["width"], card["height"]],
            "fill_hex": card.get("fill_hex"),
            "area_pct_of_slide": card_area_pct,
            "children": [_describe_child(c, card) for c in children],
            "composition_summary": _card_summary(card, children),
        })

    return {
        "slide_dimensions": [sw, sh],
        "title_candidates": titles,
        "cards": card_descriptions,
        "global_alignment": _detect_alignment_grid(objects),
    }
