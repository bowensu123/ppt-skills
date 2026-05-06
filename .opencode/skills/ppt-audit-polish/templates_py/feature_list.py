"""Feature-list — left vertical list of N items, right side is a hero panel
with the title/subtitle/footer.

Best fit when items are independent features/benefits (3-6) and you want a
strong title visual on one side.
"""
from __future__ import annotations

from _base import (
    add_blank_slide, add_circle, add_rect, add_rounded_rect, add_text,
    truncate, wipe_slides,
)


NAME = "feature-list"
DESCRIPTION = "Left vertical list of items, right hero panel with title/subtitle"
MIN_ITEMS = 3
MAX_ITEMS = 7
PREFERRED_THEME = "clean-tech"


def render(prs, content: dict, theme) -> dict:
    SW = int(prs.slide_width)
    SH = int(prs.slide_height)

    wipe_slides(prs)
    slide = add_blank_slide(prs)

    primary = theme.color("primary")
    primary_soft = theme.color("primary_soft") if theme.palette.get("primary_soft") else "EAF1FF"
    text_strong = theme.color("text_strong")
    text_muted = theme.color("text_muted")
    background = theme.color("background")
    family = theme.font_family() or "Microsoft YaHei"

    # 1. Right hero panel (40% width, full height).
    hero_left = int(SW * 0.55)
    hero_w = SW - hero_left
    add_rect(slide, hero_left, 0, hero_w, SH, fill=primary_soft, line="none")

    # 2. Hero title.
    title = content.get("title") or "Slide Title"
    add_text(slide, hero_left + 457200, 822960, hero_w - 914400, 1280160,
             truncate(title, 60),
             size_pt=28, bold=True, color=text_strong, family=family,
             align="left", v_align="top")

    subtitle = content.get("subtitle") or ""
    if subtitle:
        add_text(slide, hero_left + 457200, 2148840, hero_w - 914400, 1280160,
                 truncate(subtitle, 220),
                 size_pt=12, color=text_muted, family=family,
                 align="left", v_align="top")

    # 3. Hero accent shape (a quarter-circle decorative element, simulated by overlapping rect+circle).
    add_circle(slide, hero_left + hero_w - 685800, SH - 685800, 411480,
               fill=primary, line="none")

    # 4. Footer text in hero.
    footer = content.get("footer") or ""
    if footer:
        add_text(slide, hero_left + 457200, SH - 1097280, hero_w - 914400, 685800,
                 truncate(footer, 180),
                 size_pt=10, color=text_strong, family=family,
                 align="left", v_align="top")

    # 5. Left list of items.
    items = (content.get("items") or [])[:MAX_ITEMS]
    n = len(items)
    warnings: list[str] = []
    if n == 0:
        warnings.append("no items found")
        return {"warnings": warnings, "items_rendered": 0}
    if n < MIN_ITEMS:
        warnings.append(f"only {n} item(s) — template prefers >= {MIN_ITEMS}")

    list_top = 685800
    list_left = 457200
    list_right = hero_left - 274320
    list_h = SH - list_top - 457200
    item_gap = 182880
    item_h = (list_h - (n - 1) * item_gap) // n

    for i, item in enumerate(items):
        y = list_top + i * (item_h + item_gap)

        # Number circle on the left.
        circle_cx = list_left + 274320
        circle_cy = y + item_h // 2
        circle_r = 274320
        add_circle(slide, circle_cx, circle_cy, circle_r, fill=primary, line="none")
        add_text(slide, circle_cx - circle_r, circle_cy - circle_r, circle_r * 2, circle_r * 2,
                 f"{i + 1}",
                 size_pt=14, bold=True, color=background, family=family,
                 align="center", v_align="middle")

        # Item name + description to the right.
        text_x = list_left + circle_r * 2 + 274320
        text_w = list_right - text_x
        name = truncate(item.get("name") or f"Item {i + 1}", 32)
        add_text(slide, text_x, y, text_w, item_h * 4 // 10, name,
                 size_pt=14, bold=True, color=text_strong, family=family,
                 align="left", v_align="middle")

        description = item.get("description") or ""
        extra = item.get("extra") or ""
        body = (extra + " — " + description).strip(" —") if extra else description
        body = truncate(body, 140)
        add_text(slide, text_x, y + item_h * 4 // 10, text_w, item_h * 6 // 10, body,
                 size_pt=10, color=text_muted, family=family,
                 align="left", v_align="top")

    return {"warnings": warnings, "items_rendered": n}
