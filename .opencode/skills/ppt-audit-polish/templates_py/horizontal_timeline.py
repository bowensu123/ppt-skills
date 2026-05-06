"""Horizontal-timeline template — top accent bar, title, subtitle band,
horizontal connector line with N dots/numbered cards beneath.

Best fit when content has 3-7 peer items (process steps, comparison classes,
stages of a flow).
"""
from __future__ import annotations

from _base import (
    add_blank_slide, add_circle, add_rect, add_rounded_rect, add_text,
    horizontal_distribute, truncate, wipe_slides,
)


NAME = "horizontal-timeline"
DESCRIPTION = "Top accent bar + title + subtitle, then a horizontal connector with N numbered cards"
MIN_ITEMS = 3
MAX_ITEMS = 7
PREFERRED_THEME = "business-warm"


def render(prs, content: dict, theme) -> dict:
    """Returns {"warnings": [...], "items_rendered": N}."""
    SW = int(prs.slide_width)
    SH = int(prs.slide_height)

    wipe_slides(prs)
    slide = add_blank_slide(prs)

    primary = theme.color("primary")
    text_strong = theme.color("text_strong")
    text_muted = theme.color("text_muted")
    background = theme.color("background")
    surface = theme.color("surface") if theme.palette.get("surface") else background
    border = theme.color("border")

    family = theme.font_family() or "Microsoft YaHei"

    # 1. Top accent bar.
    add_rect(slide, 0, 0, SW, 365760, fill=primary, line="none")

    # 2. Header text in the accent bar (right-aligned brand line).
    badge = (content.get("badge") or "").strip()
    if badge:
        add_text(slide, SW - 3000000, 91440, 2700000, 200000, badge,
                 size_pt=11, bold=True, color=background, family=family,
                 align="right", v_align="middle")

    # 3. Title beneath the bar.
    title = content.get("title") or "Slide Title"
    add_text(slide, 457200, 487680, SW - 914400, 365760, title,
             size_pt=24, bold=True, color=text_strong, family=family,
             align="left", v_align="middle")

    # 4. Subtitle.
    subtitle = content.get("subtitle") or ""
    if subtitle:
        add_text(slide, 457200, 914400, SW - 914400, 320040, subtitle,
                 size_pt=12, color=text_muted, family=family,
                 align="left", v_align="middle")

    # 5. Timeline horizontal connector.
    timeline_y = 1539240
    timeline_left = 853440
    timeline_right = SW - 853440
    add_rect(slide, timeline_left, timeline_y, timeline_right - timeline_left, 22860,
             fill=primary, line="none")

    # 6. N evenly-distributed numbered cards.
    items = content.get("items") or []
    items = items[:MAX_ITEMS] if items else []
    n = len(items)
    warnings: list[str] = []
    if n == 0:
        warnings.append("no items found in content; rendered title bar only")
        return {"warnings": warnings, "items_rendered": 0}
    if n < MIN_ITEMS:
        warnings.append(f"only {n} item(s) — template prefers >= {MIN_ITEMS}")

    margin = 685800
    available = SW - 2 * margin
    slot_w = available // n
    card_w = int(slot_w * 0.85)
    card_x_centers = [margin + i * slot_w + slot_w // 2 for i in range(n)]

    # Dots on the timeline, then numbered card below.
    dot_radius = 91440
    card_top = 1828800
    card_height = SH - card_top - 1280160  # leave room for footer
    number_height = 320040
    name_height = 365760
    body_height = card_height - number_height - name_height - 200000
    if body_height < 600000:
        body_height = 600000

    for i, (cx, item) in enumerate(zip(card_x_centers, items)):
        add_circle(slide, cx, timeline_y + 11430, dot_radius, fill=primary, line="none")

        # Number "01", "02", ...
        num_text = f"{i + 1:02d}"
        add_text(slide, cx - card_w // 2, card_top, card_w, number_height, num_text,
                 size_pt=18, bold=True, color=primary, family=family,
                 align="left", v_align="top")

        # Name (e.g., "Zero-shot")
        name = truncate(item.get("name") or f"Step {i + 1}", 24)
        add_text(slide, cx - card_w // 2, card_top + number_height,
                 card_w, name_height, name,
                 size_pt=18, bold=True, color=text_strong, family=family,
                 align="left", v_align="top")

        # Description body (concise paragraph)
        description = item.get("description") or ""
        extra = item.get("extra") or ""
        body_text = (extra + "\n" + description).strip() if extra else description
        body_text = truncate(body_text, 180)
        add_text(slide, cx - card_w // 2, card_top + number_height + name_height,
                 card_w, body_height, body_text,
                 size_pt=10, color=text_muted, family=family,
                 align="left", v_align="top")

        # Detail labels (small chips at bottom of card slot)
        details = (item.get("details") or [])[:3]
        if details:
            chip_y = card_top + number_height + name_height + body_height + 91440
            chip_h = 240000
            chip_x = cx - card_w // 2
            for chip_text in details:
                chip_w = max(int(len(chip_text) * 91440 * 0.8 + 182880), 480000)
                add_rounded_rect(slide, chip_x, chip_y, chip_w, chip_h,
                                 fill=primary, line="none", corner_ratio=0.4)
                add_text(slide, chip_x, chip_y, chip_w, chip_h, chip_text,
                         size_pt=8, bold=True, color=background, family=family,
                         align="center", v_align="middle")
                chip_x += chip_w + 91440
                if chip_x > cx + card_w // 2:
                    break

    # 7. Footer band.
    footer = content.get("footer") or ""
    footer_secondary = content.get("footer_secondary") or ""
    footer_top = SH - 822960
    add_rect(slide, 0, footer_top, SW, 6096, fill=border, line="none")
    if footer:
        add_text(slide, 457200, footer_top + 100000, SW - 914400, 280000,
                 truncate(footer, 200),
                 size_pt=10, color=text_strong, family=family,
                 align="left", v_align="top")
    if footer_secondary:
        add_text(slide, 457200, footer_top + 400000, SW - 914400, 280000,
                 truncate(footer_secondary, 200),
                 size_pt=9, color=text_muted, family=family,
                 align="left", v_align="top")

    return {"warnings": warnings, "items_rendered": n}
