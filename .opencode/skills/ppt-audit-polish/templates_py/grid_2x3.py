"""Grid 2x3 (or 2x2 / 2x4) — items arranged in a 2-column grid of cards.

Best fit for 4-8 items where horizontal-timeline would feel too narrow per
card.
"""
from __future__ import annotations

from _base import (
    add_blank_slide, add_rect, add_rounded_rect, add_text,
    truncate, wipe_slides,
)


NAME = "grid-2x3"
DESCRIPTION = "Title above a 2-column grid of N rounded cards (2x2 / 2x3 / 2x4)"
MIN_ITEMS = 4
MAX_ITEMS = 8
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
    border = theme.color("border")
    family = theme.font_family() or "Microsoft YaHei"

    # Title row.
    title = content.get("title") or "Slide Title"
    add_text(slide, 457200, 365760, SW - 914400, 411480, title,
             size_pt=24, bold=True, color=text_strong, family=family,
             align="left", v_align="middle")

    subtitle = content.get("subtitle") or ""
    if subtitle:
        add_text(slide, 457200, 822960, SW - 914400, 274320, subtitle,
                 size_pt=11, color=text_muted, family=family,
                 align="left", v_align="middle")

    # Decorative accent bar under title.
    add_rect(slide, 457200, 1188720, 685800, 38100, fill=primary, line="none")

    # Grid.
    items = (content.get("items") or [])[:MAX_ITEMS]
    n = len(items)
    warnings: list[str] = []
    if n == 0:
        warnings.append("no items found")
        return {"warnings": warnings, "items_rendered": 0}
    if n < MIN_ITEMS:
        warnings.append(f"only {n} item(s) — template prefers >= {MIN_ITEMS}")

    cols = 2 if n <= 8 else 3
    rows = (n + cols - 1) // cols

    grid_top = 1402080
    grid_left = 457200
    grid_right_pad = 457200
    grid_bottom_pad = 685800

    grid_w = SW - grid_left - grid_right_pad
    grid_h = SH - grid_top - grid_bottom_pad
    cell_gap = 274320
    cell_w = (grid_w - (cols - 1) * cell_gap) // cols
    cell_h = (grid_h - (rows - 1) * cell_gap) // rows

    for idx, item in enumerate(items):
        col = idx % cols
        row = idx // cols
        x = grid_left + col * (cell_w + cell_gap)
        y = grid_top + row * (cell_h + cell_gap)

        # Card background.
        add_rounded_rect(slide, x, y, cell_w, cell_h,
                         fill=background, line=border, line_pt=0.75, corner_ratio=0.04)

        # Soft accent strip on the left edge of each card.
        add_rect(slide, x, y, 91440, cell_h, fill=primary, line="none")

        # Number badge (top-left of card body).
        pad = 274320
        badge_size = 365760
        add_rounded_rect(slide, x + pad, y + pad, badge_size, badge_size,
                         fill=primary_soft, line="none", corner_ratio=0.3)
        add_text(slide, x + pad, y + pad, badge_size, badge_size, f"{idx + 1:02d}",
                 size_pt=14, bold=True, color=primary, family=family,
                 align="center", v_align="middle")

        # Name.
        name_x = x + pad + badge_size + 182880
        name_w = cell_w - (name_x - x) - pad
        name = truncate(item.get("name") or f"Item {idx + 1}", 28)
        add_text(slide, name_x, y + pad, name_w, badge_size, name,
                 size_pt=15, bold=True, color=text_strong, family=family,
                 align="left", v_align="middle")

        # Body description below.
        body_y = y + pad + badge_size + 182880
        body_h = cell_h - (body_y - y) - pad
        description = item.get("description") or ""
        extra = item.get("extra") or ""
        body = (extra + "\n" + description).strip() if extra else description
        body = truncate(body, 240)
        add_text(slide, x + pad, body_y, cell_w - 2 * pad, body_h, body,
                 size_pt=10, color=text_muted, family=family,
                 align="left", v_align="top")

        # Detail chips (max 2)
        details = (item.get("details") or [])[:2]
        if details:
            chip_y = y + cell_h - 320040 - pad // 2
            chip_x = x + pad
            for chip_text in details:
                chip_w = max(int(len(chip_text) * 91440 * 0.85 + 182880), 480000)
                add_rounded_rect(slide, chip_x, chip_y, chip_w, 256032,
                                 fill=primary_soft, line="none", corner_ratio=0.4)
                add_text(slide, chip_x, chip_y, chip_w, 256032, chip_text,
                         size_pt=9, bold=True, color=primary, family=family,
                         align="center", v_align="middle")
                chip_x += chip_w + 91440

    return {"warnings": warnings, "items_rendered": n}
