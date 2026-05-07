"""Claude-Code-aesthetic template — terminal-styled grid for code/dev/AI content.

Visual identity:
  * Dark near-black background (#0F0F0F), card surfaces slightly lighter
  * Anthropic coral primary (#D97757) for accents, prompts, dividers
  * Monospace font (Consolas / JetBrains Mono / Cascadia Code) for that
    terminal feel; falls back gracefully when unavailable
  * Each item rendered as a "code block card" with a coral chevron
    prefix + numeric badge in the top-left corner
  * Title prefixed with a `$` prompt; subtitle styled as `// comment`
  * Footer rendered as a terminal status line with `$` prefix

Best fit for: dev tools, AI products, code/agent showcases, technical
roadmaps. Pairs with the `claude-code` theme JSON for the palette.
Works with any theme but looks best on dark palettes.
"""
from __future__ import annotations

from _base import (
    add_blank_slide, add_image_from_path, add_rect, add_rounded_rect,
    add_text, aspect_fit_box, truncate, wipe_slides,
)


NAME = "claude-code"
DESCRIPTION = "Terminal-styled dark grid with coral accents and monospace type — for dev/AI/code content"
MIN_ITEMS = 3
MAX_ITEMS = 8
PREFERRED_THEME = "claude-code"


def render(prs, content: dict, theme) -> dict:
    SW = int(prs.slide_width)
    SH = int(prs.slide_height)

    wipe_slides(prs)
    slide = add_blank_slide(prs)

    # Pull palette with fallbacks. The template is most distinctive on
    # dark themes but usable on light ones too.
    primary = theme.color("primary")
    text_strong = theme.color("text_strong")
    text = theme.color("text")
    text_muted = theme.color("text_muted")
    background = theme.color("background")
    surface = theme.color("surface") if theme.palette.get("surface") else background
    border = theme.color("border")
    family = theme.font_family() or "Consolas"

    # 1. Full-bleed dark background to set the terminal mood. We DRAW
    #    a rect rather than relying on the slide background so the
    #    template is portable across master slides.
    add_rect(slide, 0, 0, SW, SH, fill=background, line="none")

    # 2. Top coral accent bar (1px-ish) — terminal "title-bar separator"
    add_rect(slide, 0, 0, SW, 60960, fill=primary, line="none")

    # 3. Title row: "$ <title>"
    margin = 457200
    title_y = 365760
    title_text = content.get("title") or "Untitled"
    # Coral `$` prefix in its own textbox so we can color it independently.
    prompt_w = 274320
    add_text(slide, margin, title_y, prompt_w, 411480, "$",
             size_pt=24, bold=True, color=primary, family=family,
             align="left", v_align="middle")
    add_text(slide, margin + prompt_w, title_y, SW - 2 * margin - prompt_w, 411480,
             title_text,
             size_pt=24, bold=True, color=text_strong, family=family,
             align="left", v_align="middle")

    # 4. Subtitle as "// comment"
    sub_y = title_y + 411480
    subtitle = content.get("subtitle") or ""
    if subtitle:
        add_text(slide, margin, sub_y, SW - 2 * margin, 274320,
                 f"// {subtitle}",
                 size_pt=12, italic=False, color=text_muted, family=family,
                 align="left", v_align="middle")
        sub_y += 274320

    # 5. Coral divider line under the title block.
    div_y = sub_y + 91440
    add_rect(slide, margin, div_y, SW - 2 * margin, 13716,
             fill=primary, line="none")

    # 6. Items grid.
    items = (content.get("items") or [])[:MAX_ITEMS]
    n = len(items)
    warnings: list[str] = []
    if n == 0:
        warnings.append("no items found")
        return {"warnings": warnings, "items_rendered": 0}
    if n < MIN_ITEMS:
        warnings.append(f"only {n} item(s) — template prefers >= {MIN_ITEMS}")

    # Layout: 2 columns when n>=4, 1 column otherwise (single-col reads
    # like a vertical command log which fits 3-or-fewer items well).
    cols = 2 if n >= 4 else 1
    rows = (n + cols - 1) // cols

    grid_top = div_y + 274320
    grid_left = margin
    footer_pad = 685800
    grid_h = SH - grid_top - footer_pad
    grid_w = SW - 2 * margin
    cell_gap = 228600
    cell_w = (grid_w - (cols - 1) * cell_gap) // cols
    cell_h = (grid_h - (rows - 1) * cell_gap) // rows

    for idx, item in enumerate(items):
        col = idx % cols
        row = idx // cols
        x = grid_left + col * (cell_w + cell_gap)
        y = grid_top + row * (cell_h + cell_gap)

        # Card surface with subtle border (no shadow — terminal cards
        # have hard edges).
        add_rounded_rect(
            slide, x, y, cell_w, cell_h,
            fill=surface, line=border, line_pt=0.75, corner_ratio=0.03,
        )

        # Left coral edge — looks like a focused-pane indicator.
        add_rect(slide, x, y, 36576, cell_h, fill=primary, line="none")

        pad = 228600
        # Top-left header: image if attributed, else "▸ NN" chevron + number
        chev_w = 137160
        num_w = 365760
        header_y = y + pad
        item_image = item.get("image")
        if item_image and item_image.get("path"):
            # Replace the chevron+number with a small icon (square fits the
            # header band height ~274320 EMU). Aspect-fit centered.
            icon_box = 274320
            ox, oy, fw, fh = aspect_fit_box(icon_box, icon_box, icon_box, icon_box)
            add_image_from_path(slide, item_image["path"],
                                x + pad + ox, header_y + oy, fw, fh)
            name_x = x + pad + icon_box + 91440
        else:
            add_text(slide, x + pad, header_y, chev_w, 274320, "▸",
                     size_pt=12, bold=True, color=primary, family=family,
                     align="left", v_align="middle")
            add_text(slide, x + pad + chev_w, header_y, num_w, 274320,
                     f"{idx + 1:02d}",
                     size_pt=11, bold=True, color=primary, family=family,
                     align="left", v_align="middle")
            name_x = x + pad + chev_w + num_w
        name_w = cell_w - (name_x - x) - pad
        name = truncate(item.get("name") or f"item-{idx + 1:02d}", 28)
        add_text(slide, name_x, header_y, name_w, 274320, name,
                 size_pt=14, bold=True, color=text_strong, family=family,
                 align="left", v_align="middle")

        # Subtle divider beneath the header.
        divider_y = header_y + 297180
        add_rect(slide, x + pad, divider_y, cell_w - 2 * pad, 9144,
                 fill=border, line="none")

        # Body description (smaller, muted).
        body_y = divider_y + 91440
        body_h = cell_h - (body_y - y) - pad
        description = item.get("description") or ""
        extra = item.get("extra") or ""
        if extra:
            description = (extra + "\n" + description).strip()
        body = truncate(description, 220)
        add_text(slide, x + pad, body_y, cell_w - 2 * pad, body_h, body,
                 size_pt=10, color=text, family=family,
                 align="left", v_align="top")

        # Detail chips at the bottom rendered as "// detail" comments
        details = (item.get("details") or [])[:2]
        if details:
            chip_y = y + cell_h - 274320 - pad // 2
            chip_text = "  ".join(f"// {d}" for d in details)
            add_text(slide, x + pad, chip_y, cell_w - 2 * pad, 228600,
                     truncate(chip_text, 80),
                     size_pt=9, color=text_muted, family=family,
                     align="left", v_align="middle")

    # 7. Footer status line: "$ status: <footer>"
    footer = content.get("footer") or ""
    footer_secondary = content.get("footer_secondary") or ""
    footer_combined = " · ".join(p for p in (footer, footer_secondary) if p)
    if footer_combined:
        footer_y = SH - footer_pad + 91440
        # Coral `$` prefix
        add_text(slide, margin, footer_y, prompt_w, 274320, "$",
                 size_pt=11, bold=True, color=primary, family=family,
                 align="left", v_align="middle")
        add_text(slide, margin + prompt_w, footer_y,
                 SW - 2 * margin - prompt_w, 274320,
                 truncate(footer_combined, 140),
                 size_pt=10, color=text_muted, family=family,
                 align="left", v_align="middle")

    # Slide-level decorations (logos / chrome): render at original coords.
    for deco in (content.get("decorations") or []):
        path = deco.get("path")
        if not path:
            continue
        add_image_from_path(slide, path,
                            int(deco.get("left", 0)),
                            int(deco.get("top", 0)),
                            int(deco.get("width", 0)),
                            int(deco.get("height", 0)))

    return {"warnings": warnings, "items_rendered": n}
