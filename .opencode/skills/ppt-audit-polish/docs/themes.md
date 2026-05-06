# Themes

A theme is a JSON file under `themes/` that defines:
- **palette** — named colors used everywhere downstream by role lookup
- **typography** — font family, per-role size/bold/color
- **spacing** — slide and card padding in EMU
- **decoration** — card fill/border/corner/shadow, badge fill, connector style

Validation schema: [_schema.json](../themes/_schema.json).

## Built-in themes

| File | Vibe | Notes |
|---|---|---|
| `clean-tech.json` | IBM-blue + neutral grays | Default. Best for product/tech architecture diagrams. |
| `academic-soft.json` | Slate-blue + warm beige surface | Best for research / paper-style decks. Lower contrast, more muted. |
| `business-warm.json` | Burnt-orange + cream | Larger title (30pt), shadowed cards. Best for client-facing pitches. |
| `editorial-dark.json` | Dark slate + sky-blue accents | Dark mode. Title 32pt. Use for code-heavy or all-figure decks. |

## Authoring a new theme

Copy `clean-tech.json` as a starting point. Required keys (validated):
- `palette` must include `text` and `background`
- `typography` must include `font_family`, `size_pt`, `bold`, `color_role`
- `spacing.slide_padding_emu` must be set

The polish engine looks up `color_role` strings in `palette`. Add new palette slots freely; just reference them from the relevant role.

### Role conventions used by `apply-typography` and recipes

| Role | Used for |
|---|---|
| `title` | Slide H1 |
| `subtitle` | One-line abstract under H1 |
| `h2` | Section/card header |
| `body` | Paragraph / list content |
| `caption` | Footnote / arrow label / small annotation |
| `badge` | Top-right corner pill |

### Font availability

Default font is `Microsoft YaHei` (works on Windows). For Linux deployments, switch to `Noto Sans CJK SC`, `WenQuanYi Micro Hei`, or `Source Han Sans SC`. For macOS, `PingFang SC` works.

The skill never embeds fonts; it sets the family name only. If the rendering machine doesn't have the font, LibreOffice will fall back, possibly with worse CJK kerning.

### Quick swap

```bash
python scripts/run_variants.py --input deck.pptx --work-dir out/ \
  --variants balanced  # only run the balanced recipe with default theme

python scripts/run_ppt_audit_polish.py --input deck.pptx --mode polish \
  --work-dir out/ --theme themes/business-warm.json
```
