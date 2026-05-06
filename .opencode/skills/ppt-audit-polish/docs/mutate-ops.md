# mutate.py - Op Catalog

**Total ops: 44**. Each takes `--in <pptx>` and `--out <pptx>`. Output is a JSON record on stdout.

> Auto-generated from `mutate.py list-ops --json`.

## geometry (11 ops)

| Op | Summary |
|---|---|
| `align` | Align multiple shapes to a common edge (left/right/top/bottom/center-h/center-v). |
| `center-on-slide` | Center a shape horizontally, vertically, or both. |
| `distribute` | Distribute shapes evenly along an axis (horizontal/vertical). |
| `equalize-gaps` | Set every shape's gap to a fixed value or auto-equalized one (axis horizontal|vertical). |
| `equalize-size` | Equalize widths and/or heights of multiple shapes to median. |
| `fit-to-slide` | Pull a shape inside slide bounds with the given padding (no-op if already inside). |
| `move` | Set absolute position (left/top in EMU). |
| `nudge` | Move shape by relative offset (dx/dy in EMU). |
| `resize` | Set absolute size (width/height in EMU). |
| `rotate` | Rotate a shape by absolute degrees. |
| `snap-to-grid` | Snap shape's position to the nearest grid step. |

### Examples

**align** - Align multiple shapes to a common edge (left/right/top/bottom/center-h/center-v).

```bash
python mutate align --in X --out Y --shape-ids 5,6,7 --edge left --target 502920
```

**center-on-slide** - Center a shape horizontally, vertically, or both.

```bash
python mutate center-on-slide --in X --out Y --shape-id 5 --axis horizontal
```

**distribute** - Distribute shapes evenly along an axis (horizontal/vertical).

```bash
python mutate distribute --in X --out Y --shape-ids 5,6,7,8 --axis horizontal
```

**equalize-gaps** - Set every shape's gap to a fixed value or auto-equalized one (axis horizontal|vertical).

```bash
python mutate equalize-gaps --in X --out Y --shape-ids 5,6,7,8 --axis horizontal
```

**equalize-size** - Equalize widths and/or heights of multiple shapes to median.

```bash
python mutate equalize-size --in X --out Y --shape-ids 5,6,7 --dimension both
```

**fit-to-slide** - Pull a shape inside slide bounds with the given padding (no-op if already inside).

```bash
python mutate fit-to-slide --in X --out Y --shape-id 5 --pad-emu 457200
```

**move** - Set absolute position (left/top in EMU).

```bash
python mutate move --in X --out Y --shape-id 5 --left 502920 --top 228600
```

**nudge** - Move shape by relative offset (dx/dy in EMU).

```bash
python mutate nudge --in X --out Y --shape-id 5 --dx 50800 --dy 0
```

**resize** - Set absolute size (width/height in EMU).

```bash
python mutate resize --in X --out Y --shape-id 5 --width 1828800 --height 228600
```

**rotate** - Rotate a shape by absolute degrees.

```bash
python mutate rotate --in X --out Y --shape-id 5 --degrees 0
```

**snap-to-grid** - Snap shape's position to the nearest grid step.

```bash
python mutate snap-to-grid --in X --out Y --shape-id 5 --grid-emu 91440
```

## style (10 ops)

| Op | Summary |
|---|---|
| `apply-badge-style` | Apply theme badge styling (primary fill + white text). |
| `apply-card-style` | Apply theme card styling (fill + border + corner) in one call. |
| `clear-fill` | Clear fill (transparent). |
| `clear-line` | Remove border. |
| `clear-shadow` | Remove drop shadow. |
| `set-corner-radius` | Round corners on a rounded-rectangle (ratio 0.0-0.5). |
| `set-fill` | Set solid fill color (hex RRGGBB). |
| `set-line` | Set border color and/or width (pt). |
| `set-opacity` | Set fill alpha (0.0-1.0). |
| `set-shadow` | Add an outer drop shadow. |

### Examples

**apply-badge-style** - Apply theme badge styling (primary fill + white text).

```bash
python mutate apply-badge-style --in X --out Y --shape-id 5 [--theme path]
```

**apply-card-style** - Apply theme card styling (fill + border + corner) in one call.

```bash
python mutate apply-card-style --in X --out Y --shape-id 5 [--theme path]
```

**clear-fill** - Clear fill (transparent).

```bash
python mutate clear-fill --in X --out Y --shape-id 5
```

**clear-line** - Remove border.

```bash
python mutate clear-line --in X --out Y --shape-id 5
```

**clear-shadow** - Remove drop shadow.

```bash
python mutate clear-shadow --in X --out Y --shape-id 5
```

**set-corner-radius** - Round corners on a rounded-rectangle (ratio 0.0-0.5).

```bash
python mutate set-corner-radius --in X --out Y --shape-id 5 --ratio 0.08
```

**set-fill** - Set solid fill color (hex RRGGBB).

```bash
python mutate set-fill --in X --out Y --shape-id 5 --color 0F62FE
```

**set-line** - Set border color and/or width (pt).

```bash
python mutate set-line --in X --out Y --shape-id 5 --color DDE1E6 --width-pt 0.75
```

**set-opacity** - Set fill alpha (0.0-1.0).

```bash
python mutate set-opacity --in X --out Y --shape-id 5 --alpha 0.6
```

**set-shadow** - Add an outer drop shadow.

```bash
python mutate set-shadow --in X --out Y --shape-id 5 --color 888888 --blur-pt 4 --dist-pt 2 --alpha 0.3
```

## z-order (4 ops)

| Op | Summary |
|---|---|
| `all-connectors-to-back` | Send every connector-like shape (lines, h=0 or w=0) to the back of every slide. |
| `bring-to-front` | Bring shape(s) to front. Shortcut for z-order --position front. |
| `send-to-back` | Send shape(s) to back. Shortcut for z-order --position back. |
| `z-order` | Move shape(s) back/front/up/down in z-order. |

### Examples

**all-connectors-to-back** - Send every connector-like shape (lines, h=0 or w=0) to the back of every slide.

```bash
python mutate all-connectors-to-back --in X --out Y
```

**bring-to-front** - Bring shape(s) to front. Shortcut for z-order --position front.

```bash
python mutate bring-to-front --in X --out Y --shape-ids 5
```

**send-to-back** - Send shape(s) to back. Shortcut for z-order --position back.

```bash
python mutate send-to-back --in X --out Y --shape-ids 10,11
```

**z-order** - Move shape(s) back/front/up/down in z-order.

```bash
python mutate z-order --in X --out Y --shape-ids 10,11 --position back
```

## text (11 ops)

| Op | Summary |
|---|---|
| `apply-typography` | Apply a theme typography role (size + bold + color + family) to a shape. |
| `set-font-bold` | Toggle bold. |
| `set-font-color` | Set font color (hex RRGGBB). |
| `set-font-family` | Set font family for one shape, all shapes (--scope all), or a role (--scope role:title). |
| `set-font-italic` | Toggle italic. |
| `set-font-size` | Set font size in pt for a shape. |
| `set-line-spacing` | Set line spacing ratio (1.0 = single). |
| `set-text` | Replace the text content of a shape. |
| `set-text-align` | Horizontal text alignment. |
| `set-text-margin` | Set text-frame inner margins (EMU each). |
| `set-text-v-align` | Vertical anchor inside the text frame. |

### Examples

**apply-typography** - Apply a theme typography role (size + bold + color + family) to a shape.

```bash
python mutate apply-typography --in X --out Y --shape-id 5 --role title [--theme path]
```

**set-font-bold** - Toggle bold.

```bash
python mutate set-font-bold --in X --out Y --shape-id 5 --bold true
```

**set-font-color** - Set font color (hex RRGGBB).

```bash
python mutate set-font-color --in X --out Y --shape-id 5 --color 161616
```

**set-font-family** - Set font family for one shape, all shapes (--scope all), or a role (--scope role:title).

```bash
python mutate set-font-family --in X --out Y --scope all --family "Microsoft YaHei"
```

**set-font-italic** - Toggle italic.

```bash
python mutate set-font-italic --in X --out Y --shape-id 5 --italic false
```

**set-font-size** - Set font size in pt for a shape.

```bash
python mutate set-font-size --in X --out Y --shape-id 5 --size-pt 14
```

**set-line-spacing** - Set line spacing ratio (1.0 = single).

```bash
python mutate set-line-spacing --in X --out Y --shape-id 5 --ratio 1.25
```

**set-text** - Replace the text content of a shape.

```bash
python mutate set-text --in X --out Y --shape-id 5 --content "New text"
```

**set-text-align** - Horizontal text alignment.

```bash
python mutate set-text-align --in X --out Y --shape-id 5 --align left
```

**set-text-margin** - Set text-frame inner margins (EMU each).

```bash
python mutate set-text-margin --in X --out Y --shape-id 5 --left-emu 91440 --right-emu 91440
```

**set-text-v-align** - Vertical anchor inside the text frame.

```bash
python mutate set-text-v-align --in X --out Y --shape-id 5 --anchor middle
```

## slide (5 ops)

| Op | Summary |
|---|---|
| `add-line` | Add a connector line between two points. |
| `add-rect` | Add a rectangle on a slide. |
| `add-text` | Add a text box. |
| `delete-shape` | Delete a shape from its slide. |
| `duplicate-shape` | Duplicate a shape (new copy appended to slide, returns new shape_id). |

### Examples

**add-line** - Add a connector line between two points.

```bash
python mutate add-line --in X --out Y --slide 1 --x1 0 --y1 0 --x2 914400 --y2 0 --color 6F6F6F --width-pt 1
```

**add-rect** - Add a rectangle on a slide.

```bash
python mutate add-rect --in X --out Y --slide 1 --left 0 --top 0 --width 914400 --height 91440 --fill 0F62FE
```

**add-text** - Add a text box.

```bash
python mutate add-text --in X --out Y --slide 1 --left 457200 --top 457200 --width 8000000 --height 400000 --content "Title"
```

**delete-shape** - Delete a shape from its slide.

```bash
python mutate delete-shape --in X --out Y --shape-id 33
```

**duplicate-shape** - Duplicate a shape (new copy appended to slide, returns new shape_id).

```bash
python mutate duplicate-shape --in X --out Y --shape-id 5
```

## connector (2 ops)

| Op | Summary |
|---|---|
| `straighten-connector` | Force connector to be horizontal (h=0) or vertical (w=0). |
| `style-connector` | Apply theme connector style (color + width) to one shape. |

### Examples

**straighten-connector** - Force connector to be horizontal (h=0) or vertical (w=0).

```bash
python mutate straighten-connector --in X --out Y --shape-id 10 --axis horizontal
```

**style-connector** - Apply theme connector style (color + width) to one shape.

```bash
python mutate style-connector --in X --out Y --shape-id 10 [--theme path]
```

## meta (1 ops)

| Op | Summary |
|---|---|
| `list-ops` | List every available op (use --json for machine-readable). |

### Examples

**list-ops** - List every available op (use --json for machine-readable).

```bash
python mutate list-ops --json
```
