---
name: ppt-audit-polish
description: Audit, polish, and re-template PowerPoint decks. The OpenCode agent drives the loop using its own multimodal model — inspect, render, look, mutate or regenerate. Two diagnostic modes (audit, polish via mutate) plus three template-based regenerate modes for decks whose structural ceiling is too low to fix with mutation alone.
compatibility: opencode
metadata:
  audience: presentation-authors
  workflow: ppt-review
  ops_count: 45
  themes_count: 4
  templates_count: 3
---

## When to invoke

Trigger on any of: "美化 PPT", "polish this deck", "audit slides", "fix layout", "调整版面", "边看边改", "迭代美化", "把 X 改成 Y".

The user is asking YOU (the OpenCode agent) to drive an iterative polish loop. You use your own multimodal model — no `PPT_POLISH_MODEL_*` env vars, no separate API setup.

## Two paths the agent picks between

```
Path A: POLISH       — keep existing layout, repair structural bugs and apply theme.
                       Best when score baseline > 30 OR primary issues are
                       structural (peer-card outliers, overlap, alignment drift).
                       Output: <stem>.polished.pptx

Path B: REGENERATE   — extract content, drop into a fresh template.
                       Best when score baseline < 25 AND density/hierarchy are
                       very low (deck design itself is the problem, not bugs).
                       Output: <stem>.regen-<template>.pptx
```

### Decision rule (in STEP A of the loop)

After the FIRST `state_summary.py` call:

```
if baseline_score >= 30 OR issue_categories are mostly structural:
    → Path A (polish loop, see below)

elif baseline_score < 25 AND (density_score < 30 OR hierarchy_score < 50):
    → Suggest Path B to the user:
      "This deck's design ceiling is low — typography, spacing, and palette
       are all far from any clean target. Polish can fix structural bugs but
       the visual will still look cluttered. Want me to regenerate it with a
       template instead? Available: horizontal-timeline / grid-2x3 / feature-list."
    → If user says yes: Path B (regenerate)
    → If user says no: Path A (do what we can)
```

## Path A: Polish (one mutate per iteration)

### One key primitive: `state_summary.py`

```bash
python scripts/state_summary.py --in <pptx> --work-dir <state-dir>
```

Returns ALL of {score, top_issues, render PNG path, key_shapes by role, next_step_hints with ready-to-run mutate argv, diff vs previous iteration} in one JSON dump. Use it for every iteration so the loop stays cheap.

### Playbook

```
INIT
  python scripts/state_summary.py --in input.pptx --work-dir state-0/
  → read state-0/state-summary.json (score, top_issues, next_step_hints)
  → read state-0/render/slide-001.png with the Read tool — SEE the slide
  → record baseline_score from the JSON

LOOP (cap at 8 iterations or score plateau)
  STEP A — pick the next fix
    Look at the rendered image. Read top_issues + next_step_hints.
    Decide ONE concrete change. Each hint already has its mutate_argv ready.
    You may also pick something the hints didn't list — you can see things
    metrics can't.

  STEP B — apply
    python scripts/mutate.py <argv> --in current.pptx --out iter-N.pptx

  STEP C — observe
    python scripts/state_summary.py --in iter-N.pptx --work-dir state-N/ \
        --diff-from current.pptx
    → read state-N/state-summary.json (new score + diff block)
    → read state-N/render/slide-001.png  (SEE the result)

  STEP D — judge with your own eyes + the score (MANDATORY 3-part report)
    Every iteration MUST produce a report in EXACTLY this format:

    ▸ NUMERIC (quote `agent_report.numeric_block` from state-summary.json verbatim)
        score: X.X → Y.Y (Δ ZZZ)
          alignment: A1 → A2 (Δ a)
          density:   D1 → D2 (Δ d)
          hierarchy: H1 → H2 (Δ h)
          palette:   P1 → P2 (Δ p)
          issues:    Δ N
          verdict:   "fair → good"

    ▸ VISUAL (you read render-N PNG and describe in ONE sentence what you SEE)
        Examples:
          "Few-shot 卡片图标和'少量示例'文字现在落在卡片内部了"
          "卡片 4 头标的青色色块缩到了正确尺寸"
          "标题字号变大但文字开始溢出文本框"

    ▸ DECISION (ADOPT or REJECT, with one-sentence justification combining numeric + visual)
        ADOPT: when numeric Δ ≥ +0.5 AND visual change is positive or neutral
        REJECT: when numeric Δ ≤ -0.5 OR visual introduced new problems

    The state-summary's `score_delta.decision_recommendation` field gives a pre-computed
    suggestion (adopt/reject/neutral) — feel free to override with your visual judgment.

  STEP E — decide stop
    If best score hasn't moved in last 2 iterations AND nothing looks broken: STOP.
    Otherwise: back to STEP A.

OUTPUT
  Copy the best iter-N.pptx to <work-dir>/<stem>.polished.pptx
  Tell the user: final score, what was fixed (cite the diff blocks), file path.
```

## Path B: Regenerate (template-based)

When polish ceiling is too low, regenerate produces a fresh slide using a pre-built template.

### One-shot regeneration

```bash
# Auto-pick template based on item count
python scripts/regenerate.py --in deck.pptx --work-dir out/ --auto

# Explicit template choice
python scripts/regenerate.py --in deck.pptx --work-dir out/ --template horizontal-timeline
python scripts/regenerate.py --in deck.pptx --work-dir out/ --template grid-2x3
python scripts/regenerate.py --in deck.pptx --work-dir out/ --template feature-list

# With theme override
python scripts/regenerate.py --in deck.pptx --work-dir out/ --auto --theme themes/business-warm.json
```

### Available templates

| Template | Items | Layout | Best for |
|---|---|---|---|
| `horizontal-timeline` | 3-7 | Top accent bar + connector + N numbered cards | Process steps, comparison classes, stages |
| `grid-2x3` | 4-8 | 2-col grid of rounded cards | Independent feature/benefit lists |
| `feature-list` | 3-7 | Right hero panel + left vertical list | When title is the hero, items are secondary |

### Inspecting available templates

```bash
python scripts/apply_template.py --list
```

### How regenerate works internally

```
extract_content.py    → content.json   (title, subtitle, badge, items[], footer)
apply_template.py     → fresh.pptx     (renders content into chosen template)
state_summary.py      → state JSON     (scored vs baseline)
```

### Path B decision flow in the agent loop

```
INIT
  python scripts/state_summary.py --in input.pptx --work-dir state-0/
  → read state-0/state-summary.json + render PNG
  → record baseline_score and dominant issue categories

DECIDE
  if baseline_score < 25 and (density_score < 30 or hierarchy_score < 50):
      Tell user: "polish 上限低，建议重新套模板"
      Show available templates and item count
      → on user confirm:
          python scripts/regenerate.py --in input.pptx --work-dir regen/ --auto
          read regen/state/render/slide-001.png
          report 3-part: NUMERIC + VISUAL + DECISION
  else:
      proceed with Path A polish loop

REPORT
  baseline 0.0 (poor) → regenerated 82.79 (good)
  template: horizontal-timeline
  output: <work-dir>/iter-1.polished.pptx
```

### What you (the agent) MUST do

- **Report 3 parts every iteration**: numeric block (quoted from state-summary), visual observation (one sentence based on the render PNG), decision (ADOPT/REJECT with justification).
- **Read the render PNG every iteration** before deciding. Numeric scores alone are not sufficient — they don't catch "icon now overlaps title" type issues.
- **One mutate per iteration** so causality is clear and rollback is precise.
- **Use `--in input.pptx --out iter-N.pptx`** — never overwrite the input.

### What you (the agent) MUST NOT do

- Don't go past 12 iterations without showing progress to the user.
- Don't make many simultaneous changes.
- Don't trust the score blindly — `decision_recommendation: adopt` from state-summary is just a heuristic; visual judgment overrides.
- Don't change shapes you can't tie back to a clear issue or visual problem.

---

## Available primitives

### Read-only probes

| Script | Output | When to use |
|---|---|---|
| `inspect_ppt.py --input X --output Y.json` | geometry / fills / lines / fonts per shape | start of any session |
| `detect_roles.py --inspection ins.json --output roles.json` | per-shape role + row/col/card groups | needed for score_layout |
| `score_layout.py --inspection ins.json --roles roles.json --output find.json` | issues + 4 quantitative scores | rule-based detection |
| `render_slides.py --input X --output-dir imgs --manifest m.json` | PNG per slide via LibreOffice | get visual to look at |
| `diff_render.py --before A.png --after B.png --output d.json [--heatmap h.png]` | SSIM, pixel diff, heatmap | judge before/after |
| `self_critique.py --findings find.json --output crit.json` | weighted 0-100 score from metrics | rule-based score |
| **`state_summary.py --in X --work-dir Y`** | **all-of-the-above bundled** | **use this every iteration** |

### Granular mutate ops (45 total)

```bash
python scripts/mutate.py list-ops --json
```

Categories: geometry (11) · text (11) · style (10) · slide (5) · z-order (4) · connector (2) · repair (1).

Every op takes `--in <pptx> --out <pptx>` and emits one JSON record on stdout.

### Common patterns

**Fix peer-card outliers (oversized header strips, misplaced inner shapes, asymmetric card boxes)**
```bash
python scripts/mutate.py repair-peer-cards --in deck.pptx --out v1.pptx
```

**Send all overlapping arrows behind cards**
```bash
python scripts/mutate.py all-connectors-to-back --in deck.pptx --out v1.pptx
```

**Apply theme typography to one shape**
```bash
python scripts/mutate.py apply-typography --in deck.pptx --out v1.pptx --shape-id 2 --role title
```

**Apply theme card style (fill + border + corner) to a card**
```bash
python scripts/mutate.py apply-card-style --in deck.pptx --out v1.pptx --shape-id 5
```

**Align peers and equalize gaps**
```bash
python scripts/mutate.py align --in deck.pptx --out v1.pptx --shape-ids 23,24,25 --edge top
python scripts/mutate.py distribute --in v1.pptx --out v2.pptx --shape-ids 23,24,25 --axis horizontal
```

**Set every text to one font family**
```bash
python scripts/mutate.py set-font-family --in deck.pptx --out v1.pptx --scope all --family "Microsoft YaHei"
```

See [docs/mutate-ops.md](docs/mutate-ops.md) for all 45 ops.

## Themes

`themes/`: `clean-tech.json` (default tech blue), `academic-soft.json`, `business-warm.json`, `editorial-dark.json`.

Pass `--theme themes/<name>.json` to `apply-typography`, `apply-card-style`, `apply-badge-style`, `style-connector`. See [docs/themes.md](docs/themes.md).

## Guardrails

- Never overwrite the input `.pptx`
- Skip charts, tables, pictures, groups, anomalous-geometry shapes (logged as `skipped`)
- Tolerate missing `soffice` — only the visual rendering degrades; all polish + score logic stays functional
- Every mutation is in→out copy; chain naturally without state
- Structured JSONL logging via `PPT_POLISH_LOG=/path.jsonl` env var

## Internationalization

Default font `Microsoft YaHei` works on Windows. For Linux/macOS, edit `themes/<name>.json` `typography.font_family` to a locally available CJK font (`Noto Sans CJK SC`, `PingFang SC`, `Source Han Sans SC`).

## Example session (showing the mandatory 3-part report)

User: "用 ppt-audit-polish 美化 D:\deck.pptx，每一步给我看"

```
[INIT]
  python scripts/state_summary.py --in D:\deck.pptx --work-dir D:\out\state-0\
  → Read D:\out\state-0\render\slide-001.png with the Read tool
  → Baseline score 15.19 (poor). Top issue: peer-card-misplaced-child (shape 38, 39 drifted into card 2).

[ITER 1] candidate: repair-peer-cards
  python scripts/mutate.py repair-peer-cards --in D:\deck.pptx --out D:\out\iter-1.pptx
  python scripts/state_summary.py --in D:\out\iter-1.pptx --work-dir D:\out\state-1\ --diff-from D:\deck.pptx
  → Read state-1/render/slide-001.png

  ▸ NUMERIC (from state-1's agent_report.numeric_block):
      score: 15.19 → 23.47 (Δ +8.28)
        alignment: 83.95 → 91.20 (Δ +7.25)
        density:   0.0   → 0.0   (Δ 0.0)
        hierarchy: 30.0  → 30.0  (Δ 0.0)
        palette:   40.0  → 40.0  (Δ 0.0)
        issues:    Δ -3
        verdict:   poor → poor

  ▸ VISUAL: 卡片 3 (Few-shot) 的图标和"少量示例"文字现在正确落在卡片内部，
            卡片 4 (Many-shot) 头标的青色色块缩到了与同行一致的小徽章尺寸。

  ▸ DECISION: ADOPT — 数值 +8.28 达到阈值，视觉上 4 个明显 bug 全修了。

[ITER 2] candidate: ...
  ...

[STOP] score plateaued at 41.5 after iter 4.
  cp D:\out\iter-3\deck.pptx D:\deck.polished.pptx
  Final report: started at 15.19 (poor), ended at 41.5 (needs-work). Fixed peer-card outliers,
  unified card borders, equalized column 6. Remaining: density still low (slide is sparse by design).
```
