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

## Vision-first analysis (recommended over rule suggestions)

Rule-based detection (`score_layout`) flags issues from geometry alone — it
can't see whether "the icon visually belongs to card 4 even though its
bbox-center is in card 3". For damaged decks, **trust your eyes over the
hint list**.

`state_summary.py` emits TWO renders every iteration:

1. `render/slide-001.png` — clean visual (judge whether it looks right)
2. `annotated/slide-001.annotated.png` — the same slide with each shape's
   `#shape_id` and kind-color overlay; use this to translate "the icon I
   see misplaced" → the exact `shape_id` you pass to mutate.py

### Vision-driven decision template

```
1. Read render/slide-001.png             — does it look right?
2. If something looks wrong:
   3. Read annotated/slide-001.annotated.png  — find the shape_id of the bad element
   4. Pick a model-friendly mutate op (placement category):
        place-near                 → "put icon NEAR this card's top-left"
        mirror-peer-position       → "put icon at the SAME offset as a sibling card has"
        move-to-card               → "drop shape into card N at fractional rel pos"
   5. Run it. Re-render. Look again.
   6. Repeat or revert based on what you see, NOT the rule scores.
```

### When rule suggestions are reliable

- `boundary-overflow` (geometry is unambiguous)
- `shape-overlap` between non-card peers
- Trivially low-contrast text where peer wins WCAG by a wide margin

### When to override rule suggestions with vision

- Anything involving "this is in the wrong card" — rules use bbox-center
  containment which fails when source has oversized cards
- `repair-peer-cards` action with >5 orphan_relocations — fall back to
  `--scope safe` and place shapes individually using mirror-peer-position
- "Should this text be in card 3 or 4?" — bbox heuristics flip; your
  visual reasoning is authoritative

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

When polish ceiling is too low, regenerate produces a fresh slide using a pre-built template. **Template selection is agent-driven** — there is no `--auto` heuristic. Item count alone cannot tell whether 4 items are 4 sequential steps or 4 parallel features, so YOU (the agent) read the content + look at the original deck and decide.

### Available templates

| Template | Best for | Pick when… |
|---|---|---|
| `horizontal-timeline` | Process steps, stages, comparison classes | Items are sequential / ordered (Step 1→2→3, Zero-shot → Few-shot → Many-shot, before→after). The original deck likely had connector arrows or numeric prefixes. |
| `grid-2x3` | Independent feature/benefit lists | Items are parallel and exchangeable (Performance / Security / Cost / Reliability). Order doesn't matter. |
| `feature-list` | One hero topic, items are secondary | Title is the spotlight; items are short bullets supporting it. Works for any item count including 0. |

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

### Path B agent-driven decision flow

```
INIT (same as Path A)
  python scripts/state_summary.py --in input.pptx --work-dir state-0/
  → read state-0/state-summary.json + state-0/render/slide-001.png
  → record baseline_score and dominant issue categories

DECIDE (whether to switch to Path B)
  if baseline_score < 25 and (density_score < 30 or hierarchy_score < 50):
      Tell user: "polish 上限低，建议重新套模板"
      → on user confirm: enter Path B template-selection below
  else:
      stay in Path A polish loop

PATH B — STEP 1: Extract content
  python scripts/extract_content.py --in input.pptx --work-dir regen/
  → produces regen/content.json

PATH B — STEP 2: Look + judge (THE AGENT-DECISION STEP)
  Read regen/content.json with the Read tool — see title, items[].name,
  items[].description, badge, footer.
  Re-look at state-0/render/slide-001.png — does the original layout
  use connectors / arrows / numbered badges?

  Decide template:
    - Sequential signals (numeric prefixes "Step 1" / "①" / "第一",
      time words "First / Then / Next / Finally", connector arrows in
      original render, items reference each other) → horizontal-timeline
    - Parallel signals (parallel-noun item names, no order implied,
      grid arrangement in original) → grid-2x3
    - Title-dominant + short item bullets → feature-list

  Report your reasoning in ONE sentence to the user before running step 3.

PATH B — STEP 3: Regenerate with chosen template
  python scripts/regenerate.py --in input.pptx --work-dir regen/ \
      --template <chosen-name>
  → reads regen/content.json (already extracted), runs apply_template +
    state_summary, copies final to regen/<stem>.polished.pptx

PATH B — STEP 4: Report 3 parts (same template as Path A)
  ▸ NUMERIC: baseline → regenerated (Δ)
  ▸ VISUAL:  read regen/state/render/slide-001.png and describe it
  ▸ DECISION: ADOPT (almost always — regenerate is opt-in already) or
              REJECT and try a different template

REPORT
  baseline 0.0 (poor) → regenerated 82.79 (good)
  template: horizontal-timeline (chosen because items had "Zero-shot →
            Few-shot → Many-shot" sequence pattern)
  output: <work-dir>/<stem>.polished.pptx
```

### Backward-compatible CLI

```bash
# Explicit template (only supported form)
python scripts/regenerate.py --in deck.pptx --work-dir out/ --template horizontal-timeline
python scripts/regenerate.py --in deck.pptx --work-dir out/ --template grid-2x3
python scripts/regenerate.py --in deck.pptx --work-dir out/ --template feature-list

# With theme override
python scripts/regenerate.py --in deck.pptx --work-dir out/ \
    --template grid-2x3 --theme themes/business-warm.json
```

`--auto` was removed deliberately. Calling it now produces an explicit
error message guiding the agent to the extract → judge → regenerate flow.

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
| `render_slides.py --input X --output-dir imgs --manifest m.json [--also-svg]` | PNG per slide via LibreOffice; `--also-svg` writes svg/ for post-render signal extraction | get visual to look at |
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
# Moderately-damaged decks (typical case): applies all four fix kinds.
python scripts/mutate.py repair-peer-cards --in deck.pptx --out v1.pptx

# HEAVILY-DAMAGED decks (>5 orphan_relocations in default-run output): use safe.
# Only the well-constrained card-box-fix and header-strip-fix run; the
# aggressive orphan/displaced relocators (which cascade incorrectly when
# many shapes already sit outside their cards) are skipped.
python scripts/mutate.py repair-peer-cards --in deck.pptx --out v1.pptx --scope safe
```

When the agent loop's first run reports many orphan_relocations and the
visual gets *worse* on inspection, REVERT and re-run with `--scope safe`.

**Send all overlapping arrows behind cards**
```bash
python scripts/mutate.py all-connectors-to-back --in deck.pptx --out v1.pptx
```

**Business-grade polish (one-click, doesn't change content)**
```bash
# Auto-pick theme based on content keywords; standard polish level.
python scripts/mutate.py polish-business --in deck.pptx --out v1.pptx

# Explicit level (1=subtle, 2=standard default, 3=rich) and theme.
python scripts/mutate.py polish-business --in deck.pptx --out v1.pptx \
    --level 3 --theme themes/business-warm.json
```

What it does (idempotent — safe to re-run):
- **Smart typography**: applies the theme's type scale (size + weight +
  color) to every detected role (title / subtitle / body / caption).
- **Geometric consistency**: unifies corner radius across rounded
  containers; unifies subtle shadow on card-sized shapes.
- **Hierarchy decoration** (level ≥ 2): adds a thin primary-color
  accent bar above each title; adds a 0.5pt divider line above the
  bottom band.
- **Surface tint** (level 3 only): adds a very-pale background
  rectangle behind each subtitle.

What it never does:
- Modify any text content (run.text stays untouched).
- Delete or hide existing shapes.
- Add decoration that overlaps existing content (decorations go to
  z-order back; if there's no space above the title, the accent bar is
  skipped).

Theme is auto-picked from content keywords:
- "AI / 框架 / model / agent" → `clean-tech`
- "营收 / ROI / 客户 / Q1" → `business-warm`
- "研究 / 实验 / 论文 / hypothesis" → `academic-soft`
- "故事 / 叙事 / 品牌 / editorial" → `editorial-dark`

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
