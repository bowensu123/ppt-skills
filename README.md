# ppt-audit-polish

OpenCode skill for auditing and polishing PowerPoint decks. The skill exposes
a small set of Python scripts that the OpenCode agent calls to inspect, render,
and mutate `.pptx` files; the agent itself drives the iteration loop using its
own multimodal model.

**No separate model setup, no API keys, no env vars.** The agent uses whatever
model OpenCode is already running with.

## What's in the box

- **`scripts/state_summary.py`** — one-shot state view (score + top issues + render PNG path + ready-to-run mutate argv + diff vs previous iteration)
- **`scripts/mutate.py`** — 45 atomic mutation operations (geometry, style, text, z-order, slide-level, repair). Each takes `--in <pptx> --out <pptx>` and emits one JSON record on stdout.
- **`scripts/inspect_ppt.py` / `detect_roles.py` / `score_layout.py` / `render_slides.py` / `diff_render.py` / `self_critique.py`** — read-only probes used by `state_summary.py` and callable individually.
- **`themes/`** — 4 themes (`clean-tech`, `academic-soft`, `business-warm`, `editorial-dark`).

The OpenCode agent reads `SKILL.md`, follows the agent-driven playbook, and uses Read+Bash tools to chain the primitives.

## Install

```bash
./scripts/install_ppt_audit_polish.sh
```

This copies the skill to `~/.config/opencode/skills/ppt-audit-polish/`. OpenCode discovers it on next start.

## Dependencies

Inside the skill (Python 3.11+):

```bash
python -m pip install -e .opencode/skills/ppt-audit-polish[dev]
```

That installs:
- `python-pptx >= 1.0.2`
- `Pillow >= 11.0`
- `PyMuPDF >= 1.25`
- `numpy >= 1.26`

Plus optional but recommended:
- **LibreOffice** for slide rendering (the agent's vision input). Set `SOFFICE_PATH` if `soffice` isn't on `PATH`.

## Use

Open OpenCode, then say something like:

```
用 ppt-audit-polish 美化 D:\deck.pptx，每一步给我看
```

The agent will:

1. Run `state_summary.py` to get score + issues + a render of the slide.
2. Read the render PNG with its own vision.
3. Pick a mutate op (each top issue includes ready-to-run argv).
4. Run `mutate.py <op>` to produce the next iteration's `.pptx`.
5. Run `state_summary.py` again with `--diff-from` to see exactly what changed.
6. Look at the new render. Decide if it improved.
7. Loop until the score plateaus.
8. Copy the best iteration to `<input>.polished.pptx` and summarize.

The agent's own multimodal model does all visual reasoning. No outside model
setup.

## Run primitives directly (without OpenCode)

If you want to drive the loop yourself or inside CI:

```bash
# State snapshot
python scripts/state_summary.py --in deck.pptx --work-dir state/

# Apply a single mutation
python scripts/mutate.py repair-peer-cards --in deck.pptx --out v1.pptx

# Compare two renders pixel-wise
python scripts/diff_render.py --before a.png --after b.png --output d.json --heatmap h.png

# List every mutate op
python scripts/mutate.py list-ops --json
```

See `.opencode/skills/ppt-audit-polish/SKILL.md` for the full agent playbook
and `.opencode/skills/ppt-audit-polish/docs/` for op catalog, theme authoring,
and architecture notes.
