---
name: ppt-audit-polish
description: Audit PowerPoint decks for overflow, misalignment, spacing, hierarchy, and clutter issues, then optionally create a polished copy with a report when the user asks to fix the deck directly.
compatibility: opencode
metadata:
  audience: presentation-authors
  workflow: ppt-review
---

## What I do

- Inspect `.pptx` decks using deterministic geometry extraction
- Repair peer-row alignment, comparable text hierarchy, and uneven peer spacing in direct-fix mode
- Render original and polished slide images for visual validation when `soffice` is available
- Produce `*.audit-report.md` in all successful runs
- Produce `*.polished.pptx` only in direct-fix mode

## When to use me

Use this when the user says things like:

- "Check and beautify this PPT"
- "Audit this presentation for overflow and alignment issues"
- "Directly fix this deck and give me a polished version"
- "Review this PPT for layout problems"

## Mode selection

- Use `audit` unless the user explicitly asks to fix or polish the deck right now
- Use `direct-fix` only when the user clearly asks for immediate repair

## Guardrails

- Never overwrite the original input file
- Only normalize text boxes that are likely peers: same row and similar dimensions
- Skip or report high-risk content such as charts and complex layouts instead of heavily rewriting it
- Report missing rendering dependencies instead of failing the whole run
- For visual validation, use LibreOffice on `PATH` or set `SOFFICE_PATH` to the absolute `soffice.exe` path

## Execution

Run this command from the skill root:

`python scripts/run_ppt_audit_polish.py --input /absolute/path/to/deck.pptx --mode audit --work-dir /absolute/path/to/output-dir`

For immediate repair:

`python scripts/run_ppt_audit_polish.py --input /absolute/path/to/deck.pptx --mode direct-fix --work-dir /absolute/path/to/output-dir`

Expected outputs:

- `/absolute/path/to/output-dir/<deck>.audit-report.md`
- `/absolute/path/to/output-dir/<deck>.polished.pptx` in direct-fix mode
- `/absolute/path/to/output-dir/<deck>.audit-artifacts/`
- `/absolute/path/to/output-dir/<deck>.audit-artifacts/images/` and `polished-images/` when rendering is available
