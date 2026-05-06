# ai-ppt

Standalone OpenCode skill repository for `ppt-audit-polish`.

## What it does

- Audits `.pptx` files for overflow, misalignment, spacing drift, hierarchy inconsistency, and crowded layouts
- Repairs peer-row alignment, peer spacing, and comparable text hierarchy in `direct-fix` mode
- Renders original and polished slide images for visual validation when `LibreOffice` (`soffice`) is available
- Produces a `*.audit-report.md` report for every successful run
- Produces a `*.polished.pptx` copy in `direct-fix` mode without overwriting the original deck

## Repository layout

- `.opencode/skills/ppt-audit-polish/`
- `scripts/install_ppt_audit_polish.sh`

## Install as a global OpenCode skill

```bash
./scripts/install_ppt_audit_polish.sh
```

This installs the skill to:

```text
~/.config/opencode/skills/ppt-audit-polish
```

## Use in OpenCode

Examples:

- `Check and beautify this PPT`
- `Audit this presentation for overflow and alignment issues`
- `Directly fix this deck and give me a polished version`

## Run directly

Install dependencies first:

```bash
python -m pip install -e .opencode/skills/ppt-audit-polish[dev]
```

For visual validation, install LibreOffice and make `soffice` available on `PATH`, or set `SOFFICE_PATH` to the absolute path of `soffice.exe`.

Audit only:

```bash
python ~/.config/opencode/skills/ppt-audit-polish/scripts/run_ppt_audit_polish.py \
  --input /absolute/path/to/deck.pptx \
  --mode audit \
  --work-dir /absolute/path/to/output-dir
```

Direct fix:

```bash
python ~/.config/opencode/skills/ppt-audit-polish/scripts/run_ppt_audit_polish.py \
  --input /absolute/path/to/deck.pptx \
  --mode direct-fix \
  --work-dir /absolute/path/to/output-dir
```
