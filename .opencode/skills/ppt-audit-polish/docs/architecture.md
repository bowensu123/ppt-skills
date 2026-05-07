# Architecture

Four layers, each composed of single-purpose scripts. Higher layers only call lower-layer scripts via subprocess + JSON, never via Python imports across boundaries (so you can swap any layer's implementation without breaking callers).

```
                                     +---------------------------+
                                     | Mode 1: variants (batch)  |
                                     | Mode 2: orchestrator      |
                                     | Mode 3: model + mutate    |
                                     +---------+-----------------+
                                               |
                  +----------------------------+----------------------------+
                  v                                                         v
+-----------------------------+                          +-----------------------------+
| L4 self-improvement         |                          | L3 recipes                  |
|  - self_critique.py         |                          |  - polish_engine.py         |
|  - polish_orchestrator.py   |                          |  - run_variants.py          |
+--------------+--------------+                          +--------------+--------------+
               |                                                        |
               +-------------------------+------------------------------+
                                         v
                       +-----------------------------+
                       | L2 granular mutate          |
                       |  - mutate.py (44 ops)       |
                       +--------------+--------------+
                                      |
                                      v
                       +-----------------------------+
                       | L1 probes (read-only)       |
                       |  - inspect_ppt.py           |
                       |  - detect_roles.py          |
                       |  - score_layout.py          |
                       |  - render_slides.py         |
                       |  - _svg_geom.py             |
                       +--------------+--------------+
                                      |
                                      v
                       +-----------------------------+
                       | L0 foundations              |
                       |  - _common.py               |
                       |  - _shape_ops.py            |
                       |  - themes/, variants/       |
                       +-----------------------------+
```

## Contracts

Every L2 op:
- Reads `--in <file>` and writes `--out <file>` (deck mutation contract)
- Emits one JSON record on stdout (machine-readable result)
- Emits zero or more JSONL records on stderr or `$PPT_POLISH_LOG` (telemetry)
- Exits 0 on success, 2 on usage error, 3 on runtime error
- Never modifies the input file

Every L1 probe:
- Reads `--input` or `--inspection` JSON
- Writes `--output` JSON (or images, for `render_slides.py`)
- Idempotent — running twice with same inputs gives same outputs

Every L3 recipe execution:
- Pure function of (input deck, theme, options JSON)
- All inputs hash-stable

Every L4 candidate:
- Generates a candidate deck via L2 mutate or L3 polish
- Scores via `self_critique.py`
- Compared by score only — no side state, no LLM judgment

## Telemetry

Set `PPT_POLISH_LOG=/path.jsonl` to collect JSON Lines events from every script. Each event has:
- `ts` (unix seconds)
- `session` (12-char hex id, persists across a single chain via `PPT_POLISH_SESSION`)
- `component` (script name)
- `event` (verb, e.g. `apply-typography`, `baseline-scored`)
- ...arbitrary structured fields

Useful for:
- Reproducing a model session step-by-step
- Building dashboards on which ops the model picks most often
- Auditing automated runs in CI

## Adding a new mutate op

1. Add the primitive to `scripts/_shape_ops.py` (returns a small dict or None).
2. Add a `cmd_<name>(args)` function in `scripts/mutate.py`, decorated with `@op(category, summary, example)`.
3. Add an argparse subparser entry in `build_parser()`.
4. Add a unit test in `tests/test_mutate_cli.py`.
5. Re-run the docs generator (see [README inside docs/](.) — `mutate-ops.md` is regenerated from `mutate.py list-ops --json`).

That's it. The op auto-appears in `list-ops`, becomes available to the orchestrator's candidate generator (if you add a mapping), and slots into recipes by adding a new option toggle.
