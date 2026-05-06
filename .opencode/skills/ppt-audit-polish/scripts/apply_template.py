"""Apply a template to a content JSON, producing a fresh .pptx.

The output deck has ONE slide regenerated from scratch using the template.
The original deck is never modified.

CLI:
  python apply_template.py --content content.json --template horizontal-timeline \
      --out out.pptx [--theme themes/clean-tech.json] [--slide-width 12192000]
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

from _common import load_theme


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
TEMPLATES_DIR = SKILL_ROOT / "templates_py"


def _load_template_module(template_name: str):
    """Load a template module by its NAME constant (e.g., 'horizontal-timeline')."""
    candidates = list(TEMPLATES_DIR.glob("*.py"))
    for path in candidates:
        if path.name.startswith("_"):
            continue
        spec = importlib.util.spec_from_file_location(path.stem, path)
        mod = importlib.util.module_from_spec(spec)
        # Temporarily expose template directory on sys.path so ``_base`` import works.
        if str(TEMPLATES_DIR) not in sys.path:
            sys.path.insert(0, str(TEMPLATES_DIR))
        spec.loader.exec_module(mod)
        if getattr(mod, "NAME", path.stem) == template_name:
            return mod
    raise FileNotFoundError(f"template '{template_name}' not found in {TEMPLATES_DIR}")


def list_templates() -> list[dict]:
    out = []
    if str(TEMPLATES_DIR) not in sys.path:
        sys.path.insert(0, str(TEMPLATES_DIR))
    for path in TEMPLATES_DIR.glob("*.py"):
        if path.name.startswith("_"):
            continue
        spec = importlib.util.spec_from_file_location(path.stem, path)
        mod = importlib.util.module_from_spec(spec)
        try:
            spec.loader.exec_module(mod)
        except Exception as exc:
            out.append({"name": path.stem, "error": str(exc)[:200]})
            continue
        out.append({
            "name": getattr(mod, "NAME", path.stem),
            "file": path.name,
            "description": getattr(mod, "DESCRIPTION", ""),
            "min_items": getattr(mod, "MIN_ITEMS", 0),
            "max_items": getattr(mod, "MAX_ITEMS", 0),
            "preferred_theme": getattr(mod, "PREFERRED_THEME", None),
        })
    return out


def apply_template(
    content: dict,
    template_name: str,
    out_path: Path,
    theme_path: Path | None = None,
    slide_width_emu: int = 12192000,
    slide_height_emu: int = 6858000,
) -> dict:
    template = _load_template_module(template_name)
    theme = load_theme(theme_path)

    prs = Presentation()
    prs.slide_width = Emu(slide_width_emu)
    prs.slide_height = Emu(slide_height_emu)

    result = template.render(prs, content, theme)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return {
        "template": template_name,
        "theme": theme.name,
        "out": str(out_path),
        "items_rendered": result.get("items_rendered", 0),
        "warnings": result.get("warnings", []),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--content", type=Path, help="content JSON from extract_content.py")
    parser.add_argument("--template", help="template name (e.g., horizontal-timeline)")
    parser.add_argument("--out", type=Path, help="output .pptx")
    parser.add_argument("--theme", type=Path, help="theme JSON; defaults to template's preferred or clean-tech")
    parser.add_argument("--slide-width", type=int, default=12192000)
    parser.add_argument("--slide-height", type=int, default=6858000)
    parser.add_argument("--list", action="store_true", help="list available templates")
    args = parser.parse_args()

    if args.list:
        print(json.dumps(list_templates(), ensure_ascii=False, indent=2))
        return 0

    if not (args.content and args.template and args.out):
        parser.error("--content, --template, and --out are required (or pass --list)")

    content = json.loads(args.content.read_text(encoding="utf-8"))
    payload = apply_template(
        content=content,
        template_name=args.template,
        out_path=args.out,
        theme_path=args.theme,
        slide_width_emu=args.slide_width,
        slide_height_emu=args.slide_height,
    )
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
