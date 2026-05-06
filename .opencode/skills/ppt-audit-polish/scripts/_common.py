"""Shared utilities for the ppt-audit-polish toolchain.

All low-level building blocks (color parsing, EMU/Pt math, theme loading,
JSONL logging) live here so every script is built on the same foundation.
Keeping these centralized is what allows the L2 mutate operations and the
L4 orchestrator to chain reliably without spec drift.
"""
from __future__ import annotations

import json
import os
import sys
import time
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt


SKILL_ROOT = Path(__file__).resolve().parent.parent
THEMES_DIR = SKILL_ROOT / "themes"
VARIANTS_DIR = SKILL_ROOT / "variants"


# ---------- Color ----------

def hex_to_rgb(hex_str: str) -> RGBColor:
    """Parse '#RRGGBB' or 'RRGGBB' into an RGBColor."""
    s = hex_str.lstrip("#")
    if len(s) != 6:
        raise ValueError(f"hex color must be 6 chars, got {hex_str!r}")
    return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def rgb_to_hex(rgb: RGBColor) -> str:
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def relative_luminance(hex_str: str) -> float:
    """WCAG relative luminance for an sRGB color."""
    s = hex_str.lstrip("#")
    r, g, b = (int(s[i : i + 2], 16) / 255.0 for i in (0, 2, 4))

    def channel(c: float) -> float:
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)


def contrast_ratio(fg_hex: str, bg_hex: str) -> float:
    """WCAG contrast ratio in [1, 21]."""
    l1 = relative_luminance(fg_hex)
    l2 = relative_luminance(bg_hex)
    lighter, darker = max(l1, l2), min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def best_text_on(bg_hex: str, light: str = "FFFFFF", dark: str = "161616") -> str:
    """Pick the better-contrast text color for a given background."""
    return light if contrast_ratio(light, bg_hex) >= contrast_ratio(dark, bg_hex) else dark


# ---------- Geometry ----------

def emu_to_inches(emu: int) -> float:
    return emu / 914400.0


def inches_to_emu(inches: float) -> int:
    return int(inches * 914400)


def pt_from_emu(emu: int) -> float:
    return emu / 12700.0


def normalize_geometry(left: int, top: int, width: int, height: int) -> tuple[int, int, int, int, bool]:
    """Convert a possibly-flipped bbox into a positive-dim bbox, plus flip flag."""
    flipped = width < 0 or height < 0
    if width < 0:
        left += width
        width = -width
    if height < 0:
        top += height
        height = -height
    return left, top, width, height, flipped


# ---------- Theme / config loading ----------

@dataclass(frozen=True)
class Theme:
    name: str
    palette: dict[str, str]
    typography: dict[str, Any]
    spacing: dict[str, int]
    decoration: dict[str, Any]
    raw: dict[str, Any]

    def color(self, role: str) -> str:
        return self.palette.get(role, self.palette.get("text", "393939"))

    def rgb(self, role: str) -> RGBColor:
        return hex_to_rgb(self.color(role))

    def font_size_pt(self, role: str) -> float | None:
        return self.typography.get("size_pt", {}).get(role)

    def font_bold(self, role: str) -> bool | None:
        return self.typography.get("bold", {}).get(role)

    def font_color_role(self, role: str) -> str | None:
        return self.typography.get("color_role", {}).get(role)

    def font_family(self) -> str | None:
        return self.typography.get("font_family")


def load_theme(path: str | Path | None = None) -> Theme:
    """Load and lightly-validate a theme JSON file."""
    if path is None:
        path = THEMES_DIR / "clean-tech.json"
    path = Path(path)
    raw = json.loads(path.read_text(encoding="utf-8"))
    _validate_theme(raw, source=str(path))
    return Theme(
        name=raw.get("name", path.stem),
        palette=raw["palette"],
        typography=raw["typography"],
        spacing=raw["spacing"],
        decoration=raw["decoration"],
        raw=raw,
    )


def _validate_theme(raw: dict, source: str) -> None:
    required_top = {"palette", "typography", "spacing", "decoration"}
    missing = required_top - raw.keys()
    if missing:
        raise ValueError(f"{source}: theme missing required keys: {sorted(missing)}")
    palette = raw["palette"]
    if "text" not in palette or "background" not in palette:
        raise ValueError(f"{source}: palette must define `text` and `background`")
    typo = raw["typography"]
    for key in ("size_pt", "bold", "color_role"):
        if key not in typo:
            raise ValueError(f"{source}: typography missing `{key}`")
    spacing = raw["spacing"]
    if "slide_padding_emu" not in spacing:
        raise ValueError(f"{source}: spacing missing `slide_padding_emu`")


# ---------- Structured logging ----------

class JsonlLogger:
    """Append-only JSON Lines logger with stable session id.

    Usage:
        log = JsonlLogger.from_env(component="mutate")
        log.event("apply-typography", shape_id=5, role="title")
    """

    def __init__(self, path: Path | None, component: str, session_id: str | None = None):
        self.path = path
        self.component = component
        self.session_id = session_id or uuid.uuid4().hex[:12]
        if path is not None:
            path.parent.mkdir(parents=True, exist_ok=True)

    @classmethod
    def from_env(cls, component: str) -> "JsonlLogger":
        env_path = os.environ.get("PPT_POLISH_LOG")
        path = Path(env_path) if env_path else None
        return cls(path=path, component=component, session_id=os.environ.get("PPT_POLISH_SESSION"))

    def event(self, event: str, **fields: Any) -> None:
        record = {
            "ts": time.time(),
            "session": self.session_id,
            "component": self.component,
            "event": event,
            **fields,
        }
        line = json.dumps(record, ensure_ascii=False)
        if self.path is None:
            sys.stderr.write(line + "\n")
        else:
            with self.path.open("a", encoding="utf-8") as fh:
                fh.write(line + "\n")


# ---------- File helpers ----------

def ensure_parent(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def write_json(path: Path, payload: Any) -> Path:
    ensure_parent(path)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


def read_json(path: Path) -> Any:
    return json.loads(Path(path).read_text(encoding="utf-8"))


# ---------- Shape ID helpers ----------

def parse_shape_ids(spec: str) -> list[int]:
    """Parse '5,6,7' or '5 6 7' or '5' into list[int]."""
    parts: Iterable[str]
    if "," in spec:
        parts = spec.split(",")
    else:
        parts = spec.split()
    return [int(p.strip()) for p in parts if p.strip()]
