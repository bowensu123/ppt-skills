"""Slide-master / slide-layout inheritance probe.

A slide silently inherits shapes from its layout (and the layout from its
master). The current ``inspect_ppt`` only walks ``slide.shapes`` so any
inherited title/footer/page-number is invisible to the detectors. That
causes false-positive "missing-title" reports when the title placeholder
is actually rendered from the layout.

This module exposes inheritance metadata WITHOUT injecting inherited
shapes into the primary objects list (which would break peer-row
clustering and density math). Detectors that care can read this metadata
to suppress false positives.
"""
from __future__ import annotations


def collect_inheritance_info(slide_obj) -> dict:
    """Walk slide -> slide_layout -> slide_master and summarize.

    Returns a dict the inspect step can attach to each slide:
      {
        "layout_name": str,
        "layout_shape_count": int,
        "layout_placeholders": [{"idx": int, "type": str, "name": str}, ...],
        "master_shape_count": int,
        "master_placeholders": [...],
      }
    """
    info: dict = {
        "layout_name": "",
        "layout_shape_count": 0,
        "layout_placeholders": [],
        "master_shape_count": 0,
        "master_placeholders": [],
    }
    try:
        layout = slide_obj.slide_layout
    except (AttributeError, KeyError):
        return info
    info["layout_name"] = getattr(layout, "name", "") or ""
    layout_shapes = list(getattr(layout, "shapes", []) or [])
    info["layout_shape_count"] = len(layout_shapes)
    info["layout_placeholders"] = _placeholders_summary(layout_shapes)

    try:
        master = layout.slide_master
    except (AttributeError, KeyError):
        return info
    master_shapes = list(getattr(master, "shapes", []) or [])
    info["master_shape_count"] = len(master_shapes)
    info["master_placeholders"] = _placeholders_summary(master_shapes)
    return info


def _placeholders_summary(shapes) -> list[dict]:
    out = []
    for shape in shapes:
        try:
            ph = shape.placeholder_format
        except AttributeError:
            continue
        if ph is None:
            continue
        out.append({
            "idx": ph.idx if ph.idx is not None else -1,
            "type": str(ph.type) if ph.type is not None else "",
            "name": getattr(shape, "name", ""),
        })
    return out


def has_inherited_title(inheritance: dict) -> bool:
    """Heuristic: layout/master has a TITLE placeholder."""
    candidates = inheritance.get("layout_placeholders", []) + inheritance.get("master_placeholders", [])
    for ph in candidates:
        ptype = (ph.get("type") or "").upper()
        if "TITLE" in ptype:
            return True
    return False
