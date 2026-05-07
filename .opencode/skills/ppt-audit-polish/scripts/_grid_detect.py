"""2D grid layout detection.

Complements `_card_repair.py` which only handles 1xN horizontal rows of
cards. Many real decks lay content out as MxN grids (e.g., 2-column x
3-row dashboard, 3x3 feature matrix). With < 3 peers per row, the 1D
peer-row algorithm refuses to fire — leaving misaligned grid cells
undetected.

This module:
  1. Filters candidate panel-like containers by similar W/H.
  2. Clusters them by `top` (rows) and `left` (columns) using fixed eps.
  3. Validates the grid: panel_count must be ≥ 80% of rows × cols.
  4. Computes per-row top anchors and per-column left anchors via median.
  5. Identifies outlier panels and proposes snap targets.
  6. Optionally recurses into each panel to find sub-grids (nested layouts).

Output is structurally compatible with `_card_repair.diagnose_repair` so
`apply_repair` can apply both 1D card fixes and 2D grid fixes via the
same code path.
"""
from __future__ import annotations

from collections import defaultdict
from statistics import median
from typing import Iterable


# Tolerances (EMU). 914400 EMU = 1 inch.
GRID_TOP_CLUSTER_EPS = 200000      # ~0.22" — same row if top differs less than this
GRID_LEFT_CLUSTER_EPS = 200000     # same col if left differs less than this
GRID_OUTLIER_OFFSET = 250000       # ~0.27" off anchor → outlier
GRID_PANEL_MIN_W = 800000          # candidate panel needs this minimum width
GRID_PANEL_MIN_H = 600000          # ...and this minimum height
GRID_SIZE_RATIO_TOL = 0.40         # peer panels can vary up to 40% in W/H
GRID_VALIDITY_RATIO = 0.60         # need at least 60% of (rows × cols) cells filled
HEADER_STRIP_OFFSET_TOL = 100000   # ~0.11" relative-offset deviation = strip outlier


# ----- 1-D fixed-eps clustering -----

def _cluster_1d(values_with_ids: list[tuple[int, int]], eps: int) -> list[list[int]]:
    """Sort by value, split where consecutive gap > eps. Returns clusters of ids."""
    if not values_with_ids:
        return []
    ordered = sorted(values_with_ids, key=lambda p: p[0])
    clusters: list[list[int]] = [[ordered[0][1]]]
    last = ordered[0][0]
    for v, sid in ordered[1:]:
        if v - last > eps:
            clusters.append([sid])
        else:
            clusters[-1].append(sid)
        last = v
    return clusters


# ----- candidate panel filtering -----

def _candidate_panels(objects: Iterable[dict]) -> list[dict]:
    """Containers / shapes that are large enough to be 'grid cells'."""
    return [
        o for o in objects
        if not o.get("anomalous")
        and o.get("kind") in ("container", "shape")
        and o.get("width", 0) >= GRID_PANEL_MIN_W
        and o.get("height", 0) >= GRID_PANEL_MIN_H
    ]


def _largest_size_cohort(panels: list[dict]) -> list[dict]:
    """Pick the cohort of largest INDIVIDUAL panel area among groups with
    ≥ 4 members. This biases outer-grid detection toward the biggest
    elements (so nested layouts are walked from outside in).

    Cohorts are formed by binning (W, H) to the nearest 100K EMU cell —
    panels of the same role tend to be near-identical in size.
    """
    if not panels:
        return []
    by_size: dict[tuple[int, int], list[dict]] = defaultdict(list)
    for p in panels:
        bucket = (
            round(p["width"] / 100000) * 100000,
            round(p["height"] / 100000) * 100000,
        )
        by_size[bucket].append(p)

    eligible = [(bucket, group) for bucket, group in by_size.items() if len(group) >= 4]
    if eligible:
        # Largest individual area wins → outer grid in nested layouts.
        bucket, cohort = max(eligible, key=lambda x: x[0][0] * x[0][1])
        return cohort

    # Fall back: most populous group of any size (handles 2x2 and small grids).
    return max(by_size.values(), key=len)


# ----- grid detection -----

def _reliable_anchors(values: list[int], eps: int, min_count: int = 2) -> list[int]:
    """Find anchor values that ≥ min_count panels share (within eps).

    These are RELIABLE anchors because multiple panels agree on them.
    Singleton clusters (one panel = one cluster) are NOT reliable —
    they're often the outlier itself.
    """
    if not values:
        return []
    clusters = _cluster_1d([(v, i) for i, v in enumerate(values)], eps=eps)
    anchors = []
    for cluster in clusters:
        if len(cluster) >= min_count:
            anchor_values = [values[i] for i in cluster]
            anchors.append(int(median(anchor_values)))
    return sorted(anchors)


def _extrapolate_anchors(anchors: list[int]) -> list[int]:
    """Given ≥ 2 reliable anchors, infer the full set of expected positions
    by extending in both directions using the median spacing.
    """
    if len(anchors) < 2:
        return list(anchors)
    spacings = [anchors[i + 1] - anchors[i] for i in range(len(anchors) - 1)]
    med_spacing = int(median(spacings))
    if med_spacing <= 0:
        return list(anchors)
    full = list(anchors)
    # Backward
    candidate = anchors[0] - med_spacing
    while candidate > -med_spacing:  # allow one slot past 0
        if candidate >= 0:
            full.insert(0, candidate)
        candidate -= med_spacing
    # Forward
    candidate = anchors[-1] + med_spacing
    # We don't know the slide bounds here; just project a couple extra slots
    for _ in range(3):
        full.append(candidate)
        candidate += med_spacing
    return sorted(set(full))


def detect_grid(panels: list[dict]) -> dict | None:
    """Detect an MxN grid using majority-vote reliable anchors.

    Algorithm (more robust than naive top/left clustering):
    1. Cluster tops; anchors are clusters with ≥ 2 panels (so an outlier in
       its own row doesn't pollute the anchor set).
    2. Same for lefts.
    3. Extrapolate missing row/col anchors using median spacing.
    4. For each panel, snap to the NEAREST (row_anchor, col_anchor) cell.
    5. If snap distance > tolerance, the panel is an outlier and its
       target is that nearest cell.
    """
    if len(panels) < 4:
        return None

    panel_by_id = {p["shape_id"]: p for p in panels}
    tops = [p["top"] for p in panels]
    lefts = [p["left"] for p in panels]

    reliable_rows = _reliable_anchors(tops, eps=GRID_TOP_CLUSTER_EPS)
    reliable_cols = _reliable_anchors(lefts, eps=GRID_LEFT_CLUSTER_EPS)

    # If we have <2 reliable in either dim, fall back to using all clusters
    # (necessary for perfectly aligned grids where every position is a
    # reliable anchor by definition).
    if len(reliable_rows) < 2:
        all_row_clusters = _cluster_1d([(t, i) for i, t in enumerate(tops)], eps=GRID_TOP_CLUSTER_EPS)
        reliable_rows = sorted(int(median([tops[i] for i in c])) for c in all_row_clusters)
    if len(reliable_cols) < 2:
        all_col_clusters = _cluster_1d([(l, i) for i, l in enumerate(lefts)], eps=GRID_LEFT_CLUSTER_EPS)
        reliable_cols = sorted(int(median([lefts[i] for i in c])) for c in all_col_clusters)

    # Extrapolate to fill gaps where outliers create missing rows/cols.
    full_row_anchors = _extrapolate_anchors(reliable_rows)
    full_col_anchors = _extrapolate_anchors(reliable_cols)

    # Trim extrapolated anchors that no panel comes near (avoids inventing
    # cells that don't exist).
    def _used_anchors(anchors: list[int], values: list[int]) -> list[int]:
        used = set()
        for v in values:
            nearest = min(range(len(anchors)), key=lambda i: abs(anchors[i] - v))
            if abs(anchors[nearest] - v) <= GRID_OUTLIER_OFFSET * 2:
                used.add(nearest)
        return sorted(anchors[i] for i in used)

    row_anchors = _used_anchors(full_row_anchors, tops)
    col_anchors = _used_anchors(full_col_anchors, lefts)

    rows = len(row_anchors)
    cols = len(col_anchors)
    if rows < 2 or cols < 2:
        return None

    # Map each panel to its (row, col) by nearest snap.
    sid_to_pos: dict[int, tuple[int, int]] = {}
    outliers: list[dict] = []

    for sid, panel in panel_by_id.items():
        r = min(range(rows), key=lambda i: abs(row_anchors[i] - panel["top"]))
        c = min(range(cols), key=lambda i: abs(col_anchors[i] - panel["left"]))
        sid_to_pos[sid] = (r, c)
        target_top = row_anchors[r]
        target_left = col_anchors[c]
        dx = panel["left"] - target_left
        dy = panel["top"] - target_top
        if abs(dx) > GRID_OUTLIER_OFFSET or abs(dy) > GRID_OUTLIER_OFFSET:
            outliers.append({
                "shape_id": sid,
                "name": panel["name"],
                "row": r, "col": c,
                "actual": [panel["left"], panel["top"]],
                "target": [target_left, target_top],
                "dx_dy": [dx, dy],
            })

    fill_ratio = len(panels) / (rows * cols)
    if fill_ratio < GRID_VALIDITY_RATIO:
        return None

    return {
        "rows": rows,
        "cols": cols,
        "fill_ratio": round(fill_ratio, 3),
        "row_anchors": row_anchors,
        "col_anchors": col_anchors,
        "panel_ids": [p["shape_id"] for p in panels],
        "panel_grid_pos": {sid: list(pos) for sid, pos in sid_to_pos.items()},
        "outliers": outliers,
    }


# ----- nested grid detection (one level down) -----

def _bbox_contains(outer: dict, inner: dict, slack: int = 91440) -> bool:
    return (
        inner["left"] >= outer["left"] - slack
        and inner["top"] >= outer["top"] - slack
        and inner["left"] + inner["width"] <= outer["left"] + outer["width"] + slack
        and inner["top"] + inner["height"] <= outer["top"] + outer["height"] + slack
    )


def detect_grids_nested(slide_objects: list[dict], max_depth: int = 2) -> list[dict]:
    """Detect grids at the slide level, then recurse into each panel.

    Returns a flat list of grids; outer grid first, sub-grids tagged with
    `parent_panel_id` so callers can correlate.
    """
    candidates = _candidate_panels(slide_objects)
    cohort = _largest_size_cohort(candidates)
    grid = detect_grid(cohort)
    if not grid:
        return []
    grid["depth"] = 0
    grid["parent_panel_id"] = None
    out = [grid]

    if max_depth <= 1:
        return out

    panel_lookup = {p["shape_id"]: p for p in cohort}
    for panel_sid in grid["panel_ids"]:
        panel = panel_lookup[panel_sid]
        children = [
            o for o in slide_objects
            if o["shape_id"] != panel_sid
            and not o.get("anomalous")
            and _bbox_contains(panel, o)
        ]
        # For a nested grid, scale down min sizes
        smaller_candidates = [
            o for o in children
            if o.get("kind") in ("container", "shape")
            and o.get("width", 0) >= 400000
            and o.get("height", 0) >= 250000
        ]
        sub_cohort = _largest_size_cohort(smaller_candidates)
        sub_grid = detect_grid(sub_cohort)
        if sub_grid:
            sub_grid["depth"] = 1
            sub_grid["parent_panel_id"] = panel_sid
            out.append(sub_grid)

    return out


# ----- header-strip relative-offset outlier detection -----

def _strip_candidates_for_panel(panel: dict, all_objects: list[dict]) -> list[dict]:
    """Return filled-color shapes that look like a header strip for `panel`.

    A "header strip" is a small filled rectangle that lives at or near the
    panel's top edge. Width is comparable to the panel; height is a small
    fraction of the panel's height. We allow some vertical slack ABOVE the
    panel because that's exactly the failure mode we're trying to detect
    (strip floating above its panel).
    """
    out = []
    pw = panel["width"]; ph = panel["height"]
    pl = panel["left"]; pt = panel["top"]
    # Vertical band: from 1.5x panel-height ABOVE panel top down to 35% INTO the panel.
    band_top_min = pt - int(ph * 1.5)
    band_top_max = pt + int(ph * 0.35)
    for o in all_objects:
        if o["shape_id"] == panel["shape_id"]:
            continue
        if o.get("anomalous"):
            continue
        if o.get("kind") not in ("container", "shape"):
            continue
        if not o.get("fill_hex"):
            continue
        # Same fill as the panel itself = it's the panel, skip
        if o.get("fill_hex") == panel.get("fill_hex"):
            continue
        # Geometry filters
        if not (band_top_min <= o["top"] <= band_top_max):
            continue
        if o["height"] >= ph * 0.5:           # too tall to be a strip
            continue
        if o["width"] < pw * 0.20:            # too narrow to be a strip
            continue
        if o["width"] > pw * 1.20:            # way wider than the panel
            continue
        # Horizontal containment: strip's x range must mostly overlap panel's
        ox_l = o["left"]; ox_r = o["left"] + o["width"]
        if ox_r < pl or ox_l > pl + pw:
            continue
        out.append(o)
    return out


def _detect_header_strip_outliers(
    grid: dict, panel_lookup: dict[int, dict], all_objects: list[dict]
) -> list[dict]:
    """For each panel in `grid`, find its header strip, compute the relative
    (dx, dy) offset to the panel top-left, and flag strips whose offset
    deviates from the peer median.

    Returns a list of fix dicts:
        {shape_id, name, left, top}  — target absolute position for the strip.
    """
    # Build per-panel best-strip candidate.
    strips_per_panel: list[tuple[dict, dict]] = []
    for panel_sid in grid["panel_ids"]:
        panel = panel_lookup.get(panel_sid)
        if panel is None:
            continue
        cands = _strip_candidates_for_panel(panel, all_objects)
        if not cands:
            continue
        # Prefer the candidate closest to the panel's top edge AND inside the
        # panel's x-range. Score = |dy| + |dx|.
        cands.sort(
            key=lambda c: abs(c["top"] - panel["top"]) + abs(c["left"] - panel["left"])
        )
        strips_per_panel.append((panel, cands[0]))

    if len(strips_per_panel) < 3:
        # Need at least 3 peer strips to vote on the median offset.
        return []

    rel_dx = [s["left"] - p["left"] for p, s in strips_per_panel]
    rel_dy = [s["top"] - p["top"] for p, s in strips_per_panel]
    med_dx = int(median(rel_dx))
    med_dy = int(median(rel_dy))

    # Also vote on width and height to detect oversized/undersized strips.
    rel_w = [s["width"] for _, s in strips_per_panel]
    rel_h = [s["height"] for _, s in strips_per_panel]
    med_w = int(median(rel_w))
    med_h = int(median(rel_h))

    fixes = []
    seen_ids = set()
    for panel, strip in strips_per_panel:
        if strip["shape_id"] in seen_ids:
            continue
        seen_ids.add(strip["shape_id"])
        dx = strip["left"] - panel["left"]
        dy = strip["top"] - panel["top"]
        wants: dict[str, int] = {}
        if abs(dx - med_dx) > HEADER_STRIP_OFFSET_TOL:
            wants["left"] = panel["left"] + med_dx
        if abs(dy - med_dy) > HEADER_STRIP_OFFSET_TOL:
            wants["top"] = panel["top"] + med_dy
        # Size voting: > 25% deviation from peer median = outlier
        if abs(strip["width"] - med_w) / max(med_w, 1) > 0.25:
            wants["width"] = med_w
        if abs(strip["height"] - med_h) / max(med_h, 1) > 0.25:
            wants["height"] = med_h
        if wants:
            fixes.append({
                "shape_id": strip["shape_id"],
                "name": strip["name"],
                **wants,
            })
    return fixes


# ----- diagnose-style API for integration with apply_repair -----

def diagnose_grid_repair(slide_inspection: dict, max_depth: int = 2) -> dict:
    """Return a repair plan structurally similar to _card_repair.diagnose_repair.

    For each detected grid we emit:
      - card_box_fixes: panels that need to snap to a row/col anchor
      - header_strip_fixes: header strips whose relative offset to their
        panel disagrees with the peer median (catches strips floating above
        their panel, oversized strips, etc.)
    """
    all_objects = slide_inspection["objects"]
    grids = detect_grids_nested(all_objects, max_depth=max_depth)
    rows_out = []
    for grid in grids:
        panel_lookup = {
            sid: next((o for o in all_objects if o["shape_id"] == sid), None)
            for sid in grid["panel_ids"]
        }
        # Box fixes from panel-grid outliers.
        box_fixes = []
        for o in grid["outliers"]:
            box_fixes.append({
                "shape_id": o["shape_id"],
                "name": o["name"],
                "left": o["target"][0],
                "top": o["target"][1],
                "row": o["row"], "col": o["col"],
                "depth": grid["depth"],
                "parent_panel_id": grid.get("parent_panel_id"),
            })
        # Header-strip outliers (must run AFTER virtually applying box fixes
        # so a panel that itself moved doesn't poison the strip detection).
        # We construct a virtual panel_lookup with corrected positions.
        virtual_lookup = {}
        for sid, panel in panel_lookup.items():
            if panel is None:
                continue
            box_fix = next((b for b in box_fixes if b["shape_id"] == sid), None)
            if box_fix:
                vp = dict(panel)
                if "left" in box_fix: vp["left"] = box_fix["left"]
                if "top" in box_fix: vp["top"] = box_fix["top"]
                virtual_lookup[sid] = vp
            else:
                virtual_lookup[sid] = panel
        header_fixes = _detect_header_strip_outliers(grid, virtual_lookup, all_objects)

        if not (box_fixes or header_fixes):
            continue

        rows_out.append({
            "card_shape_ids": grid["panel_ids"],
            "card_box_fixes": box_fixes,
            "header_strip_fixes": header_fixes,
            "orphan_relocations": [],
            "displaced_relocations": [],
            "_grid_meta": {
                "rows": grid["rows"], "cols": grid["cols"],
                "fill_ratio": grid["fill_ratio"],
                "depth": grid["depth"],
                "parent_panel_id": grid.get("parent_panel_id"),
            },
        })
    return {"rows": rows_out}
