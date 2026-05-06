from __future__ import annotations

import argparse
import json
from pathlib import Path
from statistics import median


# 1D fixed-epsilon clustering. Each cluster is a contiguous run of points whose
# consecutive gap stays within `eps_emu`. Adaptive thresholds are unsafe for dense
# layouts (many shapes packed close together inflate the median gap and merge
# unrelated columns), so we use a tight fixed tolerance instead.
def _cluster_1d(points: list[tuple[int, int]], eps_emu: int) -> list[list[int]]:
    """points: list of (value, item_id). Returns list of clusters, each a list of item_ids."""
    if not points:
        return []
    ordered = sorted(points, key=lambda p: p[0])
    if len(ordered) == 1:
        return [[ordered[0][1]]]

    clusters: list[list[int]] = [[ordered[0][1]]]
    last_value = ordered[0][0]
    for value, item_id in ordered[1:]:
        if value - last_value > eps_emu:
            clusters.append([item_id])
        else:
            clusters[-1].append(item_id)
        last_value = value
    return clusters


def _bbox_contains(outer: dict, inner: dict, slack_emu: int = 91440) -> bool:
    return (
        inner["left"] >= outer["left"] - slack_emu
        and inner["top"] >= outer["top"] - slack_emu
        and inner["left"] + inner["width"] <= outer["left"] + outer["width"] + slack_emu
        and inner["top"] + inner["height"] <= outer["top"] + outer["height"] + slack_emu
    )


def _classify_text_roles(slide: dict) -> dict[int, str]:
    """Assign role to each text shape: title / subtitle / badge / h2 / body / caption."""
    width = slide["width_emu"]
    height = slide["height_emu"]
    text_objs = [
        obj
        for obj in slide["objects"]
        if obj["kind"] == "text" and obj["text"] and not obj["anomalous"]
    ]
    if not text_objs:
        return {}

    # Primary font size per shape: max font size used (more reliable than first paragraph).
    def primary_size(obj: dict) -> int:
        return max(obj["font_sizes"]) if obj["font_sizes"] else 0

    sized = [(obj, primary_size(obj)) for obj in text_objs]
    sized_with_known = [(obj, sz) for obj, sz in sized if sz > 0]

    roles: dict[int, str] = {}

    # Title: largest font, top in upper 30% of slide.
    title_candidates = [
        (obj, sz)
        for obj, sz in sized_with_known
        if obj["top"] < height * 0.30
    ]
    title_obj = None
    if title_candidates:
        title_obj, _ = max(title_candidates, key=lambda p: (p[1], -p[0]["top"]))
        roles[title_obj["shape_id"]] = "title"

    # Badge: short text PINNED to the extreme top-right corner. The right edge
    # must sit close to the slide right edge AND the shape must be high up
    # (within ~10% of slide height). Anything further inside is treated as a
    # column header and left alone.
    badge_top_max = int(height * 0.10)
    badge_right_min = int(width * 0.82)
    for obj, sz in sized:
        if obj["shape_id"] in roles:
            continue
        right_edge = obj["left"] + obj["width"]
        if (
            right_edge > badge_right_min
            and obj["top"] < badge_top_max
            and len(obj["text"]) <= 24
            and obj["width"] < width * 0.25
        ):
            roles[obj["shape_id"]] = "badge"

    # Subtitle: 2nd largest font sitting just below the title and spanning wide.
    if title_obj is not None:
        title_bottom = title_obj["top"] + title_obj["height"]
        below_title = [
            (obj, sz)
            for obj, sz in sized_with_known
            if obj["shape_id"] not in roles
            and obj["top"] >= title_bottom - 91440
            and obj["top"] - title_bottom < height * 0.10
            and obj["width"] > width * 0.40
        ]
        if below_title:
            subtitle_obj, _ = max(below_title, key=lambda p: p[1])
            roles[subtitle_obj["shape_id"]] = "subtitle"

    # Remaining text: bucket into h2 / body / caption by font size clustering.
    remaining = [(obj, sz) for obj, sz in sized_with_known if obj["shape_id"] not in roles]
    if remaining:
        sizes_sorted = sorted({sz for _, sz in remaining})
        if len(sizes_sorted) == 1:
            uniform_role = "body"
            for obj, _ in remaining:
                roles[obj["shape_id"]] = uniform_role
        else:
            top_size = sizes_sorted[-1]
            bottom_size = sizes_sorted[0]
            mid_threshold_high = top_size - (top_size - bottom_size) * 0.25
            mid_threshold_low = bottom_size + (top_size - bottom_size) * 0.25
            for obj, sz in remaining:
                if sz >= mid_threshold_high:
                    roles[obj["shape_id"]] = "h2"
                elif sz <= mid_threshold_low:
                    roles[obj["shape_id"]] = "caption"
                else:
                    roles[obj["shape_id"]] = "body"

    # Text without known font size: default to body.
    for obj, sz in sized:
        if obj["shape_id"] not in roles and sz == 0:
            roles[obj["shape_id"]] = "body"

    return roles


def _detect_cards(slide: dict, roles: dict[int, str]) -> list[dict]:
    """A card = empty container shape that wraps at least one text shape."""
    containers = [
        obj
        for obj in slide["objects"]
        if obj["kind"] == "container" and not obj["anomalous"] and obj["width"] > 457200 and obj["height"] > 228600
    ]
    text_objs = [obj for obj in slide["objects"] if obj["kind"] == "text" and obj["text"]]
    cards: list[dict] = []
    for card_idx, container in enumerate(containers):
        contained_ids = [
            text["shape_id"]
            for text in text_objs
            if _bbox_contains(container, text)
        ]
        if contained_ids:
            cards.append(
                {
                    "card_id": card_idx,
                    "container_shape_id": container["shape_id"],
                    "contained_shape_ids": contained_ids,
                }
            )
    return cards


def _detect_rows_columns(slide: dict, roles: dict[int, str]) -> tuple[list[dict], list[dict]]:
    """Group peer text shapes into horizontal rows and vertical columns.

    Only peer shapes (same role family) participate. Title/subtitle/badge are excluded
    from row/column groups so they stay in their designed positions.
    """
    peer_roles = {"h2", "body", "caption"}
    peers = [
        obj
        for obj in slide["objects"]
        if obj["kind"] == "text"
        and obj["text"]
        and not obj["anomalous"]
        and roles.get(obj["shape_id"]) in peer_roles
    ]
    if not peers:
        return [], []

    by_id = {obj["shape_id"]: obj for obj in peers}

    # Row clustering uses `top` (not center) with a tight fixed tolerance: shapes
    # only count as the same row if their top edges land within ~0.08 inch of
    # each other.
    tops = [(obj["top"], obj["shape_id"]) for obj in peers]
    raw_row_clusters = _cluster_1d(tops, eps_emu=80000)
    rows = []
    for row_idx, ids in enumerate(raw_row_clusters):
        if len(ids) < 2:
            continue
        members = [by_id[i] for i in ids]
        rows.append(
            {
                "row_id": row_idx,
                "shape_ids": [m["shape_id"] for m in sorted(members, key=lambda m: m["left"])],
                "top_median": int(median(m["top"] for m in members)),
                "height_median": int(median(m["height"] for m in members)),
            }
        )

    # Column clustering uses `left` with the same tight tolerance — shapes that
    # actually share a left anchor (within ~0.08 inch) are a column. Centers and
    # adaptive thresholds incorrectly merge unrelated columns when slides are
    # densely packed.
    lefts = [(obj["left"], obj["shape_id"]) for obj in peers]
    raw_col_clusters = _cluster_1d(lefts, eps_emu=80000)
    cols = []
    for col_idx, ids in enumerate(raw_col_clusters):
        if len(ids) < 2:
            continue
        members = [by_id[i] for i in ids]
        cols.append(
            {
                "col_id": col_idx,
                "shape_ids": [m["shape_id"] for m in sorted(members, key=lambda m: m["top"])],
                "left_median": int(median(m["left"] for m in members)),
                "width_median": int(median(m["width"] for m in members)),
            }
        )

    return rows, cols


def detect_roles(inspection: dict) -> dict:
    out_slides: list[dict] = []
    for slide in inspection["slides"]:
        roles = _classify_text_roles(slide)
        cards = _detect_cards(slide, roles)
        rows, cols = _detect_rows_columns(slide, roles)

        shape_role_entries = []
        for obj in slide["objects"]:
            shape_role_entries.append(
                {
                    "shape_id": obj["shape_id"],
                    "name": obj["name"],
                    "kind": obj["kind"],
                    "role": roles.get(obj["shape_id"]),
                    "anomalous": obj["anomalous"],
                }
            )

        out_slides.append(
            {
                "slide_index": slide["slide_index"],
                "shapes": shape_role_entries,
                "cards": cards,
                "rows": rows,
                "columns": cols,
            }
        )

    return {"input": inspection["input"], "slides": out_slides}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--inspection", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    inspection = json.loads(Path(args.inspection).read_text(encoding="utf-8"))
    payload = detect_roles(inspection)
    Path(args.output).write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
