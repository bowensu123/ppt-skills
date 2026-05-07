"""Peer-card outlier detection and repair.

Many decks have a "card row" of N similarly-shaped cards (e.g., a
five-step process diagram). When one card or one shape inside a card
drifts away from its peers - oversized header strip, misplaced icon,
container that extends past the card row's vertical band - the result
looks broken even though the typography is fine.

This module:
  1. Identifies a peer-card row (containers with similar height/width and
     close top edges).
  2. For each container, derives the "peer template" (median top, height,
     width across the row).
  3. Snaps any outlier container to the template.
  4. For shapes inside each card (bbox-center inside container), computes
     each shape's offset *relative* to its card's top-left.
  5. For each role-equivalent shape across cards (same z-order index
     among cards' children), uses the peer-median relative offset as the
     ground truth and snaps outliers to it.

The algorithm is deliberately conservative: only shapes that are clearly
peers (same kind, similar size) participate; anything ambiguous is left
untouched.
"""
from __future__ import annotations

from statistics import median
from typing import Any


PEER_HEIGHT_RATIO_TOL = 1.0    # generous inclusion (we WANT outliers in the row so we can fix them)
PEER_WIDTH_RATIO_TOL = 1.0
PEER_TOP_BAND_EMU = 600000     # cards whose top is within ~0.65" of base count as same-row
OUTLIER_SIZE_RATIO = 0.20      # >20% deviation from peer median = outlier
OUTLIER_OFFSET_EMU = 250000    # >0.27" relative offset from peer median = outlier


def _bbox_center(obj: dict) -> tuple[int, int]:
    return obj["left"] + obj["width"] // 2, obj["top"] + obj["height"] // 2


def _bbox_contains(outer: dict, inner_center: tuple[int, int], slack: int = 91440) -> bool:
    cx, cy = inner_center
    return (
        outer["left"] - slack <= cx <= outer["left"] + outer["width"] + slack
        and outer["top"] - slack <= cy <= outer["top"] + outer["height"] + slack
    )


def _identify_peer_card_rows(slide_objects: list[dict]) -> list[list[dict]]:
    """Cluster card-like containers (>= 1M EMU tall, same band, similar size)."""
    candidates = [
        obj for obj in slide_objects
        if obj["kind"] == "container"
        and not obj.get("anomalous")
        and obj["width"] >= 800000
        and obj["height"] >= 1000000
    ]
    if len(candidates) < 2:
        return []

    rows: list[list[dict]] = []
    used = set()
    for i, base in enumerate(candidates):
        if i in used:
            continue
        peers = [base]
        used.add(i)
        base_h = base["height"]
        base_w = base["width"]
        base_top = base["top"]
        for j, other in enumerate(candidates):
            if j in used:
                continue
            if abs(other["top"] - base_top) > PEER_TOP_BAND_EMU:
                continue
            if abs(other["height"] - base_h) / max(base_h, 1) > PEER_HEIGHT_RATIO_TOL:
                continue
            if abs(other["width"] - base_w) / max(base_w, 1) > PEER_WIDTH_RATIO_TOL:
                continue
            peers.append(other)
            used.add(j)
        if len(peers) >= 3:
            rows.append(peers)
    return rows


def _identify_card_children(card: dict, all_objects: list[dict]) -> list[dict]:
    """Return shapes whose bbox-center sits inside this card."""
    return [
        obj for obj in all_objects
        if obj["shape_id"] != card["shape_id"]
        and not obj.get("anomalous")
        and _bbox_contains(card, _bbox_center(obj))
    ]


def _identify_outlier_children(
    cards: list[dict], all_objects: list[dict]
) -> list[dict]:
    """For each card, find children whose center is OUTSIDE this card but
    might belong to it based on similar children in peer cards.

    Heuristic: walk peer cards in left-to-right order. For each card, the
    set of children should mirror peers'. If one card has fewer children
    inside than peers do at the same relative slot, look for orphan shapes
    (centers outside any card) whose offset matches peer slots.
    """
    cards_lr = sorted(cards, key=lambda c: c["left"])
    children_per_card = [_identify_card_children(c, all_objects) for c in cards_lr]

    # Build peer slot signatures: for each child role-equivalence index,
    # the median (rel_x, rel_y, w, h, kind, fill_hex bucket).
    # We approximate role-equivalence by sorting each card's children by (rel_y, rel_x).
    sorted_children = []
    for card, kids in zip(cards_lr, children_per_card):
        rel = sorted(
            [
                (
                    kid,
                    kid["left"] - card["left"],
                    kid["top"] - card["top"],
                )
                for kid in kids
            ],
            key=lambda r: (r[2], r[1]),
        )
        sorted_children.append(rel)

    # Find orphan shapes - bbox-center outside ALL cards but inside the
    # vertical band of the row.
    band_top = min(c["top"] for c in cards_lr) - 200000
    band_bottom = max(c["top"] + c["height"] for c in cards_lr) + 200000
    candidate_orphans = []
    for obj in all_objects:
        if obj.get("anomalous"):
            continue
        cx, cy = _bbox_center(obj)
        if not (band_top <= cy <= band_bottom):
            continue
        in_any_card = any(_bbox_contains(c, (cx, cy)) for c in cards_lr)
        if not in_any_card:
            candidate_orphans.append(obj)
    return candidate_orphans


def _suggest_card_box_fixes(cards: list[dict]) -> list[dict]:
    """For each card, suggest top/height/width corrections to match peer median."""
    if len(cards) < 3:
        return []
    med_top = int(median(c["top"] for c in cards))
    med_h = int(median(c["height"] for c in cards))
    med_w = int(median(c["width"] for c in cards))
    fixes = []
    for c in cards:
        wants: dict[str, Any] = {}
        if abs(c["top"] - med_top) > OUTLIER_OFFSET_EMU:
            wants["top"] = med_top
        if abs(c["height"] - med_h) / max(med_h, 1) > OUTLIER_SIZE_RATIO:
            wants["height"] = med_h
        if abs(c["width"] - med_w) / max(med_w, 1) > OUTLIER_SIZE_RATIO:
            wants["width"] = med_w
        if wants:
            fixes.append({"shape_id": c["shape_id"], "name": c["name"], **wants})
    return fixes


def _suggest_header_strip_fixes(cards: list[dict], all_objects: list[dict]) -> list[dict]:
    """Header strips are small filled rectangles in the top-left of each
    card (e.g., a numbered badge). Width is consistent across peers; an
    outlier whose width is >> peer median should be shrunk.

    The outlier may have grown so wide that its bbox center has slipped
    OUTSIDE its parent card; we still want to find it. We do this by
    searching the entire row top-band, not only "inside-card" children.
    """
    # Step 1: per-card peer badges (children whose center IS inside a card).
    headers_per_card: list[list[dict]] = []
    for card in cards:
        kids = _identify_card_children(card, all_objects)
        candidates = [
            k for k in kids
            if k["kind"] == "container"
            and k.get("fill_hex")
            and k["top"] - card["top"] < card["height"] * 0.20
            and k["height"] < card["height"] * 0.20
            and k["width"] < card["width"] * 0.30
        ]
        headers_per_card.append(candidates)

    badges = []
    for headers in headers_per_card:
        if headers:
            badges.append(min(headers, key=lambda h: h["width"]))
    if len(badges) < 3:
        return []
    median_w = int(median(b["width"] for b in badges))
    median_h = int(median(b["height"] for b in badges))

    # Step 2: also sweep the row's top-band for filled rects that look like
    # outlier badges - too wide to fit the peer pattern but in the right band.
    band_top = min(c["top"] for c in cards) - 100000
    band_max_top = min(c["top"] for c in cards) + int(median(c["height"] for c in cards)) // 4
    seen = set()
    fixes = []
    for headers in headers_per_card:
        for header in headers:
            if header["shape_id"] in seen:
                continue
            seen.add(header["shape_id"])
            if header["width"] > median_w * 2 and header["height"] < median_h * 2.5:
                fixes.append(
                    {
                        "shape_id": header["shape_id"],
                        "name": header["name"],
                        "width": median_w,
                        "height": median_h,
                    }
                )

    for obj in all_objects:
        sid = obj["shape_id"]
        if sid in seen:
            continue
        if obj["kind"] != "container" or not obj.get("fill_hex"):
            continue
        if obj.get("anomalous"):
            continue
        if not (band_top <= obj["top"] <= band_max_top):
            continue
        if not (median_h * 0.5 <= obj["height"] <= median_h * 2.5):
            continue
        if obj["width"] > median_w * 2:
            seen.add(sid)
            fixes.append(
                {"shape_id": sid, "name": obj["name"], "width": median_w, "height": median_h}
            )
    return fixes


def _suggest_orphan_relocation(cards: list[dict], all_objects: list[dict]) -> list[dict]:
    """For shapes whose center is outside every card in the row but lies
    in the row's vertical band, snap to the nearest card whose relative
    slot would receive them best.

    For each peer card we have a list of (child, rel_x, rel_y). Build the
    median rel_x/rel_y for each role-equivalent slot; pick the orphan's
    nearest slot and place it accordingly.
    """
    cards_lr = sorted(cards, key=lambda c: c["left"])
    if not cards_lr:
        return []
    children_per_card = [_identify_card_children(c, all_objects) for c in cards_lr]
    # Sort each card's children by (rel_y, rel_x) so similar-position children align across cards.
    rel_sorted_per_card = []
    for card, kids in zip(cards_lr, children_per_card):
        rel = sorted(
            [(kid, kid["left"] - card["left"], kid["top"] - card["top"]) for kid in kids],
            key=lambda r: (r[2], r[1]),
        )
        rel_sorted_per_card.append(rel)

    # Build peer template from cards that have all expected children.
    max_slots = max((len(rs) for rs in rel_sorted_per_card), default=0)
    if max_slots == 0:
        return []

    band_top = min(c["top"] for c in cards_lr) - 200000
    band_bottom = max(c["top"] + c["height"] for c in cards_lr) + 200000

    fixes = []
    for orphan in all_objects:
        if orphan.get("anomalous"):
            continue
        if orphan["kind"] not in ("text", "container"):
            continue
        cx, cy = _bbox_center(orphan)
        if not (band_top <= cy <= band_bottom):
            continue
        if any(_bbox_contains(c, (cx, cy)) for c in cards_lr):
            continue
        # This is an orphan in the row band. Decide which card it belongs to.
        # Strategy: use the orphan's text or fill to find a matching peer slot.
        best_card = None
        best_slot_rel: tuple[int, int] | None = None
        best_score = -1.0
        for ci, (card, rs) in enumerate(zip(cards_lr, rel_sorted_per_card)):
            for slot_idx, (peer_kid, rx, ry) in enumerate(rs):
                if peer_kid["kind"] != orphan["kind"]:
                    continue
                if peer_kid["text"] and orphan["text"]:
                    # text similarity by character length proxy
                    sim = 1.0 - abs(len(peer_kid["text"]) - len(orphan["text"])) / max(
                        len(peer_kid["text"]), len(orphan["text"]), 1
                    )
                    if sim > best_score:
                        # Find nearest card MISSING this slot
                        for cj, ors in enumerate(rel_sorted_per_card):
                            if cj == ci:
                                continue
                            existing_at_slot = next(
                                (
                                    o for o, ox, oy in ors
                                    if abs(ox - rx) < 91440 and abs(oy - ry) < 91440
                                ),
                                None,
                            )
                            if existing_at_slot is None:
                                target_card = cards_lr[cj]
                                target_left = target_card["left"] + rx
                                target_top = target_card["top"] + ry
                                # Score this match
                                if sim > best_score:
                                    best_score = sim
                                    best_card = target_card
                                    best_slot_rel = (target_left, target_top)
                                break
        if best_card and best_slot_rel and best_score > 0.5:
            fixes.append(
                {
                    "shape_id": orphan["shape_id"],
                    "name": orphan["name"],
                    "left": best_slot_rel[0],
                    "top": best_slot_rel[1],
                    "match_score": round(best_score, 2),
                }
            )
    return fixes


def _suggest_displaced_children(cards: list[dict], all_objects: list[dict]) -> list[dict]:
    """Detect shapes that bbox-inside card[i] but structurally belong to
    card[j] (because card[i] already has its own equivalent shape and
    card[j] is missing one).

    Approach:
      1. For each card, list children sorted by relative (top, left).
      2. Compute peer median child-count K.
      3. For cards with more than K children: identify the "extra" by
         looking for two children in the same approximate slot.
      4. For each extra, find the under-populated card whose missing slot
         would receive it, using peer templates.
    """
    cards_lr = sorted(cards, key=lambda c: c["left"])
    children_per_card = [_identify_card_children(c, all_objects) for c in cards_lr]
    counts = [len(kids) for kids in children_per_card]
    if len(counts) < 3:
        return []

    median_count = int(median(counts))

    # Build peer slot template from a card with >= median children.
    template_card = None
    template_kids = []
    for c, kids in zip(cards_lr, children_per_card):
        if len(kids) >= median_count and (template_card is None or len(kids) == median_count):
            template_card = c
            template_kids = kids
            break
    if template_card is None or not template_kids:
        return []

    template_slots = sorted(
        [
            (k["left"] - template_card["left"], k["top"] - template_card["top"], k)
            for k in template_kids
        ],
        key=lambda r: (r[1], r[0]),
    )

    def slot_present(card_idx: int, rx: int, ry: int, kind: str, slack: int = 200000) -> bool:
        c = cards_lr[card_idx]
        for k in children_per_card[card_idx]:
            if k["kind"] != kind:
                continue
            if abs((k["left"] - c["left"]) - rx) <= slack and abs((k["top"] - c["top"]) - ry) <= slack:
                return True
        return False

    # First, mark each kid as "fits a template slot" or "extra" within its
    # own card. Only "extra" kids are candidates to relocate.
    def is_template_fit(card, kid, slack: int = 200000) -> bool:
        krx = kid["left"] - card["left"]
        kry = kid["top"] - card["top"]
        for rx, ry, ref_k in template_slots:
            if kid["kind"] != ref_k["kind"]:
                continue
            if abs(krx - rx) <= slack and abs(kry - ry) <= slack:
                return True
        return False

    extras_per_card: list[list[dict]] = []
    for c, kids in zip(cards_lr, children_per_card):
        extras_per_card.append([k for k in kids if not is_template_fit(c, k)])

    fixes = []
    used = set()
    for under_idx, c in enumerate(cards_lr):
        if len(children_per_card[under_idx]) >= median_count:
            continue
        missing = []
        for rx, ry, ref_k in template_slots:
            if not slot_present(under_idx, rx, ry, ref_k["kind"]):
                missing.append((rx, ry, ref_k))
        if not missing:
            continue
        for rx, ry, ref_k in missing:
            best = None
            best_score = -1.0
            for over_idx, extras in enumerate(extras_per_card):
                if over_idx == under_idx:
                    continue
                if not extras:
                    continue
                over_card = cards_lr[over_idx]
                for kid in extras:
                    if kid["shape_id"] in used:
                        continue
                    if kid["kind"] != ref_k["kind"]:
                        continue
                    # Compute the kid's intent relative position: would it
                    # have FIT the missing slot if this kid were already in
                    # the under-populated card? Use the kid's absolute
                    # offset from over_card and check size/text similarity.
                    size_sim = 1.0 - abs(kid["width"] - ref_k["width"]) / max(ref_k["width"], 1)
                    text_sim = 1.0
                    if ref_k.get("text") and kid.get("text"):
                        text_sim = 1.0 - abs(len(ref_k["text"]) - len(kid["text"])) / max(
                            len(ref_k["text"]), len(kid["text"]), 1
                        )
                    score = 0.5 * size_sim + 0.5 * text_sim
                    if score > best_score:
                        best_score = score
                        best = (kid, over_idx)
            if best and best_score > 0.4:
                kid, _over_idx = best
                used.add(kid["shape_id"])
                fixes.append({
                    "shape_id": kid["shape_id"],
                    "name": kid["name"],
                    "left": c["left"] + rx,
                    "top": c["top"] + ry,
                    "match_score": round(best_score, 2),
                })
    return fixes


def diagnose_repair(slide_inspection: dict, scope: str = "all", include_grid: bool = True) -> dict:
    """Build a repair plan for one slide.

    scope:
      'all'           — every fix kind (default; matches previous behavior)
      'safe'          — only card-box-fix + header-strip-fix
                        (skips orphan + displaced relocation, which can
                        misfire on heavily damaged decks)
      'no-orphans'    — everything except orphan-relocation

    include_grid:
      When the 1D row-based detection produces no full row of >=3 peers,
      fall back to the 2D grid detector (handles 2x3 / 3x2 dashboards
      and nested layouts).
    """
    # Work on a deep-copied object list so we can apply fixes virtually
    # before running later detection passes (otherwise an oversized header
    # whose center sits OUTSIDE its real card poisons displaced-child
    # detection).
    import copy
    objs = copy.deepcopy(slide_inspection["objects"])
    obj_by_id = {o["shape_id"]: o for o in objs}

    rows = _identify_peer_card_rows(objs)
    out_rows = []
    for cards in rows:
        box_fixes = _suggest_card_box_fixes(cards)
        for fix in box_fixes:
            o = obj_by_id.get(fix["shape_id"])
            if o is not None:
                if "top" in fix: o["top"] = fix["top"]
                if "height" in fix: o["height"] = fix["height"]
                if "width" in fix: o["width"] = fix["width"]

        header_fixes = _suggest_header_strip_fixes(cards, objs)
        for fix in header_fixes:
            o = obj_by_id.get(fix["shape_id"])
            if o is not None:
                o["width"] = fix["width"]
                if "height" in fix:
                    o["height"] = fix["height"]

        # After in-memory updates, recompute card list with updated geometry.
        updated_cards = [obj_by_id[c["shape_id"]] for c in cards]

        orphan_relocs = []
        displaced_relocs = []
        if scope == "all":
            orphan_relocs = _suggest_orphan_relocation(updated_cards, objs)
            displaced_relocs = _suggest_displaced_children(updated_cards, objs)
        elif scope == "no-orphans":
            displaced_relocs = _suggest_displaced_children(updated_cards, objs)
        # 'safe' produces only box + header fixes.

        out_rows.append({
            "card_shape_ids": [c["shape_id"] for c in cards],
            "card_box_fixes": box_fixes,
            "header_strip_fixes": header_fixes,
            "orphan_relocations": orphan_relocs,
            "displaced_relocations": displaced_relocs,
        })

    # If 1D row detection produced no actionable plan, fall back to 2D grid.
    if include_grid and not any(
        r["card_box_fixes"] or r["header_strip_fixes"] or r["orphan_relocations"] or r["displaced_relocations"]
        for r in out_rows
    ):
        try:
            from _grid_detect import diagnose_grid_repair
            grid_plan = diagnose_grid_repair(slide_inspection)
            out_rows.extend(grid_plan["rows"])
        except Exception:
            pass

    return {"rows": out_rows}


def apply_repair(slide, slide_inspection: dict, plan: dict, action_log: list[dict], slide_idx: int) -> None:
    """Apply a repair plan to a python-pptx slide in place.

    For grid-box fixes that include a `left`/`top` delta, also move every
    shape that was nested INSIDE the panel's old bbox by the same delta —
    otherwise children stay at their old absolute positions and appear as
    orphaned leftovers when the panel moves.
    """
    from pptx.util import Emu

    sid_to_shape = {int(s.shape_id): s for s in slide.shapes if s.shape_id is not None}
    obj_lookup = {o["shape_id"]: o for o in slide_inspection["objects"]}

    def _bbox_contains_center(L, T, R, B, obj, slack=91440):
        cx = obj["left"] + obj["width"] // 2
        cy = obj["top"] + obj["height"] // 2
        return L - slack <= cx <= R + slack and T - slack <= cy <= B + slack

    def _panel_attached_children(old_panel: dict, new_left: int, new_top: int) -> list[int]:
        """Return shape_ids that are children of OLD panel position but
        would NOT be inside the NEW panel position. These are "attached"
        decorations (header strips, badges) that should travel with the
        panel. Shapes that fit BOTH old and new are independent content
        already correctly placed — leave them alone.
        """
        old_L = old_panel["left"]; old_T = old_panel["top"]
        old_R = old_L + old_panel["width"]; old_B = old_T + old_panel["height"]
        new_R = new_left + old_panel["width"]; new_B = new_top + old_panel["height"]
        out = []
        for o in slide_inspection["objects"]:
            if o["shape_id"] == old_panel["shape_id"]:
                continue
            if o.get("anomalous"):
                continue
            in_old = _bbox_contains_center(old_L, old_T, old_R, old_B, o)
            in_new = _bbox_contains_center(new_left, new_top, new_R, new_B, o)
            if in_old and not in_new:
                out.append(o["shape_id"])
        return out

    for row in plan["rows"]:
        for fix in row["card_box_fixes"]:
            shape = sid_to_shape.get(fix["shape_id"])
            if shape is None:
                continue
            old_obj = obj_lookup.get(fix["shape_id"])
            old_left = int(shape.left) if shape.left is not None else 0
            old_top = int(shape.top) if shape.top is not None else 0

            # Identify panel-attached decorations (header strip, badge,
            # accent bar) that lived in the panel's OLD bbox but won't fit
            # the NEW position. These travel with the panel. Content that
            # fits both old and new bboxes is left alone — it's already at
            # the correct slide coordinate.
            children_to_move: list[int] = []
            new_left_for_calc = fix.get("left", old_left)
            new_top_for_calc = fix.get("top", old_top)
            if old_obj is not None and ("left" in fix or "top" in fix):
                children_to_move = _panel_attached_children(
                    old_obj, new_left_for_calc, new_top_for_calc,
                )
            dx = (fix["left"] - old_left) if "left" in fix else 0
            dy = (fix["top"] - old_top) if "top" in fix else 0

            if "top" in fix:
                shape.top = Emu(fix["top"])
            if "height" in fix:
                shape.height = Emu(fix["height"])
            if "width" in fix:
                shape.width = Emu(fix["width"])
            if "left" in fix:
                shape.left = Emu(fix["left"])

            if (dx or dy) and children_to_move:
                moved_n = 0
                for child_sid in children_to_move:
                    child_shape = sid_to_shape.get(int(child_sid))
                    if child_shape is None:
                        continue
                    try:
                        child_shape.left = Emu(int(child_shape.left) + dx)
                        child_shape.top = Emu(int(child_shape.top) + dy)
                        moved_n += 1
                    except (AttributeError, ValueError):
                        continue
                if moved_n:
                    action_log.append({
                        "slide_index": slide_idx, "action": "repair-card-box-children",
                        "target": fix["name"],
                        "detail": f"moved {moved_n} child shapes by ({dx}, {dy})",
                    })

            action_log.append({
                "slide_index": slide_idx, "action": "repair-card-box",
                "target": fix["name"], "detail": str({k: v for k, v in fix.items() if k not in ("shape_id", "name")}),
            })
        for fix in row["header_strip_fixes"]:
            shape = sid_to_shape.get(fix["shape_id"])
            if shape is None:
                continue
            shape.width = Emu(fix["width"])
            if "height" in fix:
                shape.height = Emu(fix["height"])
            action_log.append({
                "slide_index": slide_idx, "action": "repair-header-strip",
                "target": fix["name"],
                "detail": f"width={fix['width']}{' height=' + str(fix['height']) if 'height' in fix else ''}",
            })
        for fix in row["orphan_relocations"]:
            shape = sid_to_shape.get(fix["shape_id"])
            if shape is None:
                continue
            shape.left = Emu(fix["left"])
            shape.top = Emu(fix["top"])
            action_log.append({
                "slide_index": slide_idx, "action": "repair-orphan-relocate",
                "target": fix["name"],
                "detail": f"left={fix['left']} top={fix['top']} score={fix['match_score']}",
            })
        for fix in row.get("displaced_relocations", []):
            shape = sid_to_shape.get(fix["shape_id"])
            if shape is None:
                continue
            shape.left = Emu(fix["left"])
            shape.top = Emu(fix["top"])
            action_log.append({
                "slide_index": slide_idx, "action": "repair-displaced-child",
                "target": fix["name"],
                "detail": f"left={fix['left']} top={fix['top']} score={fix['match_score']}",
            })
