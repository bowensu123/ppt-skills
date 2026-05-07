"""Agent-driven peer-group repair.

The existing `repair-grid` and `repair-peer-cards` ops detect peer
groups by GEOMETRIC clustering (similar size, shared row band).
That works ~80% of the time but mis-groups cases where:

  * A section title is geometrically similar to content cards (false
    positive — it gets pulled into the peer cohort and "repaired" to
    match the cards, breaking the design).
  * A deliberately-smaller card is treated as an outlier and resized
    (false negative — design intent destroyed).

This module flips the model: the AGENT decides what's a peer group
(reading the rendered slide + inspection.json + the shape inventory),
writes peer-groups.json, and this op applies geometric uniformity
WITHIN each group as instructed. No auto-detection, no false-positive
clustering — pure execution of the agent's categorization.

Children attached to each peer (icons, text inside cards, decorations
at fixed offsets) MOVE WITH their parent so visuals don't drift.
Pictures (extracted via _asset_extract) are preserved binary-faithful:
we only adjust their position, never re-encode.

peer-groups.json schema:

  {
    "groups": [
      {
        "name": "ai-tool-cards",
        "shape_ids": [10, 12, 14, 25, 27, 29],
        "axis": "horizontal",          // distribute axis
        "uniform_size": true,           // equalize W and H to median
        "uniform_spacing": true,        // distribute to even gaps
        "uniform_alignment": true,      // align cross-axis (top for h, left for v)
        "children_per_peer": [          // OPTIONAL - children that move with each peer
          [15, 18, 22],                 // sids inside peer at index 0
          [16, 19, 23],                 // ...peer at index 1
          ...
        ],
        "target_size": null,            // null = use median; or [w_emu, h_emu]
        "target_gap_emu": null          // null = use peer-median gap
      },
      ...
    ]
  }
"""
from __future__ import annotations

from statistics import median
from typing import Any


def _shape_index(prs) -> dict[int, Any]:
    """Map shape_id -> (slide, shape) across the deck."""
    out: dict[int, Any] = {}
    for slide in prs.slides:
        for shape in slide.shapes:
            sid = getattr(shape, "shape_id", None)
            if sid is not None:
                out[int(sid)] = (slide, shape)
    return out


def _move_with_children(shape, dx: int, dy: int, children: list,
                         action_log: list[dict], *, ctx: str) -> None:
    """Apply (dx, dy) to a parent shape AND every child shape provided.

    Pictures, text boxes, decorations all move uniformly. Children that
    moved as part of the parent's bbox keep their relative position.
    """
    from pptx.util import Emu
    if dx == 0 and dy == 0:
        return
    try:
        shape.left = Emu(int(shape.left or 0) + dx)
        shape.top = Emu(int(shape.top or 0) + dy)
    except (AttributeError, ValueError):
        return
    moved = 1
    for child in children:
        try:
            child.left = Emu(int(child.left or 0) + dx)
            child.top = Emu(int(child.top or 0) + dy)
            moved += 1
        except (AttributeError, ValueError):
            continue
    action_log.append({
        "action": "move-with-children",
        "ctx": ctx,
        "shape_id": shape.shape_id,
        "dx": dx, "dy": dy,
        "children_moved": moved - 1,
    })


def _equalize_size(shapes: list, target_w: int, target_h: int,
                    action_log: list[dict], group_name: str) -> None:
    """Resize each shape to (target_w, target_h)."""
    from pptx.util import Emu
    for s in shapes:
        try:
            old_w = int(s.width or 0); old_h = int(s.height or 0)
            if old_w != target_w:
                s.width = Emu(target_w)
            if old_h != target_h:
                s.height = Emu(target_h)
            if old_w != target_w or old_h != target_h:
                action_log.append({
                    "action": "equalize-size",
                    "group": group_name,
                    "shape_id": s.shape_id,
                    "from": [old_w, old_h],
                    "to": [target_w, target_h],
                })
        except (AttributeError, ValueError):
            continue


def _distribute_uniformly(shapes_with_children: list[tuple],
                          axis: str, target_gap: int,
                          action_log: list[dict],
                          group_name: str) -> None:
    """Place peers along `axis` with uniform `target_gap`. Children
    move with their parent (uniformly so relative positions hold)."""
    from pptx.util import Emu
    # Sort peers by their current axis position (left for h, top for v)
    if axis == "horizontal":
        ordered = sorted(shapes_with_children,
                         key=lambda p: int(p[0].left or 0))
    else:
        ordered = sorted(shapes_with_children,
                         key=lambda p: int(p[0].top or 0))
    if not ordered:
        return
    # First peer stays at its current position.
    first_shape, first_children = ordered[0]
    cursor: int
    if axis == "horizontal":
        cursor = int(first_shape.left or 0) + int(first_shape.width or 0) + target_gap
    else:
        cursor = int(first_shape.top or 0) + int(first_shape.height or 0) + target_gap

    for shape, children in ordered[1:]:
        if axis == "horizontal":
            target_left = cursor
            old_left = int(shape.left or 0)
            dx = target_left - old_left
            _move_with_children(shape, dx, 0, children, action_log,
                                ctx=f"distribute:{group_name}")
            cursor = target_left + int(shape.width or 0) + target_gap
        else:
            target_top = cursor
            old_top = int(shape.top or 0)
            dy = target_top - old_top
            _move_with_children(shape, 0, dy, children, action_log,
                                ctx=f"distribute:{group_name}")
            cursor = target_top + int(shape.height or 0) + target_gap


def _align_cross_axis(shapes_with_children: list[tuple], axis: str,
                       action_log: list[dict], group_name: str) -> None:
    """For horizontal-axis groups, align all `top`s to the median.
    For vertical-axis groups, align all `left`s to the median.
    Children follow."""
    if not shapes_with_children:
        return
    if axis == "horizontal":
        target_top = int(median(int(s.top or 0)
                                  for s, _ in shapes_with_children))
        for shape, children in shapes_with_children:
            old = int(shape.top or 0)
            if old != target_top:
                _move_with_children(shape, 0, target_top - old, children,
                                    action_log,
                                    ctx=f"align-cross:{group_name}")
    else:
        target_left = int(median(int(s.left or 0)
                                   for s, _ in shapes_with_children))
        for shape, children in shapes_with_children:
            old = int(shape.left or 0)
            if old != target_left:
                _move_with_children(shape, target_left - old, 0, children,
                                    action_log,
                                    ctx=f"align-cross:{group_name}")


def repair_peers_smart(prs, peer_groups: dict) -> dict:
    """Apply uniformity per agent-defined peer group. Returns action log."""
    sid_index = _shape_index(prs)
    actions: list[dict] = []
    skipped: list[dict] = []

    for group in peer_groups.get("groups", []):
        name = group.get("name", "<unnamed>")
        sids = group.get("shape_ids", [])
        children_per_peer = group.get("children_per_peer", [])
        axis = group.get("axis", "horizontal")
        if axis not in ("horizontal", "vertical"):
            skipped.append({"group": name, "reason": f"invalid-axis:{axis}"})
            continue

        # Resolve shapes; missing shape_ids skipped silently.
        shapes_with_children: list[tuple] = []
        for i, sid in enumerate(sids):
            if sid not in sid_index:
                continue
            _, shape = sid_index[sid]
            children = []
            if i < len(children_per_peer):
                for child_sid in children_per_peer[i]:
                    if child_sid in sid_index:
                        children.append(sid_index[child_sid][1])
            shapes_with_children.append((shape, children))
        if len(shapes_with_children) < 2:
            skipped.append({"group": name, "reason": "fewer-than-2-shapes"})
            continue

        peers_only = [s for s, _ in shapes_with_children]

        # Uniform size
        if group.get("uniform_size", True):
            target = group.get("target_size")
            if target and len(target) == 2:
                tw, th = int(target[0]), int(target[1])
            else:
                tw = int(median(int(s.width or 0) for s in peers_only))
                th = int(median(int(s.height or 0) for s in peers_only))
            _equalize_size(peers_only, tw, th, actions, name)

        # Uniform alignment (cross-axis)
        if group.get("uniform_alignment", True):
            _align_cross_axis(shapes_with_children, axis, actions, name)

        # Uniform spacing (along axis)
        if group.get("uniform_spacing", True):
            target_gap = group.get("target_gap_emu")
            if target_gap is None:
                # Compute peer-median gap on current positions.
                if axis == "horizontal":
                    sorted_peers = sorted(
                        shapes_with_children,
                        key=lambda p: int(p[0].left or 0),
                    )
                    gaps = []
                    for i in range(len(sorted_peers) - 1):
                        a, _ = sorted_peers[i]; b, _ = sorted_peers[i + 1]
                        gaps.append(
                            int(b.left or 0)
                            - (int(a.left or 0) + int(a.width or 0))
                        )
                else:
                    sorted_peers = sorted(
                        shapes_with_children,
                        key=lambda p: int(p[0].top or 0),
                    )
                    gaps = []
                    for i in range(len(sorted_peers) - 1):
                        a, _ = sorted_peers[i]; b, _ = sorted_peers[i + 1]
                        gaps.append(
                            int(b.top or 0)
                            - (int(a.top or 0) + int(a.height or 0))
                        )
                if gaps:
                    target_gap = int(median(gaps))
                else:
                    target_gap = 91440
            target_gap = max(int(target_gap), 0)
            _distribute_uniformly(shapes_with_children, axis, target_gap,
                                   actions, name)

    return {
        "groups_processed": len(peer_groups.get("groups", []) ) - len(skipped),
        "actions_applied": len(actions),
        "actions": actions,
        "skipped": skipped,
    }
