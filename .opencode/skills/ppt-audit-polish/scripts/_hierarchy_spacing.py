"""Vertical-rhythm analysis between hierarchical roles.

Typographic design has a well-known principle: the vertical gap between
hierarchical elements should be PROPORTIONAL to the larger element's
height. A title pushed too close to the body looks cramped; a title with
3x the title height of breathing room below it looks broken.

Rules of thumb (collected from typographic style guides; lenient ranges):
  * subtitle's top - title's bottom should be in 0.30-0.80 of title height
  * body's top - subtitle's bottom should be in 0.50-1.50 of subtitle height
  * body's top - title's bottom (if no subtitle) in 0.80-1.50 of title height

Inputs:
  * `slide`        — inspect_ppt slide entry (has objects, width_emu)
  * `role_data`    — detect_roles per-slide entry with shapes[role/shape_id]

Outputs:
  * Issue list with `category="hierarchy-spacing-drift"`. Each issue
    proposes a vertical NUDGE (positive dy = move down, negative dy =
    move up) for the LATER element so the ratio falls in the ideal range.

The detector is conservative: it only fires when the deck has a clear
single-section structure (one title + one subtitle/body run). Multi-
section slides are left untouched (would need section partitioning).
"""
from __future__ import annotations

# Ratio bounds. Below low → too tight (move LATER element DOWN).
# Above high → too loose (move LATER element UP).
TITLE_TO_SUBTITLE_LOW = 0.30
TITLE_TO_SUBTITLE_HIGH = 0.80
SUBTITLE_TO_BODY_LOW = 0.50
SUBTITLE_TO_BODY_HIGH = 1.50
TITLE_TO_BODY_LOW = 0.80
TITLE_TO_BODY_HIGH = 1.50

# Below this gap (EMU) we don't even compute ratios — it's overlap or
# touching, handled by other detectors.
MIN_MEASURABLE_GAP_EMU = 50000


def _role_lookup(role_data: dict) -> dict[str, list[int]]:
    """Return {role_name: [shape_id, ...]} from detect_roles output."""
    out: dict[str, list[int]] = {}
    for entry in role_data.get("shapes", []):
        role = entry.get("role")
        sid = entry.get("shape_id")
        if not role or sid is None:
            continue
        out.setdefault(role, []).append(sid)
    return out


def _bottom(shape: dict) -> int:
    return shape["top"] + shape["height"]


def _pick_top_role(roles: dict[str, list[int]], role: str,
                   obj_by_id: dict[int, dict]) -> dict | None:
    """For a role with possibly multiple shapes, pick the topmost one."""
    if role not in roles:
        return None
    candidates = [obj_by_id[sid] for sid in roles[role] if sid in obj_by_id]
    if not candidates:
        return None
    return min(candidates, key=lambda o: o["top"])


def _ideal_dy_to_target_ratio(
    above: dict, below: dict, target_low: float, target_high: float,
) -> int:
    """Compute the vertical nudge (dy) needed to bring the (below.top -
    above.bottom) / above.height ratio into [target_low, target_high].

    Snap to the closer boundary to minimize movement.
    """
    above_h = max(above["height"], 1)
    current_gap = below["top"] - _bottom(above)
    current_ratio = current_gap / above_h
    target = target_high if current_ratio > target_high else target_low
    target_gap = int(above_h * (target_low + target_high) / 2)
    return target_gap - current_gap


def detect_hierarchy_spacing_issues(slide: dict, role_data: dict) -> list[dict]:
    """Run the rhythm checks. Returns issue dicts ready for score_layout."""
    obj_by_id = {o["shape_id"]: o for o in slide["objects"]}
    roles = _role_lookup(role_data)

    title    = _pick_top_role(roles, "title", obj_by_id)
    subtitle = _pick_top_role(roles, "subtitle", obj_by_id)
    # body / heading roles are detected as different names by detect_roles;
    # consider both as "body candidates" and pick the topmost.
    body_candidates = []
    for r in ("body", "heading", "paragraph", "list"):
        if r in roles:
            body_candidates.extend(obj_by_id[s] for s in roles[r] if s in obj_by_id)
    body = min(body_candidates, key=lambda o: o["top"]) if body_candidates else None

    issues: list[dict] = []

    def _check_pair(above: dict, below: dict,
                    low: float, high: float, label: str) -> None:
        gap = below["top"] - _bottom(above)
        if gap < MIN_MEASURABLE_GAP_EMU:
            return
        above_h = max(above["height"], 1)
        ratio = gap / above_h
        if low <= ratio <= high:
            return
        dy = _ideal_dy_to_target_ratio(above, below, low, high)
        direction = "down" if dy > 0 else "up"
        issues.append({
            "category": "hierarchy-spacing-drift",
            "severity": "info",
            "shape_id": below["shape_id"],
            "message": (
                f"{label}: gap {gap} EMU is {round(ratio,2)}x of {label.split('-')[0]} "
                f"height {above_h} EMU (target range {low}-{high}x). "
                f"Suggest moving '{below.get('name')}' {direction} by {abs(dy)} EMU."
            ),
            "suggested_fix": "nudge-vertical",
            "suggested_argv": [
                "nudge",
                "--shape-id", str(below["shape_id"]),
                "--dx", "0",
                "--dy", str(dy),
            ],
            "rhythm_report": {
                "above_shape_id": above["shape_id"],
                "above_role": label.split("-")[0],
                "below_shape_id": below["shape_id"],
                "below_role": label.split("-")[2],
                "current_gap_emu": gap,
                "above_height_emu": above_h,
                "current_ratio": round(ratio, 3),
                "target_low": low, "target_high": high,
                "proposed_dy": dy,
            },
        })

    # Run the three checks in priority order.
    if title and subtitle:
        _check_pair(title, subtitle, TITLE_TO_SUBTITLE_LOW,
                    TITLE_TO_SUBTITLE_HIGH, "title-to-subtitle")
        if body:
            _check_pair(subtitle, body, SUBTITLE_TO_BODY_LOW,
                        SUBTITLE_TO_BODY_HIGH, "subtitle-to-body")
    elif title and body:
        _check_pair(title, body, TITLE_TO_BODY_LOW,
                    TITLE_TO_BODY_HIGH, "title-to-body")

    return issues
