"""Auto-discover peer groups and analyze their spacing.

Existing infrastructure (`role_data.rows`, `role_data.columns`) can only
detect spacing problems within groups that detect_roles already
identified — typically peer cards in a clean horizontal row. Many real
decks have spacing problems in groups that detect_roles missed:
  * vertical lists where items aren't strict cards (icons + labels)
  * mixed-kind groups (text + container alternating)
  * grid cells in nested layouts
  * stacks of bullet points or feature rows

This module:
  1. Auto-discovers peer groups via geometric clustering (same kind,
     similar size, shared row/column band) without depending on roles.
  2. For each discovered group, computes a quantitative spacing report:
       * current gaps (sorted by axis position)
       * mean / median / std / max-min
       * gap-to-shape-width ratios (gives the agent "is this tight or
         loose?" context that absolute EMU numbers don't convey)
       * an ideal-equalized target = current median (most peers agree)
  3. Emits two issue categories:
       * `spacing-uneven`     — gap variance exceeds tolerance
       * `spacing-extreme`    — gaps are very tight (<10% avg width)
                                or very loose (>200% avg width)

The agent reads the report — including the "why" in each message —
then decides whether to apply the suggested mutate_argv.
"""
from __future__ import annotations

from statistics import mean, median, pstdev


# ---- tolerances ----

# Same row if tops differ by less than this (~0.22").
SPACING_TOP_BAND_EMU = 200000

# Same column if lefts differ by less.
SPACING_LEFT_BAND_EMU = 200000

# Peer if w and h are within 30% of base.
SPACING_SIZE_RATIO_TOL = 0.30

# Need ≥ 3 members to talk about spacing (need ≥ 2 gaps).
SPACING_MIN_GROUP = 3

# Gap variance: (max - min) / mean > this → uneven.
SPACING_UNEVEN_RATIO = 0.30

# Gap extreme: gap < this fraction of average shape width = too tight,
# > this multiple of average shape width = too loose.
SPACING_TIGHT_FRAC = 0.10
SPACING_LOOSE_FRAC = 2.00

# Slide-edge margin used to compute "would it fit centered?" hints.
SLIDE_MARGIN_EMU = 457200   # 0.5"


# ---- group discovery ----

def _candidate_shapes(objects: list[dict]) -> list[dict]:
    """Filter to non-anomalous, spacing-meaningful shapes."""
    return [
        o for o in objects
        if not o.get("anomalous")
        and o.get("kind") in ("container", "text", "shape")
        and o.get("width", 0) > 0
        and o.get("height", 0) > 0
    ]


def _is_peer(base: dict, other: dict) -> bool:
    """Two shapes are peers if same kind and similar size (within tol)."""
    if base["kind"] != other["kind"]:
        return False
    bw, bh = base["width"], base["height"]
    if abs(other["width"] - bw) / max(bw, 1) > SPACING_SIZE_RATIO_TOL:
        return False
    if abs(other["height"] - bh) / max(bh, 1) > SPACING_SIZE_RATIO_TOL:
        return False
    return True


def discover_row_groups(objects: list[dict]) -> list[list[dict]]:
    """Find horizontal peer groups: shapes sharing a top band, same kind,
    similar size, ≥ 3 members.
    """
    cands = _candidate_shapes(objects)
    used = set()
    groups: list[list[dict]] = []
    for i, base in enumerate(cands):
        if i in used:
            continue
        peers = [base]; used.add(i)
        for j, other in enumerate(cands):
            if j in used or j == i:
                continue
            if abs(other["top"] - base["top"]) > SPACING_TOP_BAND_EMU:
                continue
            if not _is_peer(base, other):
                continue
            peers.append(other); used.add(j)
        if len(peers) >= SPACING_MIN_GROUP:
            peers.sort(key=lambda p: p["left"])
            groups.append(peers)
    return groups


def discover_column_groups(objects: list[dict]) -> list[list[dict]]:
    """Find vertical peer groups: shapes sharing a left band, same kind,
    similar size, ≥ 3 members.
    """
    cands = _candidate_shapes(objects)
    used = set()
    groups: list[list[dict]] = []
    for i, base in enumerate(cands):
        if i in used:
            continue
        peers = [base]; used.add(i)
        for j, other in enumerate(cands):
            if j in used or j == i:
                continue
            if abs(other["left"] - base["left"]) > SPACING_LEFT_BAND_EMU:
                continue
            if not _is_peer(base, other):
                continue
            peers.append(other); used.add(j)
        if len(peers) >= SPACING_MIN_GROUP:
            peers.sort(key=lambda p: p["top"])
            groups.append(peers)
    return groups


# ---- gap analysis ----

def compute_gaps(peers: list[dict], axis: str) -> list[int]:
    """Return [g1, g2, ...] between consecutive peers on the given axis."""
    if axis == "horizontal":
        ordered = sorted(peers, key=lambda p: p["left"])
        return [
            ordered[i + 1]["left"] - (ordered[i]["left"] + ordered[i]["width"])
            for i in range(len(ordered) - 1)
        ]
    ordered = sorted(peers, key=lambda p: p["top"])
    return [
        ordered[i + 1]["top"] - (ordered[i]["top"] + ordered[i]["height"])
        for i in range(len(ordered) - 1)
    ]


def analyze_group(peers: list[dict], axis: str, slide_size_emu: int) -> dict:
    """Build a quantitative spacing report for one group.

    `slide_size_emu` is the slide's width or height (matching the axis).
    Returns a dict with metrics + `verdicts` listing detected problems.
    """
    gaps = compute_gaps(peers, axis)
    if not gaps:
        return {"axis": axis, "shape_ids": [p["shape_id"] for p in peers],
                "gaps": [], "verdicts": []}

    gap_mean = mean(gaps); gap_med = int(median(gaps))
    gap_std = pstdev(gaps) if len(gaps) > 1 else 0.0
    gap_min = min(gaps); gap_max = max(gaps)
    spread = gap_max - gap_min

    avg_dim = mean(p["width" if axis == "horizontal" else "height"] for p in peers)
    tight_threshold = avg_dim * SPACING_TIGHT_FRAC
    loose_threshold = avg_dim * SPACING_LOOSE_FRAC

    verdicts: list[dict] = []

    # 1. Uneven detection — gaps differ from the median by more than the
    #    tolerance ratio, OR spread is large in absolute terms.
    if gap_mean > 0 and (spread / max(gap_mean, 1)) > SPACING_UNEVEN_RATIO:
        verdicts.append({
            "kind": "uneven",
            "spread_emu": spread,
            "ratio": round(spread / gap_mean, 3),
            "target_gap_emu": gap_med,    # equalize to majority
            "rationale": (
                f"Gap spread is {spread} EMU ({round(100*spread/gap_mean,0)}% of mean); "
                f"median {gap_med} EMU represents the majority pattern."
            ),
        })

    # 2. Extreme spacing — even if gaps are UNIFORM, they may be too
    #    tight (visually crowded) or too loose (relationship lost).
    #    Skip if the group is already flagged uneven; the uneven fix
    #    addresses the root cause first, then re-detection on the next
    #    iteration will catch any remaining extreme spacing.
    already_uneven = any(v["kind"] == "uneven" for v in verdicts)
    if not already_uneven and gap_med < tight_threshold:
        verdicts.append({
            "kind": "too-tight",
            "current_gap_emu": gap_med,
            "min_recommended_emu": int(tight_threshold * 1.5),
            "rationale": (
                f"Median gap {gap_med} EMU is < 10% of average peer "
                f"{'width' if axis=='horizontal' else 'height'} "
                f"({int(avg_dim)} EMU); peers visually run together."
            ),
        })
    elif not already_uneven and gap_med > loose_threshold:
        verdicts.append({
            "kind": "too-loose",
            "current_gap_emu": gap_med,
            "max_recommended_emu": int(loose_threshold * 0.6),
            "rationale": (
                f"Median gap {gap_med} EMU is > 200% of average peer "
                f"{'width' if axis=='horizontal' else 'height'} "
                f"({int(avg_dim)} EMU); peers no longer read as a group."
            ),
        })

    return {
        "axis": axis,
        "shape_ids": [p["shape_id"] for p in peers],
        "n_peers": len(peers),
        "n_gaps": len(gaps),
        "gaps_emu": gaps,
        "gap_mean_emu": int(gap_mean),
        "gap_median_emu": gap_med,
        "gap_std_emu": int(gap_std),
        "gap_max_emu": gap_max,
        "gap_min_emu": gap_min,
        "avg_peer_dim_emu": int(avg_dim),
        "verdicts": verdicts,
    }


# ---- top-level: emit issues ----

def detect_spacing_issues(slide: dict) -> list[dict]:
    """Find all spacing problems on a slide and return issue dicts ready
    to be merged into score_layout's issue list."""
    objects = slide["objects"]
    width = slide["width_emu"]
    height = slide["height_emu"]

    issues: list[dict] = []
    seen_groups: set[tuple[int, ...]] = set()

    for axis, slide_dim, groups in (
        ("horizontal", width,  discover_row_groups(objects)),
        ("vertical",   height, discover_column_groups(objects)),
    ):
        for peers in groups:
            sids = tuple(sorted(p["shape_id"] for p in peers))
            key = (axis, *sids)
            if key in seen_groups:
                continue
            seen_groups.add(key)
            report = analyze_group(peers, axis, slide_dim)
            for v in report["verdicts"]:
                if v["kind"] == "uneven":
                    issues.append({
                        "category": "spacing-uneven",
                        "severity": "warning",
                        "shape_ids": list(report["shape_ids"]),
                        "shape_id": report["shape_ids"][0],
                        "message": (
                            f"{axis.capitalize()} spacing across "
                            f"{report['n_peers']} peers is uneven: "
                            f"gaps {report['gaps_emu']}; "
                            + v["rationale"]
                        ),
                        "suggested_fix": "equalize-gaps-auto",
                        "suggested_argv": [
                            "equalize-gaps",
                            "--shape-ids", ",".join(map(str, report["shape_ids"])),
                            "--axis", axis,
                            "--gap-emu", str(v["target_gap_emu"]),
                        ],
                        "spacing_report": report,
                    })
                elif v["kind"] in ("too-tight", "too-loose"):
                    target = v.get("min_recommended_emu") or v.get("max_recommended_emu")
                    issues.append({
                        "category": "spacing-extreme",
                        "severity": "info",
                        "shape_ids": list(report["shape_ids"]),
                        "shape_id": report["shape_ids"][0],
                        "message": (
                            f"{axis.capitalize()} gap is "
                            f"{'too tight' if v['kind']=='too-tight' else 'too loose'} "
                            f"across {report['n_peers']} peers. "
                            + v["rationale"]
                        ),
                        "suggested_fix": "equalize-gaps-auto",
                        "suggested_argv": [
                            "equalize-gaps",
                            "--shape-ids", ",".join(map(str, report["shape_ids"])),
                            "--axis", axis,
                            "--gap-emu", str(target),
                        ],
                        "spacing_report": report,
                    })
    return issues
