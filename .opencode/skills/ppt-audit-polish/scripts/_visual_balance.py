"""Visual balance: where is the slide's center of mass?

Treats every shape as a uniform-density rectangle and computes the
area-weighted centroid. Distance from slide center, normalized by the
slide half-extent, is the imbalance score.

Rule of thirds (stretch goal): also report whether the dominant shape's
center sits near a third-line — but for v1 we focus on overall balance.
"""
from __future__ import annotations


IMBALANCE_THRESHOLD = 0.30  # mass center > 30% off slide center => visual imbalance


def compute_balance(slide: dict) -> dict:
    """Return {mass_x, mass_y, slide_cx, slide_cy, imbalance_x, imbalance_y, score, issue?}."""
    objects = slide.get("objects", [])
    width = slide["width_emu"]
    height = slide["height_emu"]

    weighted_x = 0.0
    weighted_y = 0.0
    total_mass = 0
    for obj in objects:
        if obj.get("anomalous") or obj.get("kind") in ("group", "connector"):
            continue
        if obj.get("width", 0) <= 0 or obj.get("height", 0) <= 0:
            continue
        cx = obj["left"] + obj["width"] / 2
        cy = obj["top"] + obj["height"] / 2
        mass = obj["width"] * obj["height"]
        weighted_x += cx * mass
        weighted_y += cy * mass
        total_mass += mass

    if total_mass == 0:
        return {
            "score": 50.0,
            "mass_x": 0, "mass_y": 0,
            "imbalance_x": 0.0, "imbalance_y": 0.0,
            "issue": None,
        }

    mass_x = weighted_x / total_mass
    mass_y = weighted_y / total_mass
    slide_cx = width / 2
    slide_cy = height / 2

    imbalance_x = abs(mass_x - slide_cx) / slide_cx if slide_cx else 0.0
    imbalance_y = abs(mass_y - slide_cy) / slide_cy if slide_cy else 0.0
    imbalance = max(imbalance_x, imbalance_y)

    # 0 imbalance -> 100, 0.5 imbalance -> 50, 1.0 -> 0.
    score = max(0.0, 100.0 - 100.0 * imbalance)

    issue = None
    if imbalance > IMBALANCE_THRESHOLD:
        direction = []
        if imbalance_x > IMBALANCE_THRESHOLD:
            direction.append("right" if mass_x > slide_cx else "left")
        if imbalance_y > IMBALANCE_THRESHOLD:
            direction.append("bottom" if mass_y > slide_cy else "top")
        issue = {
            "category": "visual-imbalance",
            "severity": "warning" if imbalance < 0.5 else "error",
            "shape_id": 0,
            "message": (
                f"Slide content center is {imbalance:.0%} off slide center "
                f"({'/'.join(direction)}); content is visually unbalanced"
            ),
            "suggested_fix": "manual-review",
            "imbalance": round(imbalance, 3),
            "mass_center": [int(mass_x), int(mass_y)],
        }

    return {
        "score": round(score, 2),
        "mass_x": int(mass_x),
        "mass_y": int(mass_y),
        "slide_cx": int(slide_cx),
        "slide_cy": int(slide_cy),
        "imbalance_x": round(imbalance_x, 3),
        "imbalance_y": round(imbalance_y, 3),
        "issue": issue,
    }
