"""Picture-specific checks: hard crops + suspicious aspect ratios.

PowerPoint stores picture crops via ``a:srcRect`` inside ``p:blipFill``:

    <p:pic>
      <p:blipFill>
        <a:blip ...>
        <a:srcRect l="10000" t="0" r="20000" b="5000"/>
      </p:blipFill>
      ...
    </p:pic>

Each value is a per-100000 fraction of the source image to clip from that
side. Anything > 0 means part of the image is hidden. We flag crops > 10%
on any side (might be intentional but worth a heads-up) and crops > 30%
(very likely truncating content).
"""
from __future__ import annotations


CROP_INFO_RATIO = 0.10
CROP_WARN_RATIO = 0.30


def detect_image_issues(slide: dict) -> list[dict]:
    issues: list[dict] = []
    for obj in slide.get("objects", []):
        if obj.get("kind") != "picture":
            continue
        if obj.get("anomalous"):
            continue
        crop = obj.get("crop") or {}
        l = crop.get("l", 0) / 100000.0
        t = crop.get("t", 0) / 100000.0
        r = crop.get("r", 0) / 100000.0
        b = crop.get("b", 0) / 100000.0
        max_crop = max(l, t, r, b)
        if max_crop > CROP_WARN_RATIO:
            issues.append({
                "category": "picture-heavy-crop",
                "severity": "warning",
                "shape_id": obj["shape_id"],
                "message": (
                    f"{obj['name']} crops {int(max_crop * 100)}% off one edge "
                    f"(L={int(l*100)}% T={int(t*100)}% R={int(r*100)}% B={int(b*100)}%)"
                ),
                "suggested_fix": "manual-review",
                "max_crop_ratio": round(max_crop, 3),
            })
        elif max_crop > CROP_INFO_RATIO:
            issues.append({
                "category": "picture-cropped",
                "severity": "info",
                "shape_id": obj["shape_id"],
                "message": (
                    f"{obj['name']} has crop "
                    f"L={int(l*100)}% T={int(t*100)}% R={int(r*100)}% B={int(b*100)}%"
                ),
                "suggested_fix": "manual-review",
                "max_crop_ratio": round(max_crop, 3),
            })
    return issues
