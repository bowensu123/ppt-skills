"""Extract structured content from a deck for template-based regeneration.

The output is a content JSON that decouples WHAT the deck says from HOW it
looks. A template renderer can then drop this content into an entirely
new layout (Mode 5: regenerate).

Output shape:
  {
    "title": "...",
    "subtitle": "...",
    "badge": "...",                 # corner pill/tag
    "items": [                      # peer-card content (Mode-5 templates iterate this)
      {
        "name": "Zero-shot",        # card title
        "description": "...",        # card body / paragraph
        "icon": "💬",                # leading character glyph if any (emoji)
        "details": ["优点: ...", "局限: ..."],  # short labels under the body
        "extra": "0 个示例",         # optional secondary header (e.g., "0 examples")
        "image": null               # OPTIONAL: {"path": "assets/sid_42.png",
                                    #  "asset_id": "a01"} — set by AGENT after
                                    # reading assets-manifest.json + annotated
                                    # render. The deterministic extractor leaves
                                    # this null; semantic image→item attribution
                                    # is left to the multimodal model.
      },
      ...
    ],
    "footer": "...",                 # bottom band line
    "footer_secondary": "...",      # second bottom-band line (smaller)
    "decorations": []               # OPTIONAL: list of slide-level images that
                                    # are NOT item icons (logos, accent bars,
                                    # background graphics). Each entry:
                                    # {"asset_id", "path", "left", "top",
                                    #  "width", "height"}. Agent populates from
                                    # assets-manifest.json after attribution.
  }

Detection mirrors detect_roles.py heuristics (title = largest font in upper
30%, items = peer-card containers in the same row band) and adds light
nesting heuristics inside each card.

If the deck doesn't fit the patterns (no clear card row) the items list
will be empty — caller can either pick a non-loop template or fall back to
polish.
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import tempfile
import unicodedata
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent


def _run(script: str, *args: str) -> None:
    subprocess.run([sys.executable, str(SCRIPT_DIR / script), *args], check=True)


def _is_emoji_or_symbol(text: str) -> bool:
    """Heuristic: short text whose first non-space char is a symbol / emoji."""
    s = text.strip()
    if not s or len(s) > 4:
        return False
    ch = s[0]
    cat = unicodedata.category(ch)
    return cat.startswith("S") or ord(ch) > 0x2600  # symbols + emoji range


def _is_numeric_badge(text: str) -> bool:
    """Single number / lettered badge like '1', '01', 'A', 'III'."""
    s = text.strip()
    if not s or len(s) > 3:
        return False
    return s.isdigit() or (len(s) <= 2 and s.isalnum())


def _is_namelike(text: str) -> bool:
    """Card-title-likely text: not emoji, not a numeric badge, has letters."""
    if _is_emoji_or_symbol(text):
        return False
    if _is_numeric_badge(text):
        return False
    s = text.strip()
    if not s:
        return False
    # Must contain at least one letter (latin or CJK).
    return bool(re.search(r"[A-Za-z一-鿿]", s))


def _bbox_contains(outer: dict, inner: dict, slack: int = 91440) -> bool:
    return (
        inner["left"] >= outer["left"] - slack
        and inner["top"] >= outer["top"] - slack
        and inner["left"] + inner["width"] <= outer["left"] + outer["width"] + slack
        and inner["top"] + inner["height"] <= outer["top"] + outer["height"] + slack
    )


def _identify_card_row(slide_objects: list[dict]) -> list[dict]:
    """Find a row of peer card containers (same as _card_repair logic)."""
    candidates = [
        o for o in slide_objects
        if o["kind"] == "container"
        and not o.get("anomalous")
        and o["width"] >= 800000 and o["height"] >= 1000000
    ]
    if len(candidates) < 2:
        return []
    rows: list[list[dict]] = []
    used = set()
    for i, base in enumerate(candidates):
        if i in used:
            continue
        peers = [base]; used.add(i)
        for j, other in enumerate(candidates):
            if j in used:
                continue
            if abs(other["top"] - base["top"]) > 600000:
                continue
            if abs(other["height"] - base["height"]) / max(base["height"], 1) > 1.0:
                continue
            if abs(other["width"] - base["width"]) / max(base["width"], 1) > 1.0:
                continue
            peers.append(other); used.add(j)
        if len(peers) >= 3:
            rows.append(peers)
    if not rows:
        return []
    rows.sort(key=lambda r: -len(r))
    return sorted(rows[0], key=lambda c: c["left"])


def _find_role(roles_data: dict, role_name: str) -> int | None:
    for slide in roles_data.get("slides", []):
        for entry in slide.get("shapes", []):
            if entry.get("role") == role_name:
                return entry["shape_id"]
    return None


def _shape_text(objects: list[dict], shape_id: int | None) -> str:
    if shape_id is None:
        return ""
    for obj in objects:
        if obj["shape_id"] == shape_id:
            return obj.get("text", "")
    return ""


def _runs_of(shape_obj: dict) -> list[dict]:
    """Surface the per-run formatting array from an inspect_ppt entry.
    Empty list when the shape has no runs (e.g. picture)."""
    return list(shape_obj.get("text_runs") or [])


def _extract_items(slide_objects: list[dict], cards: list[dict]) -> list[dict]:
    items = []
    for card in cards:
        children = [
            o for o in slide_objects
            if o["shape_id"] != card["shape_id"]
            and not o.get("anomalous")
            and _bbox_contains(card, o)
        ]
        text_children = [c for c in children if c.get("kind") == "text" and c.get("text")]
        text_children.sort(key=lambda c: (c["top"], c["left"]))

        item: dict = {
            "name": "", "description": "", "icon": "",
            "details": [], "extra": "",
            # NEW: per-run formatting preserved for faithful regeneration.
            "name_runs": [], "description_runs": [],
            # NEW: source shape ids so the agent / mutate ops can correlate.
            "source_shape_ids": {},
        }

        # Heuristic ordering: icon (emoji/symbol), name (largest font), extra (numeric/short),
        # description (longest), details (short trailing labels containing colons)
        if not text_children:
            items.append(item)
            continue

        # Find icon: very short, symbol-only.
        for c in text_children:
            if _is_emoji_or_symbol(c["text"]):
                item["icon"] = c["text"].strip()
                break

        # Find name: name-like text (has letters, not numeric) in upper half of card.
        name_candidates = [c for c in text_children if _is_namelike(c["text"])]
        if name_candidates:
            top_half = card["top"] + card["height"] * 0.40
            tops = [c for c in name_candidates if c["top"] < top_half]
            # Prefer larger font sizes among upper-half name-like shapes.
            tops.sort(key=lambda c: -(max(c["font_sizes"]) if c["font_sizes"] else 0))
            if tops:
                item["name"] = tops[0]["text"].strip()
                item["name_runs"] = _runs_of(tops[0])
                item["source_shape_ids"]["name"] = tops[0]["shape_id"]

        # Description: longest paragraph-like text in mid band.
        body_candidates = [
            c for c in text_children
            if _is_namelike(c["text"])
            and c["text"].strip() != item["name"]
            and len(c["text"]) > 10
        ]
        if body_candidates:
            body_candidates.sort(key=lambda c: -len(c["text"]))
            item["description"] = body_candidates[0]["text"].strip()
            item["description_runs"] = _runs_of(body_candidates[0])
            item["source_shape_ids"]["description"] = body_candidates[0]["shape_id"]

        # Extra: short label like "0 个示例", "1 个示例" near the icon.
        mid_top = card["top"] + card["height"] * 0.55
        for c in text_children:
            txt = c["text"].strip()
            if txt in (item["name"], item["description"]):
                continue
            if _is_emoji_or_symbol(txt):
                continue
            if c["top"] >= mid_top:
                continue
            if 1 <= len(txt) <= 10 and not _is_numeric_badge(txt):
                if not item["extra"]:
                    item["extra"] = txt
                    break

        # Trailing details: label-like short pieces in the lower half.
        for c in text_children:
            txt = c["text"].strip()
            if txt in (item["name"], item["description"], item["extra"]):
                continue
            if _is_emoji_or_symbol(txt) or _is_numeric_badge(txt):
                continue
            if c["top"] < card["top"] + card["height"] * 0.55:
                continue
            if 1 <= len(txt) <= 16:
                item["details"].append(txt)

        # Dedupe details preserving order.
        seen = set()
        item["details"] = [t for t in item["details"] if not (t in seen or seen.add(t))][:6]

        items.append(item)
    return items


def extract_content(input_path: Path, work_dir: Path) -> dict:
    work_dir.mkdir(parents=True, exist_ok=True)
    inspection = work_dir / "inspection.json"
    roles = work_dir / "roles.json"
    _run("inspect_ppt.py", "--input", str(input_path), "--output", str(inspection))
    _run("detect_roles.py", "--inspection", str(inspection), "--output", str(roles))

    insp_data = json.loads(inspection.read_text(encoding="utf-8"))
    roles_data = json.loads(roles.read_text(encoding="utf-8"))
    slide = insp_data["slides"][0]
    objects = slide["objects"]

    title_id = _find_role(roles_data, "title")
    subtitle_id = _find_role(roles_data, "subtitle")
    badge_id = _find_role(roles_data, "badge")

    # detect_roles can pick a stray emoji shape as title when it's the largest
    # font in upper 30%. Re-validate: if the role-tagged title isn't name-like,
    # search the upper-band of the slide for the actual title text.
    title_text = _shape_text(objects, title_id)
    if not _is_namelike(title_text):
        height = slide["height_emu"]
        candidates = [
            o for o in objects
            if o.get("kind") == "text"
            and _is_namelike(o.get("text", ""))
            and o["top"] < height * 0.30
            and not o.get("anomalous")
        ]
        if candidates:
            candidates.sort(key=lambda c: -(max(c["font_sizes"]) if c["font_sizes"] else 0))
            title_id = candidates[0]["shape_id"]

    cards = _identify_card_row(objects)
    items = _extract_items(objects, cards)

    # Footer: text shapes near bottom 25% of slide, large width.
    height = slide["height_emu"]
    width = slide["width_emu"]
    footer_band = [
        o for o in objects
        if o.get("kind") == "text"
        and o.get("text")
        and o["top"] > height * 0.75
        and o["width"] > width * 0.30
        and o["shape_id"] not in {title_id, subtitle_id, badge_id}
    ]
    footer_band.sort(key=lambda o: o["top"])
    footer = footer_band[0]["text"] if footer_band else ""
    footer_secondary = footer_band[1]["text"] if len(footer_band) > 1 else ""

    content = {
        "title": _shape_text(objects, title_id).strip(),
        "subtitle": _shape_text(objects, subtitle_id).strip(),
        "badge": _shape_text(objects, badge_id).strip(),
        "items": items,
        "footer": footer.strip(),
        "footer_secondary": footer_secondary.strip(),
        "_meta": {
            "input": str(input_path),
            "card_count": len(cards),
            "slide_w": width,
            "slide_h": height,
        },
    }

    out = work_dir / "content.json"
    out.write_text(json.dumps(content, ensure_ascii=False, indent=2), encoding="utf-8")
    return content


def main() -> int:
    # Force UTF-8 stdout so emoji/CJK in titles don't crash on GBK consoles.
    import io
    if hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", line_buffering=True)
    parser = argparse.ArgumentParser()
    parser.add_argument("--in", dest="in_path", required=True, type=Path)
    parser.add_argument("--work-dir", required=True, type=Path)
    args = parser.parse_args()
    content = extract_content(args.in_path, args.work_dir)
    print(json.dumps({
        "title": content["title"][:60],
        "subtitle_len": len(content["subtitle"]),
        "items_count": len(content["items"]),
        "footer_len": len(content["footer"]),
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
