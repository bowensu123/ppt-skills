"""Visual diff between two slide renders.

Produces:
  * pixel-diff ratio (fraction of changed pixels above per-channel threshold)
  * mean absolute error
  * a simplified SSIM-like score (luminance-only block comparison)
  * a heatmap PNG visualizing changed regions

The simplified SSIM avoids depending on scikit-image so the skill stays
deployable inside locked-down internal environments with only Pillow.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import numpy as np
from PIL import Image, ImageChops


def _load_grayscale(path: Path) -> np.ndarray:
    img = Image.open(path).convert("L")
    return np.asarray(img, dtype=np.float32)


def _load_rgb(path: Path) -> np.ndarray:
    img = Image.open(path).convert("RGB")
    return np.asarray(img, dtype=np.float32)


def _resize_to_match(a: Image.Image, b: Image.Image) -> tuple[Image.Image, Image.Image]:
    if a.size == b.size:
        return a, b
    target = a.size if a.size[0] * a.size[1] >= b.size[0] * b.size[1] else b.size
    return a.resize(target), b.resize(target)


def pixel_diff_ratio(before: Path, after: Path, threshold: int = 12) -> float:
    """Fraction of pixels whose per-channel diff exceeds ``threshold``."""
    a = Image.open(before).convert("RGB")
    b = Image.open(after).convert("RGB")
    a, b = _resize_to_match(a, b)
    diff = np.abs(np.asarray(a, dtype=np.int16) - np.asarray(b, dtype=np.int16))
    changed = np.any(diff > threshold, axis=2)
    return float(changed.mean())


def mean_abs_error(before: Path, after: Path) -> float:
    a = _load_rgb(before)
    b = _load_rgb(after)
    if a.shape != b.shape:
        bi = Image.open(after).convert("RGB").resize((a.shape[1], a.shape[0]))
        b = np.asarray(bi, dtype=np.float32)
    return float(np.mean(np.abs(a - b)))


def block_ssim(before: Path, after: Path, block: int = 8) -> float:
    """Simplified luminance-only SSIM. Returns score in [-1, 1]."""
    a = _load_grayscale(before)
    b = _load_grayscale(after)
    if a.shape != b.shape:
        bi = Image.open(after).convert("L").resize((a.shape[1], a.shape[0]))
        b = np.asarray(bi, dtype=np.float32)

    h = a.shape[0] - (a.shape[0] % block)
    w = a.shape[1] - (a.shape[1] % block)
    a = a[:h, :w].reshape(h // block, block, w // block, block).swapaxes(1, 2).reshape(-1, block * block)
    b = b[:h, :w].reshape(h // block, block, w // block, block).swapaxes(1, 2).reshape(-1, block * block)

    mu_a = a.mean(axis=1)
    mu_b = b.mean(axis=1)
    var_a = a.var(axis=1)
    var_b = b.var(axis=1)
    cov_ab = ((a - mu_a[:, None]) * (b - mu_b[:, None])).mean(axis=1)

    c1 = (0.01 * 255) ** 2
    c2 = (0.03 * 255) ** 2
    num = (2 * mu_a * mu_b + c1) * (2 * cov_ab + c2)
    den = (mu_a ** 2 + mu_b ** 2 + c1) * (var_a + var_b + c2)
    ssim_blocks = num / den
    return float(np.mean(ssim_blocks))


def heatmap(before: Path, after: Path, output: Path, threshold: int = 12) -> Path:
    a = Image.open(before).convert("RGB")
    b = Image.open(after).convert("RGB")
    a, b = _resize_to_match(a, b)
    diff = np.abs(np.asarray(a, dtype=np.int16) - np.asarray(b, dtype=np.int16)).max(axis=2)
    mask = (diff > threshold).astype(np.uint8) * 255

    base = np.asarray(a.convert("RGB"), dtype=np.uint8).copy()
    overlay = np.zeros_like(base)
    overlay[..., 0] = mask
    overlay[..., 2] = mask
    blended = (0.6 * base + 0.4 * overlay).astype(np.uint8)
    Image.fromarray(blended).save(output)
    return output


def diff(before: Path, after: Path, heatmap_path: Path | None = None) -> dict:
    pdr = pixel_diff_ratio(before, after)
    mae = mean_abs_error(before, after)
    ssim = block_ssim(before, after)
    payload = {
        "before": str(before),
        "after": str(after),
        "pixel_diff_ratio": round(pdr, 6),
        "mean_abs_error": round(mae, 4),
        "block_ssim": round(ssim, 6),
    }
    if heatmap_path is not None:
        payload["heatmap"] = str(heatmap(before, after, heatmap_path))
    return payload


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--before", required=True)
    parser.add_argument("--after", required=True)
    parser.add_argument("--output", required=True, help="JSON output path")
    parser.add_argument("--heatmap", help="optional heatmap PNG path")
    args = parser.parse_args()

    heatmap_path = Path(args.heatmap) if args.heatmap else None
    payload = diff(Path(args.before), Path(args.after), heatmap_path)
    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
