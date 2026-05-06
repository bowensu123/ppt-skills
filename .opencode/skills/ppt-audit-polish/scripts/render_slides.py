from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
from pathlib import Path

import fitz


def _find_soffice() -> str | None:
    env_path = os.environ.get("SOFFICE_PATH")
    if env_path and Path(env_path).exists():
        return env_path

    path_match = shutil.which("soffice") or shutil.which("soffice.exe")
    if path_match:
        return path_match

    for candidate in (
        Path(os.environ.get("ProgramFiles", "")) / "LibreOffice" / "program" / "soffice.exe",
        Path(os.environ.get("ProgramFiles(x86)", "")) / "LibreOffice" / "program" / "soffice.exe",
    ):
        if candidate.exists():
            return str(candidate)
    return None


def render_slides(input_path: Path, output_dir: Path) -> dict:
    soffice = _find_soffice()
    output_dir.mkdir(parents=True, exist_ok=True)

    if soffice is None:
        return {
            "status": "skipped",
            "reason": "soffice-not-found",
            "images": [],
            "dependency_hint": "Install LibreOffice or set SOFFICE_PATH to soffice.exe for visual validation.",
        }

    pdf_dir = output_dir / "pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(pdf_dir),
            str(input_path),
        ],
        check=True,
        capture_output=True,
        text=True,
    )

    pdf_path = pdf_dir / f"{input_path.stem}.pdf"
    document = fitz.open(pdf_path)
    images: list[str] = []
    for page_number, page in enumerate(document, start=1):
        pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        image_path = output_dir / f"slide-{page_number:03d}.png"
        pixmap.save(image_path)
        images.append(str(image_path))

    return {"status": "rendered", "reason": "", "images": images, "pdf": str(pdf_path)}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--manifest", required=True)
    args = parser.parse_args()

    manifest = render_slides(Path(args.input), Path(args.output_dir))
    manifest_path = Path(args.manifest)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
