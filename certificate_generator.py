from __future__ import annotations

import argparse
import csv
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

from PIL import Image, ImageColor, ImageDraw, ImageFont


@dataclass(frozen=True)
class TextSpec:
    field: str
    x: int
    y: int
    size: int
    color: str
    anchor: str  # Pillow anchor, e.g. "mm" (center), "lm" (left-middle)
    font_path: Optional[str] = None


def _parse_color(s: str) -> tuple[int, int, int]:
    c = ImageColor.getrgb(s)
    if len(c) == 4:
        return (c[0], c[1], c[2])
    return c


def _load_font(font_path: Optional[str], size: int) -> ImageFont.FreeTypeFont:
    if font_path:
        return ImageFont.truetype(font_path, size=size)

    # Best-effort defaults: try common fonts before falling back to bitmap.
    candidates = [
        "DejaVuSans.ttf",
        # macOS (Supplemental fonts)
        "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode MS.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ]
    for cand in candidates:
        try:
            return ImageFont.truetype(cand, size=size)
        except Exception:
            continue

    # Fall back to a tiny default bitmap font (size may not scale).
    return ImageFont.load_default()


def _read_rows(csv_path: Path) -> list[dict[str, str]]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError("CSV header row is missing.")
        return [{k: (v or "").strip() for k, v in row.items()} for row in reader]


def _render_one(
    template: Image.Image,
    row: dict[str, str],
    specs: Iterable[TextSpec],
    font_path: Optional[str] = None,
) -> Image.Image:
    img = template.copy()
    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGBA")

    draw = ImageDraw.Draw(img)
    for spec in specs:
        text = row.get(spec.field, "")
        font = _load_font(spec.font_path or font_path, spec.size)
        draw.text(
            (spec.x, spec.y),
            text,
            fill=_parse_color(spec.color),
            font=font,
            anchor=spec.anchor,
        )
    return img


def _safe_filename(s: str) -> str:
    s = s.strip().replace(os.sep, "_")
    return "".join(ch if ch.isalnum() or ch in ("_", "-", " ", ".", "(", ")", "[", "]") else "_" for ch in s).strip()


def _unique_filename_base(base: str, used: dict[str, int], fallback: str) -> str:
    normalized = _safe_filename(base) or fallback
    count = used.get(normalized, 0) + 1
    used[normalized] = count
    if count == 1:
        return normalized
    return f"{normalized}_{count}"


def main() -> int:
    p = argparse.ArgumentParser(
        description="Generate certificates/awards from an image template + CSV."
    )
    p.add_argument("--template", required=True, help="Template image path (PNG/JPG).")
    p.add_argument("--csv", required=True, help="Recipients CSV path (UTF-8).")
    p.add_argument("--out", default="out", help="Output directory (default: out).")
    p.add_argument(
        "--font",
        default=None,
        help="TTF/OTF font path. Recommended for Korean text (e.g. AppleSDGothicNeo*.ttf).",
    )
    p.add_argument(
        "--png",
        action="store_true",
        help="Save individual PNGs (default: on unless --pdf-only).",
    )
    p.add_argument(
        "--pdf",
        action="store_true",
        help="Also save a combined PDF.",
    )
    p.add_argument(
        "--pdf-only",
        action="store_true",
        help="Do not save individual PNGs; only output a PDF (implies --pdf).",
    )
    p.add_argument(
        "--filename-field",
        default="name",
        help="CSV field used in output filenames (default: name).",
    )

    # Text placements (repeatable).
    # Example: --text name 800 620 72 "#111111" mm
    p.add_argument(
        "--text",
        nargs=6,
        action="append",
        metavar=("FIELD", "X", "Y", "SIZE", "COLOR", "ANCHOR"),
        help='Draw CSV FIELD at (X,Y). COLOR like "#111111". ANCHOR like mm/lm/rm.',
    )

    args = p.parse_args()
    if args.pdf_only:
        args.pdf = True
        args.png = False
    if not args.png and not args.pdf:
        args.png = True

    template_path = Path(args.template).expanduser()
    csv_path = Path(args.csv).expanduser()
    out_dir = Path(args.out).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not args.text:
        # Sensible default placements; adjust to your template.
        # Coordinates are pixels from top-left of the template image.
        args.text = [
            ["award_title", "800", "430", "80", "#111111", "mm"],
            ["name", "800", "600", "110", "#111111", "mm"],
            ["org", "800", "780", "44", "#333333", "mm"],
            ["date", "800", "850", "36", "#333333", "mm"],
        ]

    specs: list[TextSpec] = []
    for field, x, y, size, color, anchor in args.text:
        specs.append(
            TextSpec(
                field=field,
                x=int(x),
                y=int(y),
                size=int(size),
                color=color,
                anchor=anchor,
            )
        )

    rows = _read_rows(csv_path)
    if not rows:
        raise ValueError("CSV has no data rows.")

    template_img = Image.open(template_path)

    rendered_images_rgb: list[Image.Image] = []
    used_filenames: dict[str, int] = {}
    for i, row in enumerate(rows, start=1):
        img = _render_one(template_img, row, specs, args.font)

        name_for_file = row.get(args.filename_field, f"row_{i}")
        base = _unique_filename_base(name_for_file, used_filenames, f"row_{i}")

        if args.png:
            out_png = out_dir / f"{base}.png"
            img.save(out_png)

        if args.pdf:
            rendered_images_rgb.append(img.convert("RGB"))

    if args.pdf:
        out_pdf = out_dir / "certificates.pdf"
        first, rest = rendered_images_rgb[0], rendered_images_rgb[1:]
        first.save(out_pdf, save_all=True, append_images=rest)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
