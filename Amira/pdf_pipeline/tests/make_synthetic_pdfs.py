"""Generate SYNTHETIC test PDFs (all names, companies and wells are made up).

    .venv/bin/python tests/make_synthetic_pdfs.py [output_folder]

Default output folder: tmp/synthetic_pdfs (inside the workspace).
"""

from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

WORKSPACE = Path(__file__).resolve().parent.parent

FONT_CANDIDATES = [
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/Library/Fonts/Arial.ttf",
    "/System/Library/Fonts/Helvetica.ttc",
]


def _text_page(c: canvas.Canvas, lines: list[str], size: int = 12) -> None:
    y = A4[1] - 80
    for line in lines:
        c.setFont("Helvetica", size)
        c.drawString(60, y, line)
        y -= size * 1.8
    c.showPage()


def _scanned_page(c: canvas.Canvas, lines: list[str]) -> None:
    """A page that is only an image of text (no text layer), like a scanned report."""
    img = Image.new("L", (1654, 2339), 255)  # A4 at 200 dpi
    draw = ImageDraw.Draw(img)
    font = None
    for path in FONT_CANDIDATES:
        if Path(path).exists():
            font = ImageFont.truetype(path, 44)
            break
    font = font or ImageFont.load_default(size=44)
    y = 200
    for line in lines:
        draw.text((140, y), line, fill=0, font=font)
        y += 90
    c.drawImage(ImageReader(img), 0, 0, width=A4[0], height=A4[1])
    c.showPage()


def make_searchable(path: Path) -> None:
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setAuthor("Synthetic Author")
    c.setTitle("Embedded Metadata Title")
    _text_page(c, [
        "FINAL SEISMIC PROCESSING REPORT",
        "ZORVEX 3D SURVEY",
        "",
        "Prepared for: Acme Exploration Sdn Bhd",
        "Prepared by: Fictional Geodata Ltd",
        "Date: 14 March 1997",
    ], size=16)
    _text_page(c, [
        "1. Introduction",
        "The Zorvex Field lies offshore in a made-up basin.",
        "Wells Zorvex-1 and ZORVEX-1 were drilled first, followed by Zorvex A-12.",
        "See Figure 3 and Page 12. This is not a magnetic field study.",
    ])
    _text_page(c, [
        "Well: ZX-3ST1 Spud date 1998",
        "Reference is also made to Zorvex B-5.",
    ])
    c.save()


def make_scanned(path: Path) -> None:
    c = canvas.Canvas(str(path), pagesize=A4)
    _scanned_page(c, [
        "SEISMIC INTERPRETATION REPORT",
        "",
        "Prepared for: Acme Exploration Sdn Bhd",
        "Date: 02 June 2001",
    ])
    _text_page(c, ["Appendix A: log data for well Zorvex-7."])
    c.save()


def make_no_fields(path: Path) -> None:
    c = canvas.Canvas(str(path), pagesize=A4)
    _text_page(c, ["Meeting notes, nothing structured on this page."])
    c.save()


def main(out_dir: Path) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    files = {
        "searchable_report.pdf": make_searchable,
        "scanned_report.pdf": make_scanned,
        "no_fields.pdf": make_no_fields,
    }
    for name, maker in files.items():
        maker(out_dir / name)
    (out_dir / "corrupt.pdf").write_bytes(b"this is not a real pdf file")
    # macOS/NAS resource-fork companion; the pipeline must skip it.
    (out_dir / "._searchable_report.pdf").write_bytes(b"\x00\x05\x16\x07")
    return sorted(out_dir.iterdir())


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else WORKSPACE / "tmp" / "synthetic_pdfs"
    for f in main(target):
        print(f)
