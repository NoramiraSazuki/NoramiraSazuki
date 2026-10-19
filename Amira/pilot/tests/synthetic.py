"""SYNTHETIC test documents and a fake inventory workbook. All names, values and wells are made up.

    ../.venv/bin/python -m pilot.tests.synthetic <output_folder>     (writes sample PDFs + inventory)
"""

from __future__ import annotations

import random
import sys
import unicodedata
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A3, A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

FONT_CANDIDATES = ["/System/Library/Fonts/Supplemental/Arial.ttf", "/Library/Fonts/Arial.ttf",
                   "/System/Library/Fonts/Helvetica.ttc"]

# Ground truth for the scanned table on mixed.pdf page 3 (signs, decimals and units matter).
TABLE_ROWS = [
    ["Line", "Offset (m)", "Static (ms)", "Velocity (m/s)"],
    ["ZX-101", "-12.5", "+3.25", "1,875.0"],
    ["ZX-102", "0.0", "-0.75", "1,902.4"],
    ["ZX-103", "25.75", "-11.40", "2,010.9"],
]
MIXED_EXPECTED = {
    1: "textlayer",   # digital cover
    2: "tesseract",   # image-only scan
    3: "tesseract",   # digital header over a scanned table (a few real characters - still OCR)
    4: "textlayer",   # two-column digital text
    5: "tesseract",   # landscape scan stored with /Rotate
    6: "textlayer",   # A3 foldout with digital text
    7: "tesseract",   # figure (synthetic "seismic" wiggles) with a short caption
    8: "tesseract",   # unmapped glyphs (cid) - garbled text layer
}


def _font(size: int):
    for p in FONT_CANDIDATES:
        if Path(p).exists():
            return ImageFont.truetype(p, size)
    return ImageFont.load_default(size=size)


def _text_image(lines: list[str], size=(1654, 2339), font_px=44, top=200, step=90) -> Image.Image:
    img = Image.new("L", size, 255)
    d = ImageDraw.Draw(img)
    f = _font(font_px)
    y = top
    for line in lines:
        d.text((140, y), line, fill=0, font=f)
        y += step
    return img


def _table_image(rows: list[list[str]], size=(1654, 1500)) -> Image.Image:
    img = Image.new("L", size, 255)
    d = ImageDraw.Draw(img)
    f = _font(40)
    x0, y0, cw, rh = 100, 120, 360, 110
    for r, row in enumerate(rows):
        for c, cell in enumerate(row):
            d.rectangle([x0 + c * cw, y0 + r * rh, x0 + (c + 1) * cw, y0 + (r + 1) * rh], outline=0, width=3)
            d.text((x0 + c * cw + 18, y0 + r * rh + 32), cell, fill=0, font=f)
    return img


def _wiggles(size=(1654, 1400), seed=7) -> Image.Image:
    """Abstract wiggle traces - looks like a seismic section but encodes nothing."""
    rnd = random.Random(seed)
    img = Image.new("L", size, 255)
    d = ImageDraw.Draw(img)
    for x in range(40, size[0] - 40, 14):
        y_prev, x_prev = 0, x
        for y in range(0, size[1], 6):
            dx = x + int(6 * rnd.uniform(-1, 1))
            d.line([(x_prev, y_prev), (dx, y)], fill=0, width=1)
            x_prev, y_prev = dx, y
    return img


def _page_marker(c: canvas.Canvas, n: int) -> None:
    c.setFont("Helvetica", 8)
    c.drawString(40, 20, f"PAGE-MARKER {n}")


def make_mixed(path: Path) -> Path:
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setTitle("Synthetic Mixed Report")
    # 1 digital cover
    c.setFont("Helvetica-Bold", 18)
    for i, line in enumerate(["FINAL SEISMIC PROCESSING REPORT", "ZORVEX 3D SURVEY"]):
        c.drawString(60, 760 - i * 28, line)
    c.setFont("Helvetica", 12)
    for i, line in enumerate(["Prepared for: Acme Exploration Sdn Bhd", "Prepared by: Fictional Geodata Ltd",
                              "Date: 14 March 1997"]):
        c.drawString(60, 680 - i * 20, line)
    _page_marker(c, 1)
    c.showPage()
    # 2 image-only scan
    c.drawImage(ImageReader(_text_image(["SCANNED PAGE TWO", "Acquisition parameters were reviewed.",
                                         "Record length 6.0 s at 2 ms sample rate."])), 0, 0, *A4)
    c.showPage()
    # 3 digital header over a scanned table
    c.drawImage(ImageReader(_table_image(TABLE_ROWS)), 0, 0, *A4)
    c.setFont("Helvetica", 9)
    c.drawString(40, 815, "Fictional Geodata Ltd - Table 4")
    c.showPage()
    # 4 two-column digital
    c.setFont("Helvetica", 10)
    left = [f"Left column line {i}: velocity analysis note." for i in range(1, 25)]
    right = [f"Right column line {i}: static correction note." for i in range(1, 25)]
    for i, (a, b) in enumerate(zip(left, right)):
        c.drawString(50, 780 - i * 16, a)
        c.drawString(310, 780 - i * 16, b)
    _page_marker(c, 4)
    c.showPage()
    # 5 landscape scan stored sideways with /Rotate 90
    land = _text_image(["ROTATED SCAN PAGE FIVE", "Shot point interval 25 m."], size=(2339, 1654), top=150)
    sideways = land.rotate(90, expand=True)  # stored portrait; the viewer turns it with /Rotate
    c.drawImage(ImageReader(sideways), 0, 0, *A4)
    c.setPageRotation(90)
    c.showPage()
    c.setPageRotation(0)  # reportlab keeps the rotation for later pages otherwise
    # 6 A3 foldout with digital table text
    c.setPageSize(landscape(A3))
    c.setFont("Helvetica", 11)
    for i, row in enumerate(TABLE_ROWS):
        for j, cell in enumerate(row):
            c.drawString(60 + j * 220, 760 - i * 22, cell)
    _page_marker(c, 6)
    c.showPage()
    # 7 figure page: wiggle image + short caption
    c.setPageSize(A4)
    c.drawImage(ImageReader(_wiggles()), 40, 200, 515, 600)
    c.setFont("Helvetica", 9)
    c.drawString(60, 170, "Figure 7. Synthetic section.")
    c.showPage()
    # 8 garbled text layer
    c.setFont("Helvetica", 10)
    for i in range(10):
        c.drawString(60, 760 - i * 16, "(cid:12)(cid:45)(cid:3)(cid:88)(cid:17)(cid:2)" * 3)
    c.showPage()
    c.save()
    return path


def make_long_digital(path: Path, pages: int = 4579) -> Path:
    c = canvas.Canvas(str(path), pagesize=A4)
    for n in range(1, pages + 1):
        c.setFont("Helvetica", 11)
        c.drawString(60, 780, f"Synthetic long report - page {n} of {pages}")
        c.drawString(60, 760, f"Reading {n * 0.25:.2f} m, value -{n % 97}.5 ms")
        _page_marker(c, n)
        c.showPage()
    c.save()
    return path


def make_scanned(path: Path, pages: int = 12) -> Path:
    c = canvas.Canvas(str(path), pagesize=A4)
    for n in range(1, pages + 1):
        img = _text_image([f"SCANNED SHEET {n}", "Synthetic observer log entry."], size=(827, 1170), font_px=40,
                          top=120, step=70)
        c.drawImage(ImageReader(img), 0, 0, *A4)
        c.showPage()
    c.save()
    return path


def make_text(path: Path, pages: int = 5, label: str = "Synthetic") -> Path:
    c = canvas.Canvas(str(path), pagesize=A4)
    for n in range(1, pages + 1):
        c.setFont("Helvetica", 12)
        c.drawString(60, 780, f"{label} page {n}")
        _page_marker(c, n)
        c.showPage()
    c.save()
    return path


def make_encrypted(path: Path) -> Path:
    """User-password protected: cannot be opened without a password (the app never tries)."""
    from reportlab.lib.pdfencrypt import StandardEncryption

    enc = StandardEncryption("synthetic-user-pw", ownerPassword="synthetic-owner-pw", strength=128)
    c = canvas.Canvas(str(path), pagesize=A4, encrypt=enc)
    c.drawString(60, 780, "Locked synthetic page")
    c.showPage()
    c.save()
    return path


def make_inventory(path: Path, rows: list[dict], title_rows: int = 2) -> Path:
    """Fake inventory with a title block above the header and columns in a different order."""
    from openpyxl import Workbook

    wb = Workbook()
    wb.active.title = "Summary"
    ws = wb.create_sheet("File Inventory")
    for i in range(title_rows):
        ws.append([f"Synthetic inventory title line {i + 1}"])
    header = ["File Name", "#", "Sub-folder", "File Path", "Total Pages", "Size (bytes)",
              "*Status of Documents/ Data Ingested", "Duplicate Group", "Copies (same ETag)", "ETag"]
    ws.append(header)
    for r in rows:
        ws.append([r.get(h) for h in header])
    wb.save(path)
    return path


def nfd(name: str) -> str:
    return unicodedata.normalize("NFD", name)


if __name__ == "__main__":
    out = Path(sys.argv[1])
    out.mkdir(parents=True, exist_ok=True)
    make_mixed(out / "mixed.pdf")
    make_scanned(out / "scanned.pdf")
    print("written to", out)
