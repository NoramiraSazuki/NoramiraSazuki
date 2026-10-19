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


def make_completion_report(path: Path) -> None:
    c = canvas.Canvas(str(path), pagesize=A4)
    _text_page(c, [
        "WELL COMPLETION REPORT",
        "ZORVEX B-5",
        "Prepared for: Acme Exploration Sdn Bhd",
        "Prepared by: Fictional Drilling Services Sdn Bhd",
        "Report Date: 20 August 1998",
        "Transmittal Date: 25/08/1998",
    ], size=14)
    _text_page(c, [
        "Well Name: Zorvex B-5     UWI: ZX-000-B5-0001",
        "Spud Date: 03 June 1998     Completion Date: 1998-07-30",
        "Top Depth: 1,250.5 m MD",
        "Base Depth (ft TVD): 8200",
        "Sampling Interval: 0.1524 m",
        "Model: ZX Static Model v2",
    ])
    _text_page(c, [
        "Operations summary: the well was re-entered.",
        "Spud Date: 05 June 1998",
    ])
    c.save()


def make_log_header(path: Path) -> None:
    c = canvas.Canvas(str(path), pagesize=A4)
    _text_page(c, [
        "COMPOSITE LOG",
        "Company: Acme Exploration Sdn Bhd",
        "Logging Company: Fictional Wireline Ltd",
        "Log Title: Gamma Ray - Resistivity",
        "Run No: 2     Trip Date: 11-Sep-1998",
        "Logged Interval: 1200 - 2450 m MDRT",
        "Scale: 1:200 and 1:500",
        "Well: Zorvex-7",
    ])
    c.save()


ARTICLE_TITLE = "Borehole Logging Report from the Quillon Deep Drilling Project: Temperature Profiles"


def make_journal_article(path: Path) -> None:
    """Mimics a published borehole paper: title over 2 lines, publication dates, and a bibliography
    whose entries contain depths, a 'model:', a field phrase and a well name that must all be ignored."""
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setTitle(ARTICLE_TITLE)
    _text_page(c, [
        "Journal of Synthetic Geoscience 12 (1999) 101-115",
        "Borehole Logging Report from the Quillon Deep",
        "Drilling Project: Temperature Profiles",
        "A. Fictional, B. Imaginary",
        "Received 12 March 1998; accepted 2 October 1998",
        "Abstract",
        "Temperature logs were acquired in the Quillon Deep Borehole.",
        "Tests were run at the Quillon Field Laboratory and relate to Earth's Thermal Field.",
    ])
    _text_page(c, [
        "2. Logging Programme",
        "Logged Interval: 1500 - 9000 m MD",
        "Data were sampled every 0.1524 m (Fictional et al., 1995).",
        "the top depth: 77 m was not reached in this sentence.",
    ])
    _text_page(c, [
        "References",
        "Fictional, A., Imaginary, B., 1995. The Geobattery model: a synthetic view. J. Synth. Geophys. 100, 1-20.",
        "Other, C., Writer, D., 1996. Logged Interval: 200 - 400 m in the Quillon Field. Proc. Synth. 5, 7-9.",
        "Writer, D., 1997. Top Depth: 50 m at well Zorvex-9. Bull. Synth. 3, 1-2.",
        "Model: Reference Model 7",
        "Appendix A",
        "Model: Thermal Model ZT-1",
    ])
    c.save()


def make_encrypted(path: Path) -> None:
    """Owner-password protected (opens without a password, like many scanned reports)."""
    from reportlab.lib.pdfencrypt import StandardEncryption

    enc = StandardEncryption("", ownerPassword=SYNTHETIC_OWNER_PASSWORD, canPrint=1, strength=128)
    c = canvas.Canvas(str(path), pagesize=A4, encrypt=enc)
    _text_page(c, ["SEISMIC SURVEY REPORT", "Prepared for: Acme Exploration Sdn Bhd", "Date: 01 May 2003"])
    c.save()


# Made-up test value; tests assert it never appears in any output file.
SYNTHETIC_OWNER_PASSWORD = "SyntheticOwnerSecret123"


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
        "completion_report.pdf": make_completion_report,
        "log_header.pdf": make_log_header,
        "encrypted_report.pdf": make_encrypted,
        "journal_article.pdf": make_journal_article,
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
