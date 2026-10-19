"""Open source PDFs read-only and render single pages with bounded size (pypdfium2)."""

from __future__ import annotations

import logging
from pathlib import Path

import pypdfium2 as pdfium

log = logging.getLogger("pilot")

PDFIUM_VERSION = f"pypdfium2-{pdfium.PYPDFIUM_INFO}/pdfium-{pdfium.PDFIUM_INFO}"


class PdfOpenError(RuntimeError):
    """The PDF cannot be processed (corrupt, or needs a password). Not retried automatically."""

    def __init__(self, message: str, code: str):
        super().__init__(message)
        self.code = code  # "encrypted" / "corrupt"


def encryption_of(path: Path) -> str:
    """Encryption algorithm only (e.g. 'Standard AES-128'), via the existing pypdf helper; '' if none."""
    from pypdf import PdfReader

    from pdf_pipeline.extract import encryption_type

    try:
        with open(path, "rb") as fh:
            return encryption_type(PdfReader(fh, strict=False))
    except Exception as exc:  # unreadable trailer: report, don't fail here
        log.debug("encryption check failed: %s", type(exc).__name__)
        return ""


def pdf_metadata(path: Path) -> dict:
    """Embedded metadata via the CLI's reader (reference only; never used as document content)."""
    from dataclasses import asdict

    from pdf_pipeline.extract import open_reader, read_metadata

    try:
        with open(path, "rb") as fh:
            return asdict(read_metadata(open_reader(fh)))
    except Exception as exc:
        log.debug("metadata read failed: %s", type(exc).__name__)
        return {}


def open_pdf(path: Path) -> pdfium.PdfDocument:
    """Open without a password. pdfium reads the file lazily, so huge PDFs are not loaded into memory."""
    try:
        return pdfium.PdfDocument(str(path))  # pdfium opens files read-only
    except pdfium.PdfiumError as exc:
        msg = str(exc)
        if "password" in msg.lower():
            raise PdfOpenError("Encrypted PDF requires a password (not attempted)", "encrypted") from exc
        if isinstance(exc.__cause__, FileNotFoundError) or not path.exists():
            raise FileNotFoundError("source file not found") from exc
        raise PdfOpenError(f"PDF could not be opened ({msg.split('(PDFium:')[-1].strip(' )') or 'unknown'})",
                           "corrupt") from exc


def render_scale(page: pdfium.PdfPage, dpi: int, max_megapixels: float) -> float:
    """Scale for ``dpi``, reduced so the image never exceeds ``max_megapixels`` (large foldouts)."""
    w, h = page.get_size()
    scale = dpi / 72.0
    pixels = (w * scale) * (h * scale)
    limit = max_megapixels * 1_000_000
    if pixels > limit:
        scale *= (limit / pixels) ** 0.5
    return scale


def render_page(page: pdfium.PdfPage, dpi: int, max_megapixels: float):
    """Render one page (page /Rotate applied, as a viewer shows it). Returns (PIL image, scale)."""
    scale = render_scale(page, dpi, max_megapixels)
    bitmap = page.render(scale=scale, grayscale=False)
    try:
        return bitmap.to_pil(), scale
    finally:
        bitmap.close()
