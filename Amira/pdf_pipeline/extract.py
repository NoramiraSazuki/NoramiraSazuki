"""Read page text and embedded metadata from a PDF with pypdf (read-only)."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import datetime

from pypdf import PdfReader

log = logging.getLogger(__name__)


class PdfReadError(RuntimeError):
    pass


@dataclass
class PageText:
    number: int  # 1-based physical page number
    text: str
    source: str  # "text" (embedded text layer) or "ocr"


@dataclass
class PdfMetadata:
    title: str = ""
    author: str = ""
    creator: str = ""
    producer: str = ""
    creation_date: str = ""


def open_reader(stream) -> PdfReader:
    reader = PdfReader(stream, strict=False)
    if reader.is_encrypted:
        try:
            ok = reader.decrypt("")
        except Exception as exc:  # pypdf raises various types for unsupported crypto
            raise PdfReadError(f"Encrypted PDF could not be opened: {type(exc).__name__}") from exc
        if not ok:
            raise PdfReadError("Encrypted PDF requires a password")
    return reader


def _format_date(value) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return str(value or "").strip()


def read_metadata(reader: PdfReader) -> PdfMetadata:
    try:
        md = reader.metadata
    except Exception as exc:
        log.debug("Metadata unreadable: %s", type(exc).__name__)
        return PdfMetadata()
    if not md:
        return PdfMetadata()
    try:
        created = md.creation_date
    except Exception:
        # Malformed date string: keep the raw value rather than guessing.
        created = md.get("/CreationDate", "")
    return PdfMetadata(
        title=str(md.title or "").strip(),
        author=str(md.author or "").strip(),
        creator=str(md.creator or "").strip(),
        producer=str(md.producer or "").strip(),
        creation_date=_format_date(created),
    )


def extract_page_texts(reader: PdfReader) -> tuple[list[str], list[int]]:
    """Return (text per page, list of page numbers that failed to extract)."""
    texts, failed = [], []
    for i, page in enumerate(reader.pages, start=1):
        try:
            texts.append(page.extract_text() or "")
        except Exception as exc:
            log.debug("Text extraction failed on page %d: %s", i, type(exc).__name__)
            texts.append("")
            failed.append(i)
    return texts, failed
