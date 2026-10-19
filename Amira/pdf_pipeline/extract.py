"""Read page text and embedded metadata from a PDF with pypdf (read-only)."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import datetime

from pypdf import PdfReader

log = logging.getLogger(__name__)


class PdfReadError(RuntimeError):
    def __init__(self, message: str, encryption: str = ""):
        super().__init__(message)
        self.encryption = encryption


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


def encryption_type(reader: PdfReader) -> str:
    """Describe the PDF's encryption scheme, e.g. 'Standard AES-128'. '' when not encrypted.

    Only the algorithm is read from the /Encrypt dictionary; passwords and keys are
    never read, derived, or returned.
    """
    try:
        enc = reader.trailer.get("/Encrypt")
        if enc is None:
            return ""
        enc = enc.get_object()
        handler = str(enc.get("/Filter", "")).lstrip("/")
        v, length = int(enc.get("/V", 0)), int(enc.get("/Length", 40))
        if v in (4, 5):
            cfm = ""
            filters = enc.get("/CF")
            if filters is not None:
                std = filters.get_object().get(str(enc.get("/StmF", "/StdCF")))
                cfm = str(std.get_object().get("/CFM", "")).lstrip("/") if std is not None else ""
            algo = {"AESV2": "AES-128", "AESV3": "AES-256", "V2": f"RC4-{length}"}.get(cfm, cfm or f"V{v}")
        elif v == 1:
            algo = "RC4-40"
        elif v == 2:
            algo = f"RC4-{length}"
        else:
            algo = f"V{v}"
        return f"{handler} {algo}".strip()
    except Exception:
        return "Encrypted (type unreadable)"


def open_reader(stream) -> PdfReader:
    """Open a PDF. Encrypted files are opened only if they need no password (owner-protected)."""
    reader = PdfReader(stream, strict=False)
    if reader.is_encrypted:
        enc = encryption_type(reader)
        try:
            ok = reader.decrypt("")
        except Exception as exc:  # pypdf raises various types for unsupported crypto
            raise PdfReadError(f"Encrypted PDF could not be opened: {type(exc).__name__}: {str(exc)[:120]}",
                               enc) from exc
        if not ok:
            raise PdfReadError("Encrypted PDF requires a password (not attempted)", enc)
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
