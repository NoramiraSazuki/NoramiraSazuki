"""Read the file-inventory workbook (read-only) and resolve each row to a file in the source folder.

Columns are found by header name, never by position or row number. Resolution never falls back
to matching by file name: a row resolves only when its exact relative path (after Unicode NFC
normalisation) exists under the source root. Anything else is reported.
"""

from __future__ import annotations

import os
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path, PurePosixPath

from pdf_pipeline.safety import is_within

REQUIRED_COLUMNS = ("File Path", "File Name")
# Optional inventory columns copied into the catalogue when present (shown, never changed).
OPTIONAL_COLUMNS = {
    "#": "inv_row_id",
    "Total Pages": "inv_pages",
    "Size (bytes)": "inv_size",
    "*Status of Documents/ Data Ingested": "inv_ingest_status",
    "Duplicate Group": "inv_duplicate_group",
    "Copies (same ETag)": "inv_copies",
    "ETag": "inv_etag",
}

# Resolution status codes
OK = "OK"
MISSING = "MISSING"
AMBIGUOUS = "AMBIGUOUS"
OUTSIDE_SOURCE = "OUTSIDE_SOURCE"
BAD_PATH = "BAD_PATH"
PREFIX_MISMATCH = "PREFIX_MISMATCH"


class InventoryError(RuntimeError):
    pass


@dataclass
class InventoryRow:
    rel_path: str                # relative to source_root, POSIX, NFC
    file_name: str
    resolve_status: str = ""
    resolved_path: str = ""
    note: str = ""
    extra: dict = field(default_factory=dict)  # optional inventory columns (inv_*)


@dataclass
class InventoryResult:
    rows: list[InventoryRow]
    unlisted: list[str]          # files on disk under source_root that no inventory row names
    source_available: bool
    header_row: int              # 1-based sheet row that held the header (informational)


def _nfc(text: str) -> str:
    return unicodedata.normalize("NFC", text)


def _cell(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return value


def read_inventory(workbook: Path, sheet: str, filter_column: str, filter_value: str,
                   header_search_rows: int = 20) -> tuple[list[dict], int]:
    """Return (matching rows as dicts keyed by header, 1-based header row number)."""
    import openpyxl

    if not workbook.is_file():
        raise InventoryError(f"Inventory workbook not found: {workbook}")
    wb = openpyxl.load_workbook(workbook, read_only=True, data_only=True)  # never saved
    try:
        if sheet not in wb.sheetnames:
            raise InventoryError(f"Worksheet '{sheet}' not found (have: {', '.join(wb.sheetnames)})")
        rows = wb[sheet].iter_rows(values_only=True)
        header, header_row = None, 0
        for i, r in enumerate(rows, start=1):
            names = [str(c).strip() if c is not None else "" for c in r]
            if filter_column in names and all(c in names for c in REQUIRED_COLUMNS):
                header, header_row = names, i
                break
            if i >= header_search_rows:
                break
        if header is None:
            raise InventoryError(f"No header row with {', '.join((filter_column,) + REQUIRED_COLUMNS)} "
                                 f"in the first {header_search_rows} rows of '{sheet}'")
        fi = header.index(filter_column)
        out = []
        for r in rows:
            if r is None or fi >= len(r) or r[fi] is None:
                continue
            if str(r[fi]).strip() != filter_value:
                continue
            out.append({h: _cell(v) for h, v in zip(header, r) if h})
        return out, header_row
    finally:
        wb.close()


def _disk_index(source_root: Path) -> dict[str, list[str]]:
    """NFC relative path -> actual relative paths on disk (more than one = ambiguous)."""
    index: dict[str, list[str]] = {}
    for dirpath, dirnames, filenames in os.walk(source_root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        for name in filenames:
            if name.startswith(".") or not name.lower().endswith(".pdf"):
                continue  # hidden files and macOS/NAS "._" companions
            rel = Path(dirpath, name).relative_to(source_root).as_posix()
            index.setdefault(_nfc(rel), []).append(rel)
    return index


def _relative(file_path: str, prefix: str) -> tuple[str | None, str]:
    """Strip the inventory prefix and reject unsafe relative paths. Returns (rel, problem)."""
    fp = _nfc(str(file_path).strip().replace("\\", "/"))
    pre = _nfc(prefix.strip().replace("\\", "/"))
    if pre and not pre.endswith("/"):
        pre += "/"
    if not fp.startswith(pre):
        return None, PREFIX_MISMATCH
    rel = fp[len(pre):]
    parts = PurePosixPath(rel).parts
    if not rel or rel.startswith("/") or any(p in ("..", ".", "") for p in parts):
        return None, BAD_PATH
    return rel, ""


def resolve(rows: list[dict], source_root: Path, prefix: str) -> InventoryResult:
    available = source_root.is_dir()
    index = _disk_index(source_root) if available else {}
    folded: dict[str, list[str]] = {}
    for key, actual in index.items():
        folded.setdefault(key.casefold(), []).extend(actual)
    root_real = source_root.resolve() if available else source_root

    result, named = [], set()
    for r in rows:
        extra = {col: r.get(src) for src, col in OPTIONAL_COLUMNS.items() if src in r}
        file_name = _nfc(str(r.get("File Name") or ""))
        rel, problem = _relative(r.get("File Path") or "", prefix)
        row = InventoryRow(rel_path=rel or _nfc(str(r.get("File Path") or "")), file_name=file_name, extra=extra)
        result.append(row)
        if problem:
            row.resolve_status = problem
            row.note = "inventory path does not start with the configured prefix" if problem == PREFIX_MISMATCH \
                else "unsafe relative path ('..', '.', or absolute)"
            continue
        named.add(rel)
        if not available:
            row.resolve_status, row.note = MISSING, "source folder not available (NAS not mounted?)"
            continue
        exact = index.get(rel, [])
        if len(exact) == 1:
            candidate = source_root / exact[0]
            if not is_within(candidate, root_real):
                row.resolve_status, row.note = OUTSIDE_SOURCE, "resolves outside the source folder (symlink?)"
                continue
            row.resolve_status, row.resolved_path = OK, str(candidate)
            if PurePosixPath(rel).name != file_name:
                row.note = "File Name column differs from the name in File Path"
            size = extra.get("inv_size")
            if isinstance(size, int):
                try:
                    if candidate.stat().st_size != size:
                        row.note = (row.note + "; " if row.note else "") + "size differs from inventory"
                except OSError:
                    pass
        elif len(exact) > 1:
            row.resolve_status = AMBIGUOUS
            row.note = f"{len(exact)} files with Unicode-equivalent paths; not choosing one"
        elif folded.get(rel.casefold()):
            row.resolve_status = AMBIGUOUS
            row.note = f"only a case-different path exists ({len(folded[rel.casefold()])}); not matched automatically"
        else:
            row.resolve_status, row.note = MISSING, "no file at this relative path"

    # The same relative path listed more than once cannot be told apart: report, don't pick.
    counts: dict[str, int] = {}
    for row in result:
        counts[row.rel_path] = counts.get(row.rel_path, 0) + 1
    for row in result:
        if counts[row.rel_path] > 1:
            row.resolve_status, row.resolved_path = AMBIGUOUS, ""
            row.note = f"listed {counts[row.rel_path]} times in the inventory"

    unlisted = sorted(a for key, actual in index.items() for a in actual if key not in named)
    return InventoryResult(result, unlisted, available, 0)


def load(settings) -> InventoryResult:
    rows, header_row = read_inventory(settings.inventory_workbook, settings.sheet,
                                      settings.filter_column, settings.filter_value)
    res = resolve(rows, settings.source_root, settings.path_prefix)
    res.header_row = header_row
    return res
