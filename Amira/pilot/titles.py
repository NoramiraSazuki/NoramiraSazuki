"""Catalogue titles from a second, read-only mapping workbook (e.g. Dulang_DIRt_v1.xlsx).

The title is taken from a configured column LETTER (default H), because in the DIRt sheet the
header above the titles is not "Title". Rows are joined to the inventory by the relative path in
the location column (default "File Location", e.g. oss://.../20260925/DOCUMENTS/Seismic Report/X.pdf),
never by file name alone. Blank titles stay blank; duplicates are reported, not resolved.
The title is catalogue metadata from that workbook - it is not extracted from the PDF.
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl.utils import column_index_from_string


@dataclass
class TitleResult:
    titles: dict[str, str]                 # NFC relative path -> title (non-blank only)
    header_row: int = 0
    title_header: str = ""                 # header text found above the title column (informational)
    rows: int = 0
    blank: int = 0
    duplicates: list[str] = field(default_factory=list)
    bad_location: int = 0
    error: str = ""


def _nfc(s) -> str:
    return unicodedata.normalize("NFC", str(s).strip().replace("\\", "/"))


def read_titles(workbook: Path, sheet: str, title_column: str, location_header: str, path_prefix: str,
                header_search_rows: int = 20) -> TitleResult:
    import openpyxl

    if not workbook.is_file():
        return TitleResult({}, error=f"title workbook not found: {workbook.name}")
    t_idx = column_index_from_string(title_column.strip().upper()) - 1
    prefix = _nfc(path_prefix).rstrip("/") + "/"
    wb = openpyxl.load_workbook(workbook, read_only=True, data_only=True)  # never saved
    try:
        if sheet not in wb.sheetnames:
            return TitleResult({}, error=f"worksheet '{sheet}' not found in title workbook")
        it = wb[sheet].iter_rows(values_only=True)
        header, hrow = None, 0
        for i, r in enumerate(it, start=1):
            names = [str(c).strip() if c is not None else "" for c in r]
            if location_header in names:
                header, hrow = names, i
                break
            if i >= header_search_rows:
                break
        if header is None:
            return TitleResult({}, error=f"no '{location_header}' header in the first {header_search_rows} rows")
        loc_idx = header.index(location_header)
        res = TitleResult({}, header_row=hrow, title_header=header[t_idx] if t_idx < len(header) else "")
        seen: dict[str, int] = {}
        for r in it:
            loc = r[loc_idx] if loc_idx < len(r) else None
            if loc in (None, ""):
                continue
            res.rows += 1
            loc = _nfc(loc)
            pos = loc.find(prefix)
            if pos < 0:
                res.bad_location += 1
                continue
            rel = loc[pos + len(prefix):]
            seen[rel] = seen.get(rel, 0) + 1
            title = r[t_idx] if t_idx < len(r) else None
            title = " ".join(str(title).split()) if title not in (None, "") else ""
            if not title:
                res.blank += 1
            elif seen[rel] == 1:
                res.titles[rel] = title
        res.duplicates = sorted(k for k, n in seen.items() if n > 1)
        for k in res.duplicates:
            res.titles.pop(k, None)  # ambiguous: do not pick one
        return res
    finally:
        wb.close()


def load(settings) -> TitleResult:
    if not settings.title_workbook:
        return TitleResult({}, error="no title workbook configured")
    return read_titles(Path(settings.title_workbook).expanduser(), settings.title_sheet, settings.title_column,
                       settings.title_location_header, settings.path_prefix)
