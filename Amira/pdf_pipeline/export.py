"""Write results/evidence as UTF-8-BOM CSV (Excel friendly). Files are never overwritten."""

from __future__ import annotations

import csv
from pathlib import Path

RESULT_COLUMNS = [
    "File Name",
    "Report Title",
    "Company on Cover",
    "Created By (Contractor)",
    "Created For (Operator)",
    "Report Date",
    "Total Pages",
    "Field Name(s)",
    "Well Name(s)",
    "PDF Author",
    "PDF Creation Date",
    "Extraction Status",
]

EVIDENCE_COLUMNS = ["File Name", "Field", "Value", "Page(s)", "Source", "Rule", "Context"]


class CsvWriter:
    """Streams rows to a new CSV file so partial results survive a crash mid-batch."""

    def __init__(self, path: Path, columns: list[str]):
        # mode "x" = create only; raises FileExistsError instead of overwriting.
        self._fh = open(path, "x", encoding="utf-8-sig", newline="")
        self._writer = csv.DictWriter(self._fh, fieldnames=columns, extrasaction="ignore")
        self._writer.writeheader()
        self._fh.flush()

    def write(self, row: dict) -> None:
        self._writer.writerow(row)
        self._fh.flush()

    def close(self) -> None:
        self._fh.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()


def write_xlsx(csv_path: Path, xlsx_path: Path) -> None:
    """Optional Excel copy of a results CSV."""
    import pandas as pd

    if xlsx_path.exists():
        raise FileExistsError(xlsx_path)
    df = pd.read_csv(csv_path, encoding="utf-8-sig", dtype=str, keep_default_na=False)
    df.to_excel(xlsx_path, index=False)
