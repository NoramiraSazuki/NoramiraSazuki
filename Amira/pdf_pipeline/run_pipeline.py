"""Extract report fields and metadata from one PDF or a folder of PDFs into CSV.

Usage (from the workspace folder):
    .venv/bin/python -m pdf_pipeline.run_pipeline "/path/to/file.pdf"
    .venv/bin/python -m pdf_pipeline.run_pipeline "/path/to/folder" --recursive

Source PDFs are only read. All outputs go to a new folder under output/, logs to logs/,
and OCR scratch files to tmp/ (removed afterwards unless --keep-temp).
"""

from __future__ import annotations

import argparse
import logging
import os
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

from . import WORKSPACE
from .config import DEFAULT_CONFIG, Rules, load_rules
from .export import EVIDENCE_COLUMNS, RESULT_COLUMNS, CsvWriter, write_xlsx
from .extract import PageText, PdfReadError, extract_page_texts, open_reader, read_metadata
from .fields import FIELD_NAMES, extract_fields
from .ocr import OcrError, ocr_available, run_ocr
from .safety import FileSnapshot, SafetyError, check_write_location, create_new_dir

log = logging.getLogger("pdf_pipeline")

# Fields whose absence makes a result PARTIAL (Field/Well names are often legitimately absent).
CORE_FIELDS = ["Report Title", "Company on Cover", "Report Date"]


def find_pdfs(source: Path, recursive: bool) -> list[Path]:
    if source.is_file():
        return [source]
    pattern = "**/*" if recursive else "*"
    return sorted(
        p for p in source.glob(pattern)
        if p.is_file() and p.suffix.lower() == ".pdf"
        # skip hidden files and macOS/NAS "._" resource-fork companions
        and not any(part.startswith(".") for part in p.relative_to(source).parts)
    )


def process_pdf(path: Path, rules: Rules, args, tmp_root: Path) -> tuple[dict, list[dict]]:
    row = {c: "" for c in RESULT_COLUMNS}
    row["File Name"] = path.name
    notes: list[str] = []
    before = FileSnapshot.of(path)
    work_dir = tmp_root / f"{time.monotonic_ns()}"

    try:
        with open(path, "rb") as fh:  # read-only
            reader = open_reader(fh)
            meta = read_metadata(reader)
            texts, failed_pages = extract_page_texts(reader)
        total = len(texts)
        row["Total Pages"] = total
        row["PDF Author"] = meta.author
        row["PDF Creation Date"] = meta.creation_date
        pages = [PageText(i, t, "text") for i, t in enumerate(texts, start=1)]
        if failed_pages:
            notes.append(f"text error on {len(failed_pages)} page(s)")

        if args.ocr == "force":
            need_ocr = list(range(1, total + 1))
        else:
            need_ocr = [p.number for p in pages if len(p.text.strip()) < args.min_chars]

        if need_ocr and args.ocr != "off" and total:
            log.debug("OCR on %d/%d pages", len(need_ocr), total)
            ocr_pdf = run_ocr(path, need_ocr, work_dir, args.lang, args.jobs, args.ocr_timeout)
            with open(ocr_pdf, "rb") as fh:
                ocr_texts, _ = extract_page_texts(open_reader(fh))
            if len(ocr_texts) != total:
                raise OcrError(f"OCR output has {len(ocr_texts)} pages, expected {total}")
            for n in need_ocr:
                pages[n - 1] = PageText(n, ocr_texts[n - 1], "ocr")
            notes.append(f"OCR {len(need_ocr)}/{total} pages")
        elif need_ocr:
            notes.append(f"{len(need_ocr)}/{total} pages without text (OCR off)")

        ex = extract_fields(pages, rules)
        for name in FIELD_NAMES:
            row[name] = "; ".join(ex.values[name])

        missing = [f for f in CORE_FIELDS if not ex.values[f]]
        if not any(p.text.strip() for p in pages):
            status = "NO TEXT"
        elif missing:
            status = "PARTIAL"
            notes.append("missing: " + ", ".join(missing))
        else:
            status = "OK"
        evidence = [
            {
                "File Name": path.name, "Field": e.field, "Value": e.value,
                "Page(s)": "; ".join(map(str, e.pages)), "Source": e.source,
                "Rule": e.rule, "Context": e.snippet,
            }
            for e in ex.evidence
        ]
        for label, value in (("PDF Title", meta.title), ("PDF Author", meta.author),
                             ("PDF Creator", meta.creator), ("PDF Producer", meta.producer),
                             ("PDF Creation Date", meta.creation_date)):
            if value:
                evidence.append({"File Name": path.name, "Field": label, "Value": value,
                                 "Page(s)": "", "Source": "pdf_metadata", "Rule": "embedded metadata",
                                 "Context": ""})
    except (PdfReadError, OcrError) as exc:
        status, evidence = "FAILED", []
        notes.append(str(exc))
    except Exception as exc:  # keep the batch going; details go to the log file only
        status, evidence = "FAILED", []
        notes.append(f"{type(exc).__name__}: {str(exc)[:150]}")
        log.debug("Unexpected error for %s", path.name, exc_info=True)
    finally:
        if not args.keep_temp:
            shutil.rmtree(work_dir, ignore_errors=True)

    if FileSnapshot.of(path) != before:
        notes.append("WARNING: source file size/mtime changed during run (not by this script)")
    row["Extraction Status"] = " | ".join([status] + notes)
    return row, evidence


def setup_logging(log_file: Path, quiet: bool) -> None:
    for handler in list(log.handlers):  # main() may be called more than once (tests)
        log.removeHandler(handler)
        handler.close()
    log.setLevel(logging.DEBUG)
    log.propagate = False
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    log.addHandler(fh)
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.WARNING if quiet else logging.INFO)
    ch.setFormatter(logging.Formatter("%(message)s"))
    log.addHandler(ch)


def parse_args(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("source", type=Path, help="PDF file or folder of PDFs (read-only)")
    ap.add_argument("--recursive", action="store_true", help="include PDFs in subfolders")
    ap.add_argument("--ocr", choices=["auto", "force", "off"], default="auto",
                    help="auto = OCR only pages without a text layer (default)")
    ap.add_argument("--lang", default="eng+msa", help="Tesseract languages (default eng+msa)")
    ap.add_argument("--min-chars", type=int, default=20,
                    help="pages with fewer text characters than this are OCR'd (default 20)")
    ap.add_argument("--jobs", type=int, default=max(1, (os.cpu_count() or 2) // 2))
    ap.add_argument("--ocr-timeout", type=int, default=3600, help="seconds per PDF (default 3600)")
    ap.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    ap.add_argument("--output-dir", type=Path, help="new folder for results (default output/run_<time>)")
    ap.add_argument("--keep-temp", action="store_true", help="keep OCR temp files in tmp/")
    ap.add_argument("--xlsx", action="store_true", help="also write results.xlsx")
    ap.add_argument("--quiet", action="store_true", help="terminal shows totals only, no file names")
    return ap.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    source = args.source.expanduser().resolve()
    if not source.exists():
        print(f"Source not found: {source}", file=sys.stderr)
        return 2

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = (args.output_dir or WORKSPACE / "output" / f"run_{stamp}").resolve()
    tmp_root = WORKSPACE / "tmp" / f"run_{stamp}"
    log_file = WORKSPACE / "logs" / f"run_{stamp}.log"
    try:
        for target in (out_dir, tmp_root, log_file):
            check_write_location(target, source)
        create_new_dir(out_dir)
    except SafetyError as exc:
        print(exc, file=sys.stderr)
        return 2
    log_file.parent.mkdir(exist_ok=True)
    tmp_root.mkdir(parents=True, exist_ok=True)
    setup_logging(log_file, args.quiet)

    rules = load_rules(args.config)
    pdfs = find_pdfs(source, args.recursive)
    if args.ocr != "off" and not ocr_available():
        log.warning("OCR tools not found; scanned pages will have no text")
    log.info("Found %d PDF(s). Output: %s", len(pdfs), out_dir)

    counts: dict[str, int] = {}
    with CsvWriter(out_dir / "results.csv", RESULT_COLUMNS) as results, \
         CsvWriter(out_dir / "evidence.csv", EVIDENCE_COLUMNS) as evidence:
        for i, pdf in enumerate(pdfs, start=1):
            t0 = time.time()
            row, ev = process_pdf(pdf, rules, args, tmp_root)
            results.write(row)
            for e in ev:
                evidence.write(e)
            status = row["Extraction Status"].split(" | ")[0]
            counts[status] = counts.get(status, 0) + 1
            # Log file name + status only; never extracted text.
            log.info("[%d/%d] %-8s %s (%.1fs)", i, len(pdfs), status, pdf.name, time.time() - t0)
            log.debug("Status detail for %s: %s", pdf.name, row["Extraction Status"])

    if args.xlsx:
        write_xlsx(out_dir / "results.csv", out_dir / "results.xlsx")
    if not args.keep_temp:
        shutil.rmtree(tmp_root, ignore_errors=True)

    summary = ", ".join(f"{k}: {v}" for k, v in sorted(counts.items())) or "no PDFs"
    print(f"Done. {len(pdfs)} PDF(s) — {summary}")
    print(f"Results: {out_dir}")
    print(f"Log:     {log_file}")
    return 0 if counts.get("FAILED", 0) == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
