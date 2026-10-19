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
import re
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

from . import WORKSPACE
from .config import DEFAULT_CONFIG, Rules, load_rules
from .export import CsvWriter, write_xlsx
from .extract import (PageText, PdfReadError, encryption_type, extract_page_texts, open_reader,
                      read_metadata)
from .fields import Evidence, Extraction, extract_fields
from .ocr import OcrError, ocr_available, run_ocr
from .safety import FileSnapshot, SafetyError, check_write_location, create_new_dir
from .schema import (CANDIDATE, CONFIG, EVIDENCE_COLUMNS, EXCLUDED, FILESYSTEM, MANUAL, MANUAL_COLUMNS,
                     NEVER_EXPORTED, NEVER_FILLED, NOT_ENCRYPTED, NOT_FOUND, PDF_METADATA, PDF_STRUCTURE,
                     REFERENCE_DATA, REFERENCE_DATA_COLUMNS, RESULT_COLUMNS, confidence)
from .sections import filter_references

log = logging.getLogger("pdf_pipeline")
STATUS = "EXTRACTION STATUS"


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


def _evidence_row(file_name: str, e: Evidence) -> dict:
    return {
        "File Name": file_name, "Attribute": e.attribute, "Value": e.value,
        "Source Page(s)": "; ".join(map(str, e.pages)), "Source": e.source,
        "Extraction Method": e.method, "Validation Status": e.validation, "Context": e.snippet,
        "Confidence": confidence(e.validation),
    }


def _empty_column_evidence(row: dict, ex: Extraction) -> None:
    """Give every results column at least one evidence row, explaining why an empty cell is empty."""
    exported = {e.attribute for e in ex.evidence if not e.validation.startswith(CANDIDATE)}
    for col in RESULT_COLUMNS:
        if col == STATUS or row[col] not in ("", None) or col in exported:
            continue
        if col in NEVER_FILLED:
            status, why = NEVER_EXPORTED, "secrets are never extracted or exported"
        elif col == "DECRYPTION_TYPE":
            status, why = NOT_ENCRYPTED, "PDF has no /Encrypt dictionary"
        elif col in REFERENCE_DATA_COLUMNS:
            status, why = REFERENCE_DATA, "needs classification_rules in dictionaries.json"
        elif col in MANUAL_COLUMNS:
            status, why = MANUAL, "set in the data system or constant_columns"
        else:
            has_candidate = any(e.attribute == col for e in ex.evidence)
            status = NOT_FOUND
            why = "only uncertain candidates found (see CANDIDATE rows)" if has_candidate else "no reliable evidence"
        ex.evidence.append(Evidence(col, "", [], "", why, status))


def _fill_row(row: dict, ex: Extraction) -> None:
    for col in RESULT_COLUMNS:
        if ex.get(col) and not row[col]:
            row[col] = "; ".join(ex.get(col))
    for attr in ("Top Depth", "Base Depth"):
        depths = ex.depths[attr]
        row[attr] = "; ".join(d.value for d in depths)
        # UOM list stays aligned with the depth list; left empty when no unit was written.
        row[f"{attr} UOM"] = "; ".join(d.uom for d in depths) if any(d.uom for d in depths) else ""


def process_pdf(path: Path, rules: Rules, args, tmp_root: Path) -> tuple[dict, list[dict]]:
    row = {c: "" for c in RESULT_COLUMNS}
    st = path.stat()
    row.update({
        "File Name": path.name,
        "ORIGINAL_FILE_NAME": path.name,  # as found in the source folder (the pipeline never renames)
        "File Size": st.st_size,
        "DIGITAL_SIZE_UOM": "B",
        "File Format": path.suffix.lstrip(".").upper(),
    })
    ex = Extraction()
    for attr in ("File Name", "ORIGINAL_FILE_NAME", "File Size", "DIGITAL_SIZE_UOM", "File Format"):
        ex.evidence.append(Evidence(attr, str(row[attr]), [], "filesystem", "file system attribute", FILESYSTEM))
    notes: list[str] = []
    before = FileSnapshot.of(path)
    work_dir = tmp_root / f"{time.monotonic_ns()}"

    try:
        with open(path, "rb") as fh:  # read-only
            reader = open_reader(fh)
            row["DECRYPTION_TYPE"] = encryption_type(reader)
            meta = read_metadata(reader)
            texts, failed_pages = extract_page_texts(reader)
        total = len(texts)
        row["PAGE_COUNT"] = total
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

        # Bibliography/citation text describes other publications: hide it from extraction.
        filtered = filter_references(pages, rules)
        doc = extract_fields(filtered.pages, rules, path.name, meta.title)
        doc.evidence[:0] = ex.evidence  # filesystem rows first
        ex = doc
        for sec in filtered.sections:
            ex.evidence.append(Evidence("Excluded Section", sec.heading, sec.pages, "text",
                                        f"reference section ({sec.lines} lines ignored)", EXCLUDED))
        if filtered.citation_lines:
            ex.evidence.append(Evidence("Excluded Citation Lines", str(sum(filtered.citation_lines.values())),
                                        sorted(filtered.citation_lines), "text",
                                        "citation-style lines outside reference sections", EXCLUDED))
        _fill_row(row, ex)
        ex.evidence.append(Evidence("PAGE_COUNT", str(total), [], "pdf_structure", "pypdf page count",
                                    PDF_STRUCTURE))
        if row["DECRYPTION_TYPE"]:
            ex.evidence.append(Evidence("DECRYPTION_TYPE", row["DECRYPTION_TYPE"], [], "pdf_structure",
                                        "/Encrypt dictionary (algorithm only)", PDF_STRUCTURE))
        # Embedded metadata: reference only, never copied into document columns.
        for label, value in (("PDF Title", meta.title), ("PDF Author", meta.author),
                             ("PDF Creator", meta.creator), ("PDF Producer", meta.producer),
                             ("PDF Creation Date", meta.creation_date)):
            if value:
                ex.evidence.append(Evidence(label, value, [], "pdf_metadata", "embedded metadata", PDF_METADATA))
        review = ex.review + ex.flag_multiple()
        if not any(p.text.strip() for p in pages):
            status = "NO TEXT"
        else:
            status = "OK"
    except (PdfReadError, OcrError) as exc:
        status = "FAILED"
        notes.append(str(exc))
        if isinstance(exc, PdfReadError) and exc.encryption:
            row["DECRYPTION_TYPE"] = exc.encryption
        review = []
    except Exception as exc:  # keep the batch going; details go to the log file only
        status = "FAILED"
        notes.append(f"{type(exc).__name__}: {str(exc)[:150]}")
        log.debug("Unexpected error for %s", path.name, exc_info=True)
        review = []
    finally:
        if not args.keep_temp:
            shutil.rmtree(work_dir, ignore_errors=True)

    if status != "FAILED":
        for col, value in rules.constant_columns.items():
            if not row[col]:
                row[col] = value
                ex.evidence.append(Evidence(col, value, [], "config", "constant_columns", CONFIG))
        missing = [c for c in rules.required_columns if row[c] in ("", None)]
        if missing and status == "OK":
            status = "PARTIAL"
        if missing:
            notes.append("missing: " + ", ".join(missing))
        if review:
            notes.append("review: " + "; ".join(review))
    for col in NEVER_FILLED:
        row[col] = ""  # secrets are never exported

    if FileSnapshot.of(path) != before:
        notes.append("WARNING: source file size/mtime changed during run (not by this script)")
    row[STATUS] = " | ".join([status] + notes)
    _empty_column_evidence(row, ex)
    return row, [_evidence_row(path.name, e) for e in ex.evidence]


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
    # pypdf's own warnings (e.g. "invalid pdf header") go to the log file, not the terminal.
    pypdf_log = logging.getLogger("pypdf")
    for handler in list(pypdf_log.handlers):
        pypdf_log.removeHandler(handler)
    pypdf_log.addHandler(fh)
    pypdf_log.propagate = False
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
    # Load rules before creating any folders, so a config error leaves nothing behind.
    try:
        rules = load_rules(args.config.expanduser())
    except (OSError, ValueError, KeyError, re.error) as exc:
        print(f"Config error: {exc}", file=sys.stderr)
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
            status = row[STATUS].split(" | ")[0]
            counts[status] = counts.get(status, 0) + 1
            # Log file name + status only; never extracted text.
            log.info("[%d/%d] %-8s %s (%.1fs)", i, len(pdfs), status, pdf.name, time.time() - t0)
            log.debug("Status detail for %s: %s", pdf.name, row[STATUS])

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
