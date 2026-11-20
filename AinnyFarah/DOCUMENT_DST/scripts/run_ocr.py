#!/usr/bin/env python3
"""Local OCR of DST report PDFs with OCRmyPDF/Tesseract.

- Reads source PDFs read-only from the NAS mount.
- Writes searchable PDFs, sidecar TXT files and logs under DOCUMENT_DST,
  mirroring the source subfolder structure.
- Verifies originals are unchanged (SHA-256 before/after).
- Prints status only; never prints document text.
"""
import hashlib
import json
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

SRC_ROOT = Path("/Volumes/homes/public/DulangOSS/20260925/DOCUMENTS/DST Document")
OUT_ROOT = Path("/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/DOCUMENT_DST")
PDF_DIR = OUT_ROOT / "searchable_pdf"
TXT_DIR = OUT_ROOT / "ocr_text"
LOG_DIR = OUT_ROOT / "logs"
TMP_DIR = OUT_ROOT / "tmp"


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def main():
    for d in (PDF_DIR, TXT_DIR, LOG_DIR, TMP_DIR):
        d.mkdir(parents=True, exist_ok=True)

    env = dict(os.environ, TMPDIR=str(TMP_DIR))
    sources = sorted(p for p in SRC_ROOT.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf")
    run_log = []

    for src in sources:
        rel = src.relative_to(SRC_ROOT)
        out_pdf = PDF_DIR / rel.parent / f"{rel.stem}_OCR.pdf"
        out_txt = TXT_DIR / rel.parent / f"{rel.stem}.txt"
        out_log = LOG_DIR / rel.parent / f"{rel.stem}.ocrmypdf.log"
        for p in (out_pdf, out_txt, out_log):
            p.parent.mkdir(parents=True, exist_ok=True)

        before = sha256(src)
        cmd = [
            "ocrmypdf",
            # Some pages carry a tiny pre-existing text layer (e.g. a stamp) over a
            # scanned image; --redo-ocr keeps that text and OCRs the rest of the page.
            "--redo-ocr",
            "--rotate-pages",
            "--language", "eng",
            "--sidecar", str(out_txt),
            "--jobs", str(os.cpu_count() or 2),
            str(src), str(out_pdf),
        ]
        started = datetime.now()
        with open(out_log, "w") as lf:
            rc = subprocess.run(cmd, stdout=lf, stderr=subprocess.STDOUT, env=env).returncode
        after = sha256(src)

        skipped = 0
        if out_txt.exists():
            skipped = out_txt.read_text(errors="ignore").count("[OCR skipped on page")
        entry = {
            "source": str(src),
            "searchable_pdf": str(out_pdf),
            "text": str(out_txt),
            "returncode": rc,
            "seconds": round((datetime.now() - started).total_seconds(), 1),
            "original_unchanged": before == after,
            "sha256": before,
            "pages_with_existing_text_skipped": skipped,
        }
        run_log.append(entry)
        print(f"[{'OK' if rc == 0 else 'FAIL rc=' + str(rc)}] {rel} | {entry['seconds']}s | "
              f"original unchanged: {entry['original_unchanged']} | pages skipped (had text): {skipped}")

    (LOG_DIR / "ocr_run_summary.json").write_text(json.dumps(run_log, indent=2))
    sys.exit(0 if all(e["returncode"] == 0 and e["original_unchanged"] for e in run_log) else 1)


if __name__ == "__main__":
    main()
