"""End-to-end tests on SYNTHETIC PDFs only.

    .venv/bin/python -m unittest tests.test_pipeline -v
"""

from __future__ import annotations

import csv
import hashlib
import json
import shutil
import subprocess
import sys
import unittest
from datetime import datetime
from pathlib import Path

WORKSPACE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(WORKSPACE))
sys.path.insert(0, str(WORKSPACE / "tests"))

import make_synthetic_pdfs  # noqa: E402
from pdf_pipeline import run_pipeline  # noqa: E402
from pdf_pipeline.fields import find_dates  # noqa: E402
from pdf_pipeline.ocr import _page_spec  # noqa: E402


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _read_csv(path: Path) -> list[dict]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


class PipelineTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base = WORKSPACE / "tmp" / f"test_{datetime.now():%Y%m%d_%H%M%S_%f}"
        cls.src = cls.base / "src"
        cls.out = cls.base / "out"
        make_synthetic_pdfs.main(cls.src)
        cls.hashes = {p.name: _sha256(p) for p in cls.src.iterdir()}

        cfg = json.loads((WORKSPACE / "config" / "dictionaries.json").read_text(encoding="utf-8"))
        cfg.update(known_fields=["Zorvex"], well_prefixes=["Zorvex", "ZX"])
        cls.config = cls.base / "config.json"
        cls.config.write_text(json.dumps(cfg), encoding="utf-8")

        cls.rc = run_pipeline.main([str(cls.src), "--config", str(cls.config),
                                    "--output-dir", str(cls.out), "--quiet", "--xlsx"])
        cls.rows = {r["File Name"]: r for r in _read_csv(cls.out / "results.csv")}
        cls.evidence = _read_csv(cls.out / "evidence.csv")

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.base, ignore_errors=True)

    # ---- environment
    def test_imports(self):
        import dateutil, openpyxl, pandas, pypdf  # noqa: F401

    def test_ocr_tools(self):
        for tool in ("ocrmypdf", "tesseract", "gs"):
            self.assertIsNotNone(shutil.which(tool), tool)
        langs = subprocess.run(["tesseract", "--list-langs"], capture_output=True, text=True).stdout.split()
        self.assertIn("eng", langs)
        self.assertIn("msa", langs)

    # ---- batch behaviour
    def test_one_row_per_pdf_and_hidden_skipped(self):
        self.assertEqual(sorted(self.rows), ["corrupt.pdf", "no_fields.pdf",
                                             "scanned_report.pdf", "searchable_report.pdf"])

    def test_failure_does_not_stop_batch(self):
        self.assertTrue(self.rows["corrupt.pdf"]["Extraction Status"].startswith("FAILED"))
        self.assertEqual(self.rc, 1)  # non-zero because one file failed

    def test_sources_unchanged_and_nothing_written_beside_them(self):
        self.assertEqual({p.name: _sha256(p) for p in self.src.iterdir()}, self.hashes)

    def test_csv_has_utf8_bom(self):
        self.assertTrue((self.out / "results.csv").read_bytes().startswith(b"\xef\xbb\xbf"))
        self.assertTrue((self.out / "results.xlsx").exists())

    # ---- searchable PDF
    def test_searchable_fields(self):
        r = self.rows["searchable_report.pdf"]
        self.assertEqual(r["Report Title"], "FINAL SEISMIC PROCESSING REPORT ZORVEX 3D SURVEY")
        self.assertEqual(r["Created For (Operator)"], "Acme Exploration Sdn Bhd")
        self.assertEqual(r["Created By (Contractor)"], "Fictional Geodata Ltd")
        self.assertEqual(r["Company on Cover"], "Acme Exploration Sdn Bhd; Fictional Geodata Ltd")
        self.assertEqual(r["Report Date"], "1997-03-14")
        self.assertEqual(r["Total Pages"], "3")
        self.assertEqual(r["Field Name(s)"], "Zorvex")
        self.assertTrue(r["Extraction Status"].startswith("OK"), r["Extraction Status"])

    def test_wells_deduplicated_case_insensitively(self):
        wells = self.rows["searchable_report.pdf"]["Well Name(s)"].split("; ")
        self.assertEqual(wells, ["Zorvex-1", "Zorvex A-12", "ZX-3ST1", "Zorvex B-5"])

    def test_metadata_kept_separate(self):
        r = self.rows["searchable_report.pdf"]
        self.assertEqual(r["PDF Author"], "Synthetic Author")
        self.assertRegex(r["PDF Creation Date"], r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")
        self.assertNotEqual(r["Report Title"], "Embedded Metadata Title")

    def test_evidence_page_references(self):
        ev = {(e["Field"], e["Value"]): e for e in self.evidence if e["File Name"] == "searchable_report.pdf"}
        self.assertEqual(ev[("Well Name(s)", "Zorvex-1")]["Page(s)"], "2")
        self.assertEqual(ev[("Well Name(s)", "ZX-3ST1")]["Page(s)"], "3")
        self.assertEqual(ev[("Report Date", "1997-03-14")]["Page(s)"], "1")
        self.assertEqual(ev[("PDF Title", "Embedded Metadata Title")]["Source"], "pdf_metadata")

    # ---- scanned PDF (OCR)
    def test_scanned_pdf_selective_ocr(self):
        r = self.rows["scanned_report.pdf"]
        self.assertIn("OCR 1/2 pages", r["Extraction Status"])
        self.assertEqual(r["Report Title"], "SEISMIC INTERPRETATION REPORT")
        self.assertEqual(r["Report Date"], "2001-06-02")
        self.assertIn("Zorvex-7", r["Well Name(s)"])
        title_ev = next(e for e in self.evidence
                        if e["File Name"] == "scanned_report.pdf" and e["Field"] == "Report Title")
        self.assertEqual(title_ev["Source"], "ocr")

    def test_missing_values_left_empty(self):
        r = self.rows["no_fields.pdf"]
        self.assertTrue(r["Extraction Status"].startswith("PARTIAL"))
        for col in ("Report Title", "Company on Cover", "Report Date", "Well Name(s)"):
            self.assertEqual(r[col], "", col)

    # ---- safety guards
    def test_refuses_to_overwrite_output(self):
        rc = run_pipeline.main([str(self.src), "--config", str(self.config),
                                "--output-dir", str(self.out), "--quiet", "--ocr", "off"])
        self.assertEqual(rc, 2)

    def test_refuses_output_inside_source(self):
        rc = run_pipeline.main([str(self.src), "--output-dir", str(self.src / "out"), "--quiet"])
        self.assertEqual(rc, 2)
        self.assertFalse((self.src / "out").exists())

    def test_refuses_output_outside_workspace(self):
        rc = run_pipeline.main([str(self.src), "--output-dir", "/private/tmp/pdf_pipeline_out", "--quiet"])
        self.assertEqual(rc, 2)


class UnitTest(unittest.TestCase):
    def test_dates(self):
        self.assertEqual([d[0] for d in find_dates("Issued 3rd Ogos 1995")], ["1995-08-03"])
        self.assertEqual([d[0] for d in find_dates("March 1997")], ["1997-03"])  # no invented day
        self.assertEqual([d[0] for d in find_dates("04/03/1997")], ["1997-03-04"])
        self.assertEqual(find_dates("31/02/1997"), [])

    def test_default_well_patterns(self):
        from pdf_pipeline.config import load_rules
        from pdf_pipeline.extract import PageText
        from pdf_pipeline.fields import Extraction, find_wells

        ex = Extraction()
        text = "DULANG 3D SURVEY. Wells Dulang A-12, DL-B5 and Dulang-3ST1. Dulang is 12 km away. See Figure 4."
        find_wells([PageText(1, text, "text")], load_rules(), ex)
        self.assertEqual(ex.values["Well Name(s)"], ["Dulang A-12", "DL-B5", "Dulang-3ST1"])

    def test_page_spec(self):
        self.assertEqual(_page_spec([1, 2, 3, 7, 9, 10]), "1-3,7,9-10")


if __name__ == "__main__":
    unittest.main()
