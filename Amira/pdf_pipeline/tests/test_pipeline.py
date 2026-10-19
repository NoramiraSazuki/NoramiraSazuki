"""End-to-end tests on SYNTHETIC PDFs only.

    python -m unittest discover -s pdf_pipeline/tests -v    (run from the folder that contains pdf_pipeline/)
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

TESTS_DIR = Path(__file__).resolve().parent
# Works whether tests/ sits inside pdf_pipeline/ or beside it.
PACKAGE_PARENT = next(p for p in TESTS_DIR.parents if (p / "pdf_pipeline" / "__init__.py").is_file())
sys.path.insert(0, str(PACKAGE_PARENT))
sys.path.insert(0, str(TESTS_DIR))

import make_synthetic_pdfs  # noqa: E402
from pdf_pipeline import WORKSPACE, run_pipeline  # noqa: E402
from pdf_pipeline.config import DEFAULT_CONFIG, load_rules  # noqa: E402
from pdf_pipeline.extract import PageText  # noqa: E402
from pdf_pipeline.fields import Extraction, find_dates, find_depths, find_wells  # noqa: E402
from pdf_pipeline.ocr import _page_spec  # noqa: E402
from pdf_pipeline.schema import RESULT_COLUMNS  # noqa: E402

EXPECTED_COLUMNS = [
    "File Name", "File Size", "File Format", "Title", "Dataset Type", "Sub Dataset Type", "Data Status",
    "Description", "Created By", "Created For", "Remarks", "Top Depth", "Top Depth UOM", "Base Depth",
    "Base Depth UOM", "Spud Date", "Completion Date", "Transmittal Date", "Unique Wellbore Identifier",
    "Well Name", "Model", "FIELD_NAME", "CREATE_DATE", "DOCUMENT_TYPE", "PAGE_COUNT", "LOGGING_COMPANY",
    "SCALE_RATIO", "MEDIA_TYPE", "WELL_LOG_CLASS_ID", "LOG_TITLE", "REPORT_LOG_RUN", "TRIP_DATE",
    "ORIGINAL_FILE_NAME", "DECRYPT_KEY", "DECRYPTION_TYPE", "DIGITAL_SIZE_UOM", "BA_LONG_NAME_1",
    "BA_TYPE_1", "DATA_STORE_NAME", "DATA_STORE_TYPE", "BAR_CODE", "LOCATION_ID", "SOURCE", "ROW_QUALITY",
    "CHECKED_BY_BA_ID", "EXTRACTION STATUS",
]


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _read_csv(path: Path) -> list[dict]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def _synthetic_config(base: Path, **overrides) -> Path:
    cfg = json.loads(DEFAULT_CONFIG.read_text(encoding="utf-8"))
    cfg.update(known_fields=["Zorvex"], well_prefixes=["Zorvex", "ZX"], **overrides)
    path = base / f"config_{len(list(base.glob('config_*.json')))}.json"
    path.write_text(json.dumps(cfg), encoding="utf-8")
    return path


class PipelineTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base = WORKSPACE / "tmp" / f"test_{datetime.now():%Y%m%d_%H%M%S_%f}"
        cls.src = cls.base / "src"
        cls.out = cls.base / "out"
        make_synthetic_pdfs.main(cls.src)
        cls.hashes = {p.name: _sha256(p) for p in cls.src.iterdir()}

        cls.config = _synthetic_config(
            cls.base,
            classification_rules=[{"name": "logs by file name", "filename_regex": "^log_header",
                                   "set": {"Dataset Type": "Well Log", "DOCUMENT_TYPE": "Composite Log"}}],
            constant_columns={"DATA_STORE_NAME": "Synthetic Store"},
        )
        cls.rc = run_pipeline.main([str(cls.src), "--config", str(cls.config),
                                    "--output-dir", str(cls.out), "--quiet", "--xlsx"])
        with open(cls.out / "results.csv", encoding="utf-8-sig", newline="") as fh:
            cls.header = next(csv.reader(fh))
        cls.rows = {r["File Name"]: r for r in _read_csv(cls.out / "results.csv")}
        cls.evidence = _read_csv(cls.out / "evidence.csv")

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.base, ignore_errors=True)

    def ev(self, file_name: str, attribute: str, value: str | None = None) -> dict:
        return next(e for e in self.evidence if e["File Name"] == file_name and e["Attribute"] == attribute
                    and (value is None or e["Value"] == value))

    # ---- environment
    def test_imports(self):
        import dateutil, openpyxl, pandas, pypdf  # noqa: F401

    def test_ocr_tools(self):
        for tool in ("ocrmypdf", "tesseract", "gs"):
            self.assertIsNotNone(shutil.which(tool), tool)
        langs = subprocess.run(["tesseract", "--list-langs"], capture_output=True, text=True).stdout.split()
        self.assertIn("eng", langs)
        self.assertIn("msa", langs)

    # ---- schema / CSV
    def test_46_columns_in_exact_order(self):
        self.assertEqual(len(EXPECTED_COLUMNS), 46)
        self.assertEqual(RESULT_COLUMNS, EXPECTED_COLUMNS)
        self.assertEqual(self.header, EXPECTED_COLUMNS)

    def test_csv_has_utf8_bom(self):
        self.assertTrue((self.out / "results.csv").read_bytes().startswith(b"\xef\xbb\xbf"))
        self.assertTrue((self.out / "evidence.csv").read_bytes().startswith(b"\xef\xbb\xbf"))
        self.assertTrue((self.out / "results.xlsx").exists())

    def test_evidence_columns(self):
        self.assertEqual(list(self.evidence[0]), ["File Name", "Attribute", "Value", "Source Page(s)", "Source",
                                                  "Extraction Method", "Validation Status", "Context",
                                                  "Confidence"])

    def test_every_column_has_evidence(self):
        for name in self.rows:
            attrs = {e["Attribute"] for e in self.evidence if e["File Name"] == name}
            missing = [c for c in EXPECTED_COLUMNS if c != "EXTRACTION STATUS" and c not in attrs]
            self.assertEqual(missing, [], name)

    def test_empty_column_reasons(self):
        f = "no_fields.pdf"
        self.assertEqual(self.ev(f, "DECRYPT_KEY")["Validation Status"], "NEVER_EXPORTED")
        self.assertEqual(self.ev(f, "Dataset Type")["Validation Status"], "REFERENCE_DATA_REQUIRED")
        self.assertEqual(self.ev(f, "BAR_CODE")["Validation Status"], "MANUAL")
        self.assertEqual(self.ev(f, "Spud Date")["Validation Status"], "NOT_FOUND")
        self.assertEqual(self.ev(f, "DECRYPTION_TYPE")["Validation Status"], "NOT_ENCRYPTED")

    def test_confidence_levels(self):
        self.assertEqual(self.ev("searchable_report.pdf", "File Size")["Confidence"], "HIGH")
        self.assertEqual(self.ev("searchable_report.pdf", "Created By")["Confidence"], "HIGH")
        self.assertEqual(self.ev("searchable_report.pdf", "Well Name", "Zorvex-1")["Confidence"], "MEDIUM")
        self.assertEqual(self.ev("searchable_report.pdf", "Title", "FINAL SEISMIC PROCESSING REPORT ZORVEX 3D SURVEY")
                         ["Confidence"], "LOW")  # inferred and contradicted by metadata
        self.assertEqual(self.ev("completion_report.pdf", "Spud Date", "1998-06-03")["Confidence"], "LOW")
        self.assertEqual(self.ev("journal_article.pdf", "Title", make_synthetic_pdfs.ARTICLE_TITLE)
                         ["Confidence"], "HIGH")

    # ---- batch behaviour
    def test_one_row_per_pdf_and_hidden_skipped(self):
        self.assertEqual(sorted(self.rows), ["completion_report.pdf", "corrupt.pdf", "encrypted_report.pdf",
                                             "journal_article.pdf", "log_header.pdf", "no_fields.pdf",
                                             "scanned_report.pdf", "searchable_report.pdf"])

    def test_failure_does_not_stop_batch(self):
        r = self.rows["corrupt.pdf"]
        self.assertTrue(r["EXTRACTION STATUS"].startswith("FAILED"))
        self.assertEqual(r["File Size"], "27")  # filesystem attributes still reported
        self.assertEqual(self.rc, 1)  # non-zero because one file failed

    def test_sources_unchanged_and_nothing_written_beside_them(self):
        self.assertEqual({p.name: _sha256(p) for p in self.src.iterdir()}, self.hashes)

    # ---- filesystem attributes
    def test_filesystem_attributes(self):
        r = self.rows["searchable_report.pdf"]
        self.assertEqual(r["File Size"], str((self.src / "searchable_report.pdf").stat().st_size))
        self.assertEqual(r["DIGITAL_SIZE_UOM"], "B")
        self.assertEqual(r["File Format"], "PDF")
        self.assertEqual(r["ORIGINAL_FILE_NAME"], "searchable_report.pdf")
        self.assertEqual(r["PAGE_COUNT"], "3")
        self.assertEqual(self.ev("searchable_report.pdf", "File Size")["Validation Status"], "FILESYSTEM")

    # ---- searchable report (original workflow)
    def test_searchable_report(self):
        r = self.rows["searchable_report.pdf"]
        self.assertEqual(r["Title"], "FINAL SEISMIC PROCESSING REPORT ZORVEX 3D SURVEY")
        self.assertEqual(r["Created For"], "Acme Exploration Sdn Bhd")
        self.assertEqual(r["Created By"], "Fictional Geodata Ltd")
        self.assertEqual(r["CREATE_DATE"], "1997-03-14")
        self.assertEqual(r["FIELD_NAME"], "Zorvex")
        self.assertTrue(r["EXTRACTION STATUS"].startswith("OK"), r["EXTRACTION STATUS"])

    def test_wells_deduplicated_case_insensitively(self):
        wells = self.rows["searchable_report.pdf"]["Well Name"].split("; ")
        self.assertEqual(wells, ["Zorvex-1", "Zorvex A-12", "ZX-3ST1", "Zorvex B-5"])

    def test_metadata_kept_separate(self):
        r = self.rows["searchable_report.pdf"]
        self.assertNotEqual(r["Title"], "Embedded Metadata Title")
        self.assertEqual(self.ev("searchable_report.pdf", "PDF Author")["Value"], "Synthetic Author")
        self.assertEqual(self.ev("searchable_report.pdf", "PDF Title")["Validation Status"], "PDF_METADATA")
        self.assertEqual(r["Created By"], "Fictional Geodata Ltd")  # not the metadata author

    def test_conflicting_title_candidates_kept(self):
        r = self.rows["searchable_report.pdf"]
        self.assertIn("review: Title (document and metadata titles differ)", r["EXTRACTION STATUS"])
        cand = self.ev("searchable_report.pdf", "Title", "Embedded Metadata Title")
        self.assertEqual((cand["Validation Status"], cand["Source"]), ("CANDIDATE", "pdf_metadata"))

    def test_evidence_page_references(self):
        f = "searchable_report.pdf"
        self.assertEqual(self.ev(f, "Well Name", "Zorvex-1")["Source Page(s)"], "2")
        self.assertEqual(self.ev(f, "Well Name", "ZX-3ST1")["Source Page(s)"], "3")
        self.assertEqual(self.ev(f, "Well Name", "ZX-3ST1")["Validation Status"], "PATTERN")
        self.assertEqual(self.ev(f, "CREATE_DATE")["Source Page(s)"], "1")
        self.assertEqual(self.ev(f, "CREATE_DATE")["Validation Status"], "LABELLED")
        self.assertEqual(self.ev(f, "Created By")["Extraction Method"], "contractor/author label")

    # ---- scanned PDF (OCR workflow)
    def test_scanned_pdf_selective_ocr(self):
        r = self.rows["scanned_report.pdf"]
        self.assertIn("OCR 1/2 pages", r["EXTRACTION STATUS"])
        self.assertEqual(r["Title"], "SEISMIC INTERPRETATION REPORT")
        self.assertEqual(r["CREATE_DATE"], "2001-06-02")
        self.assertEqual(r["Well Name"], "Zorvex-7")
        self.assertEqual(r["FIELD_NAME"], "")  # field not inferred from the well name alone
        self.assertEqual(self.ev("scanned_report.pdf", "Title")["Source"], "ocr")

    # ---- completion report: typed dates, depths, UWI
    def test_typed_dates_are_distinguished(self):
        r = self.rows["completion_report.pdf"]
        self.assertEqual(r["CREATE_DATE"], "1998-08-20")
        self.assertEqual(r["Transmittal Date"], "1998-08-25")
        self.assertEqual(r["Completion Date"], "1998-07-30")
        self.assertEqual(r["TRIP_DATE"], "")

    def test_conflicting_values_flagged_not_hidden(self):
        r = self.rows["completion_report.pdf"]
        self.assertEqual(r["Spud Date"], "1998-06-03; 1998-06-05")
        self.assertIn("review: Spud Date (multiple values)", r["EXTRACTION STATUS"])
        self.assertIn("MULTIPLE_VALUES", self.ev("completion_report.pdf", "Spud Date", "1998-06-05")["Validation Status"])

    def test_depths_units_and_reference_kept_apart(self):
        r = self.rows["completion_report.pdf"]
        self.assertEqual((r["Top Depth"], r["Top Depth UOM"]), ("1250.5", "m"))
        self.assertEqual((r["Base Depth"], r["Base Depth UOM"]), ("8200", "ft"))
        self.assertEqual(self.ev("completion_report.pdf", "Top Depth Reference")["Value"], "MD")
        self.assertEqual(self.ev("completion_report.pdf", "Base Depth Reference")["Value"], "TVD")
        self.assertNotIn("0.1524", r["Top Depth"] + r["Base Depth"])  # sampling interval ignored

    def test_uwi_and_wells(self):
        r = self.rows["completion_report.pdf"]
        self.assertEqual(r["Unique Wellbore Identifier"], "ZX-000-B5-0001")
        self.assertEqual(r["Well Name"], "ZORVEX B-5")  # UWI fragment "ZX-000" not taken as a well
        self.assertEqual(r["Model"], "ZX Static Model v2")

    # ---- log header
    def test_log_header_attributes(self):
        r = self.rows["log_header.pdf"]
        self.assertEqual(r["Title"], "COMPOSITE LOG")
        self.assertEqual(r["LOGGING_COMPANY"], "Fictional Wireline Ltd")
        self.assertEqual(r["LOG_TITLE"], "Gamma Ray - Resistivity")
        self.assertEqual(r["REPORT_LOG_RUN"], "2")
        self.assertEqual(r["TRIP_DATE"], "1998-09-11")
        self.assertEqual(r["SCALE_RATIO"], "1:200; 1:500")
        self.assertEqual((r["Top Depth"], r["Top Depth UOM"], r["Base Depth"], r["Base Depth UOM"]),
                         ("1200", "m", "2450", "m"))
        self.assertEqual(self.ev("log_header.pdf", "Top Depth Reference")["Value"], "MDRT")
        self.assertEqual(r["Created By"], "")  # "Company:" is not guessed as contractor

    def test_classification_and_constants_from_config(self):
        r = self.rows["log_header.pdf"]
        self.assertEqual((r["Dataset Type"], r["DOCUMENT_TYPE"]), ("Well Log", "Composite Log"))
        self.assertEqual(self.ev("log_header.pdf", "Dataset Type")["Validation Status"], "CONFIG")
        self.assertEqual(self.rows["searchable_report.pdf"]["Dataset Type"], "")  # rule did not match
        self.assertEqual(r["DATA_STORE_NAME"], "Synthetic Store")

    # ---- regression: published article with a bibliography (the five KTB-style problems)
    def test_article_bibliography_excluded(self):
        ev = self.ev("journal_article.pdf", "Excluded Section")
        self.assertEqual((ev["Value"], ev["Source Page(s)"], ev["Validation Status"]),
                         ("References", "3", "EXCLUDED_SECTION"))

    def test_article_depths_only_from_report_itself(self):
        r = self.rows["journal_article.pdf"]
        self.assertEqual((r["Top Depth"], r["Top Depth UOM"], r["Base Depth"], r["Base Depth UOM"]),
                         ("1500", "m", "9000", "m"))  # not 200-400 m / 50 m from the references
        self.assertNotIn("77", r["Top Depth"])  # lowercase 'the top depth:' in a sentence is not a label

    def test_article_model_not_from_publication_title(self):
        r = self.rows["journal_article.pdf"]
        self.assertEqual(r["Model"], "Thermal Model ZT-1")  # from the appendix, after the references
        self.assertNotIn("Geobattery", r["Model"])
        self.assertNotIn("Reference Model 7", r["Model"])

    def test_article_field_name_not_from_phrases(self):
        r = self.rows["journal_article.pdf"]
        self.assertEqual(r["FIELD_NAME"], "")  # not "Quillon" (Field Laboratory) nor "Thermal"
        self.assertEqual(self.ev("journal_article.pdf", "FIELD_NAME")["Validation Status"], "NOT_FOUND")

    def test_article_well_in_reference_ignored(self):
        self.assertEqual(self.rows["journal_article.pdf"]["Well Name"], "")

    def test_article_title_confirmed_by_metadata(self):
        r = self.rows["journal_article.pdf"]
        self.assertEqual(r["Title"], make_synthetic_pdfs.ARTICLE_TITLE)  # both title lines, not just the first
        self.assertEqual(self.ev("journal_article.pdf", "Title", r["Title"])["Validation Status"], "CONFIRMED")
        self.assertNotIn("Title (", r["EXTRACTION STATUS"])  # no title review needed

    def test_article_dates_kept_distinct(self):
        r = self.rows["journal_article.pdf"]
        self.assertEqual(r["CREATE_DATE"], "")  # publication date is not the report issue date
        pub = self.ev("journal_article.pdf", "Publication Date")
        self.assertEqual(pub["Value"], "1998-03-12")
        self.assertIn("not used for CREATE_DATE", pub["Extraction Method"])
        self.assertEqual(self.ev("journal_article.pdf", "PDF Creation Date")["Validation Status"], "PDF_METADATA")
        cand = self.ev("journal_article.pdf", "CREATE_DATE", "1998-10-02")
        self.assertEqual(cand["Validation Status"], "CANDIDATE")

    # ---- encryption / secrets
    def test_encrypted_pdf_type_only_never_key(self):
        r = self.rows["encrypted_report.pdf"]
        self.assertEqual(r["DECRYPTION_TYPE"], "Standard RC4-128")
        self.assertEqual(r["Title"], "SEISMIC SURVEY REPORT")
        for row in self.rows.values():
            self.assertEqual(row["DECRYPT_KEY"], "")
        secret = make_synthetic_pdfs.SYNTHETIC_OWNER_PASSWORD.encode()
        for f in self.out.iterdir():
            self.assertNotIn(secret, f.read_bytes(), f.name)

    def test_decrypt_key_cannot_be_configured(self):
        with self.assertRaises(ValueError):
            load_rules(_synthetic_config(self.base, constant_columns={"DECRYPT_KEY": "x"}))
        with self.assertRaises(ValueError):
            load_rules(_synthetic_config(self.base, constant_columns={"File Size": "1"}))

    # ---- missing values
    def test_missing_values_left_empty(self):
        r = self.rows["no_fields.pdf"]
        self.assertTrue(r["EXTRACTION STATUS"].startswith("PARTIAL"))
        filled = {"File Name", "File Size", "File Format", "PAGE_COUNT", "ORIGINAL_FILE_NAME",
                  "DIGITAL_SIZE_UOM", "DATA_STORE_NAME", "EXTRACTION STATUS"}
        for col in EXPECTED_COLUMNS:
            if col not in filled:
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
        ex = Extraction()
        text = ("DULANG 3D SURVEY. Wells Dulang A-12, DL-B5 and Dulang-3ST1. Dulang is 12 km away. "
                "See Figure 4. UWI: DL-000-B5-0001")
        find_wells([PageText(1, text, "text")], load_rules(), ex)
        self.assertEqual(ex.get("Well Name"), ["Dulang A-12", "DL-B5", "Dulang-3ST1"])

    def test_depth_variants(self):
        ex = Extraction()
        text = ("Top Log Interval (ft): 4,100\nBottom Logged Interval 8200 FT TVDSS\n"
                "Interval: 1200 to 1300 m\nLogged Interval: 10 - 20\nTop Depth: 99000 m")
        find_depths([PageText(1, text, "text")], load_rules(), ex)
        tops = [(d.value, d.uom, d.reference) for d in ex.depths["Top Depth"]]
        bases = [(d.value, d.uom, d.reference) for d in ex.depths["Base Depth"]]
        self.assertEqual(tops, [("4100", "ft", "")])  # bare "Interval" is not a depth label any more
        self.assertEqual(bases, [("8200", "ft", "TVDSS")])
        candidates = {e.value: e.method for e in ex.evidence if e.validation == "CANDIDATE"}
        self.assertIn("no unit written", candidates["10"])
        self.assertIn("above plausible maximum", candidates["99000 m"])

    def test_title_completed_from_metadata_only_if_printed(self):
        from pdf_pipeline.fields import find_title
        rules = load_rules()
        page = PageText(1, "SEISMIC REPORT\nof the Quillon Area, Second Edition\nPrepared for: X Sdn Bhd", "text")
        ex = Extraction()
        find_title([page], rules, ex, "Seismic Report of the Quillon Area, Second Edition")
        self.assertEqual(ex.get("Title"), ["Seismic Report of the Quillon Area, Second Edition"])
        ex = Extraction()
        find_title([page], rules, ex, "Some Other Title Not In The Text")
        self.assertEqual(ex.get("Title"), ["SEISMIC REPORT"])
        self.assertTrue(ex.review)
        ex = Extraction()
        find_title([PageText(1, "no keyword here", "text")], rules, ex, "Unprinted Metadata Title")
        self.assertEqual(ex.get("Title"), [])  # uncertain -> empty
        ex = Extraction()
        find_title([page], rules, ex, "Microsoft Word - draft.doc")
        self.assertFalse(ex.review)  # placeholder metadata titles are ignored

    def test_report_date_vs_publication_date(self):
        from pdf_pipeline.fields import find_create_date
        ex = Extraction()
        page = PageText(1, "Published: 01 June 2001\nReport Date: 05 May 2000\n© 2001 Synthetic Press", "text")
        find_create_date([page], load_rules(), ex)
        self.assertEqual(ex.get("CREATE_DATE"), ["2000-05-05"])

    def test_reference_heading_in_table_of_contents_is_not_a_section(self):
        from pdf_pipeline.sections import filter_references
        pages = [PageText(1, "Contents\nIntroduction\nReferences\nAppendix\n", "text"),
                 PageText(2, "Logged Interval: 1 - 2 m\nWell: Zorvex-1\n", "text")]
        result = filter_references(pages, load_rules())
        self.assertEqual(result.sections, [])
        self.assertIn("Logged Interval", result.pages[1].text)

    def test_citation_lines(self):
        from pdf_pipeline.sections import is_citation_line
        self.assertTrue(is_citation_line("Smith, J. A., 1997. Heat flow at depth. J. Geophys. Res. 102, 1-9."))
        self.assertTrue(is_citation_line("[3] Writer D. (2001) Borehole logs. doi:10.1000/xyz"))
        self.assertFalse(is_citation_line("Logged Interval: 1500 - 9000 m MD"))
        self.assertFalse(is_citation_line("Report Date: 05 May 2000"))

    def test_page_spec(self):
        self.assertEqual(_page_spec([1, 2, 3, 7, 9, 10]), "1-3,7,9-10")


if __name__ == "__main__":
    unittest.main()
