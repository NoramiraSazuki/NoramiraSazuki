# PDF Metadata Extraction — User Guide

This tool reads PDF reports (searchable or scanned) and fills in a spreadsheet of
metadata: title, companies, dates, well names, field names, depths, page count, and more.
It writes **one row per PDF**, plus a second file that shows *where* each value came from.

Everything runs **on your Mac**. Nothing is uploaded, and the original PDFs are never changed.

---

## Contents

1. [Quick start](#1-quick-start)
2. [What you get](#2-what-you-get)
3. [Reading the results](#3-reading-the-results)
4. [Which columns are filled automatically](#4-which-columns-are-filled-automatically)
5. [Customising the rules](#5-customising-the-rules)
6. [Command options](#6-command-options)
7. [Safety and privacy](#7-safety-and-privacy)
8. [Troubleshooting](#8-troubleshooting)
9. [For maintainers](#9-for-maintainers)
10. [Dulang PDF Extraction (local browser app)](#10-dulang-pdf-extraction-local-browser-app)

---

## 1. Quick start

Open **Terminal** in VS Code (**Terminal → New Terminal**) and go to this folder:

```bash
cd ~/Documents/GitHub/NoramiraSazuki/Amira
```

**One PDF:**

```bash
../.venv/bin/python -m pdf_pipeline.run_pipeline "/path/to/report.pdf"
```

**A whole folder, including subfolders:**

```bash
../.venv/bin/python -m pdf_pipeline.run_pipeline "/path/to/folder" --recursive --xlsx
```

**A long run that keeps the Mac awake** (recommended for hundreds of files):

```bash
caffeinate -i ../.venv/bin/python -m pdf_pipeline.run_pipeline "/path/to/folder" --recursive --xlsx
```

> **Tip:** Drag a file or folder from Finder into the Terminal window to paste its path.
> Keep the quotes around paths that contain spaces.

When the run finishes you will see a summary like:

```
Done. 266 PDF(s) — FAILED: 2, OK: 4, PARTIAL: 260
Results: .../Amira/output/run_20261015_101557
```

### How long does it take?

| PDF type | Speed |
|---|---|
| Searchable (text can be selected) | Seconds per file |
| Scanned (images of pages) | Minutes per file — every page is OCR'd |

As a guide, 266 mostly scanned reports (4.8 GB) took about 4 hours.

---

## 2. What you get

Every run creates a **new** folder, so earlier results are never overwritten:

```
Amira/output/run_<date>_<time>/
├── results.csv     ← one row per PDF, 46 columns (open in Excel)
├── results.xlsx    ← same as results.csv (only with --xlsx)
└── evidence.csv    ← where each value was found, and how sure the tool is
```

A log of the run is saved in `Amira/logs/run_<date>_<time>.log`.

Results are saved as each PDF finishes. If a run is interrupted, the rows already done are kept.

---

## 3. Reading the results

### `results.csv`: the EXTRACTION STATUS column

The last column says how each file went:

| Status | Meaning | What to do |
|---|---|---|
| **OK** | Title, Created By, Created For and CREATE_DATE were all found | Spot-check |
| **PARTIAL** | At least one of those four is empty. Everything else found is filled in | Normal for older reports; fill gaps by hand if needed |
| **NO TEXT** | No readable text, even after OCR | Check the PDF manually |
| **FAILED** | The file could not be read (corrupt, password-protected, OCR crash) | See the message after `FAILED \|` |

Extra notes may follow the status, separated by `|`:

- `OCR 12/40 pages` — 12 of 40 pages were scanned images and needed OCR
- `missing: Created By, CREATE_DATE` — required columns that stayed empty
- `review: Spud Date (multiple values)` — the document gives more than one value; check which is right
- `review: Title (document and metadata titles differ)` — two different titles were found; both are in `evidence.csv`

Several values in one cell are separated by semicolons, e.g. `Dulang A-12; Dulang B-5`.

### `evidence.csv`: why each value was chosen

Each row is one value for one PDF:

| Column | Meaning |
|---|---|
| Attribute | Which result column (e.g. `Spud Date`) |
| Value | The value found |
| Source Page(s) | Page number(s) where it appears |
| Source | `text`, `ocr`, `pdf_metadata`, `filesystem`, `pdf_structure` or `config` |
| Extraction Method | The rule that found it (e.g. `'Spud Date' label`) |
| Validation Status | See below |
| Context | A short piece of the surrounding text, to check by eye |
| Confidence | HIGH / MEDIUM / LOW |

**Validation Status values:**

| Status | Meaning | Confidence |
|---|---|---|
| FILESYSTEM / PDF_STRUCTURE | Read from the file itself (size, pages) | HIGH |
| LABELLED | Written next to a label, e.g. `Spud Date: 3 June 1998` | HIGH |
| CONFIRMED | Document text and PDF metadata agree | HIGH |
| CONFIG | Set by one of your rules in `dictionaries.json` | HIGH |
| PATTERN | Matched a name pattern (wells, fields) | MEDIUM |
| INFERRED | Best guess, e.g. the title block on the cover | LOW |
| …`; MULTIPLE_VALUES` | More than one value found | LOW |
| **CANDIDATE** | Seen but **not exported**, because it was too uncertain | — |
| EXCLUDED_SECTION | A bibliography/reference section that was ignored | — |
| PDF_METADATA | Embedded PDF info, for reference only | — |
| NOT_FOUND | Nothing reliable found | — |
| REFERENCE_DATA_REQUIRED | Needs your classification rules (see section 5) | — |
| MANUAL | Filled in your data system, not from the PDF | — |
| NOT_ENCRYPTED / NEVER_EXPORTED | Encryption columns | — |

> **Useful Excel filters on `evidence.csv`:**
> - `Validation Status` = **CANDIDATE** → values the tool deliberately left out; you may want some of them
> - `Confidence` = **LOW** → values worth checking first
> - `Attribute` = **Excluded Section** → confirms a references section was skipped

---

## 4. Which columns are filled automatically

### Always filled (from the file itself)

| Column | Notes |
|---|---|
| File Name, ORIGINAL_FILE_NAME | Name as found in the folder |
| File Size | In bytes; DIGITAL_SIZE_UOM = `B` |
| File Format | e.g. `PDF` |
| PAGE_COUNT | Number of pages |
| DECRYPTION_TYPE | Encryption algorithm only (e.g. `Standard AES-128`); empty if not encrypted |

### Filled when the document states them

| Column | Found from |
|---|---|
| Title | Title block on the cover pages |
| Created By | Labels like `Prepared by:`, `Submitted by:`, `Contractor:` |
| Created For | Labels like `Prepared for:`, `Client:`, `Operator:` |
| CREATE_DATE | The report's issue date: `Date:`, `Report Date:`, `Issued:` |
| Spud Date, Completion Date, Transmittal Date, TRIP_DATE | Only a date written right after its own label |
| Top Depth / Base Depth (+ UOM) | `Top Depth:`, `Bottom Log Interval`, `Logged Interval: 1200 - 2450 m` |
| Unique Wellbore Identifier | `UWI:`, `Wellbore ID:`, `API No:` |
| Well Name | Names starting with a known prefix (`Dulang-1`, `DL A-12`) or `Well:` lines |
| FIELD_NAME | Your `known_fields` list, or a `Field:` line |
| Model | A line starting with `Model:` |
| LOGGING_COMPANY, LOG_TITLE | `Logging Company:`, `Log Title:` |
| REPORT_LOG_RUN | `Run No: 2` |
| SCALE_RATIO | `Scale: 1:200` |

### Need your reference data

**Dataset Type, Sub Dataset Type, DOCUMENT_TYPE, WELL_LOG_CLASS_ID**: add `classification_rules` (section 5).

### Filled in your data system (or with `constant_columns`)

Data Status, Description, Remarks, MEDIA_TYPE, BA_LONG_NAME_1, BA_TYPE_1, DATA_STORE_NAME,
DATA_STORE_TYPE, BAR_CODE, LOCATION_ID, SOURCE, ROW_QUALITY, CHECKED_BY_BA_ID.

### Never filled

**DECRYPT_KEY** is always empty. Passwords and keys are never extracted or exported.

### How the tool avoids wrong values

- **Uncertain means empty.** If a value isn't clearly stated, the cell stays empty. Anything it noticed goes to `evidence.csv` as a CANDIDATE.
- **Bibliographies are ignored.** Text under headings like *References*, *Bibliography* or *Literatur* describes *other* publications, so it's skipped up to the next *Appendix*. Single citation lines elsewhere (e.g. footnotes) are skipped too.
- **Dates are kept apart.** The report issue date (CREATE_DATE), the publication date (`Received`, `Published`, `©`) and the PDF's own creation date are different things. Only the first is used for CREATE_DATE; the other two appear in `evidence.csv`.
- **Depths need a unit.** `1250 m` is accepted; a bare `1250` becomes a CANDIDATE. Depth references such as MD or TVD go in `evidence.csv`, never in the UOM columns.
- **No assumptions about identity.** Similar well names (`Dulang-1` vs `Dulang 1`) are kept as separate values; only differences in capital letters are merged.

---

## 5. Customising the rules

All rules are in **`pdf_pipeline/config/dictionaries.json`**. Open it in VS Code, edit it, save, and re-run. No coding needed. Keep the JSON format (quotes, commas, square brackets).

| To… | Edit this key | Example |
|---|---|---|
| Recognise more well names | `well_prefixes` | `["Dulang", "DL", "DUL", "Bekok"]` |
| Recognise more field names | `known_fields` | `["Dulang", "Bekok"]` |
| Add contractor or client labels | `contractor_labels`, `operator_labels` | add `"Company"` to `operator_labels` for log headers |
| Accept unlabelled cover dates as CREATE_DATE | `create_date_fallback` | `true` (default `false`) |
| Read 03/04/1997 as March 4 instead of April 3 | `date_dayfirst` | `false` |
| Change which empty columns make a row PARTIAL | `required_columns` | `["Title", "CREATE_DATE"]` |
| Put the same value in every row | `constant_columns` | `{"DATA_STORE_NAME": "NAS-01", "MEDIA_TYPE": "Digital"}` |

### Classifying documents (Dataset Type, DOCUMENT_TYPE…)

Add rules to `classification_rules`. The first matching rule fills its columns, but only cells that are still empty:

```json
"classification_rules": [
  {
    "name": "Seismic reports",
    "filename_regex": "^PMO3DULGSE",
    "set": {
      "Dataset Type": "Seismic",
      "DOCUMENT_TYPE": "Seismic Report"
    }
  },
  {
    "name": "Well logs by keyword",
    "text_keywords": ["COMPOSITE LOG", "WIRELINE LOG"],
    "set": { "Dataset Type": "Well Log" }
  }
]
```

- `filename_regex` matches the file name; `^PMO3DULGSE` means "starts with PMO3DULGSE".
- `text_keywords` are searched in the title and cover pages.
- If a rule has both, both must match.

The tool checks the config when it starts. If something is wrong (for example a column name that doesn't exist), it stops with a clear message before processing anything.

---

## 6. Command options

| Option | What it does |
|---|---|
| `--recursive` | Include PDFs in subfolders |
| `--xlsx` | Also write `results.xlsx` |
| `--quiet` | Show only the final totals in Terminal (no file names) |
| `--ocr auto` | Default: OCR only pages that have no text |
| `--ocr force` | OCR every page |
| `--ocr off` | No OCR (fast, but scanned pages give no text) |
| `--lang eng+msa` | OCR languages (default English + Malay) |
| `--min-chars 20` | Pages with fewer characters than this are OCR'd |
| `--jobs 5` | CPU cores used for OCR (default: half of them) |
| `--ocr-timeout 3600` | Maximum OCR seconds per PDF |
| `--output-dir output/my_run` | Choose the results folder (must not exist yet) |
| `--config my_rules.json` | Use a different rules file |
| `--keep-temp` | Keep OCR working files in `tmp/` (for troubleshooting) |

Full help: `../.venv/bin/python -m pdf_pipeline.run_pipeline --help`

---

## 7. Safety and privacy

- **Original PDFs are read-only.** They are opened for reading only; a size and date check after each file confirms nothing changed.
- **Outputs stay in this folder.** The tool refuses to write outside `Amira/`, or inside the folder being processed.
- **Nothing is overwritten.** Each run needs a new output folder.
- **Everything runs offline.** OCR (Tesseract via OCRmyPDF) and extraction run on this Mac; no cloud services or AI models are called.
- **Logs contain no document text**, only file names and statuses.
- **Not uploaded to GitHub.** `output/`, `logs/`, `tmp/` and the data folders are listed in `.gitignore`.
- **Passwords are never used or exported.** Password-protected PDFs are reported as FAILED and are not cracked.

> Results files contain extracted company data. Treat them as confidential, just like the original PDFs.

---

## 8. Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| `Config file not found` | `dictionaries.json` was moved | Keep it at `pdf_pipeline/config/dictionaries.json`, or pass `--config path` |
| `Config error: ... not a configurable column` | Typo in a column name in your rules | Use the exact column names from section 4 |
| `Output folder already exists` | `--output-dir` points to an old run | Choose a new name, or leave the option out |
| `Refusing to write inside the source folder` | Output folder is inside the PDF folder | Use the default `output/` |
| `FAILED \| Encrypted PDF requires a password` | Password-protected PDF | Get an unprotected copy |
| `FAILED \| ... DependencyError ... cryptography` | AES-encrypted PDF; support isn't installed | Ask your maintainer to install `cryptography` in `.venv` |
| `FAILED \| ocrmypdf exit code 15: DecompressionBombError` | A page image is extremely large | Known limitation; check the file manually |
| `FAILED \| ocrmypdf exit code 4: ... could not be decoded` | Part of the PDF is damaged | Known limitation; check the file manually |
| A value is missing | Not clearly labelled in the PDF | Look for CANDIDATE rows in `evidence.csv`; adjust the rules in section 5 |
| A value is wrong | A rule matched the wrong text | Check `Context` and `Source Page(s)` in `evidence.csv`, then tighten the rule |
| Run is very slow | Many scanned pages | Normal. Use `caffeinate -i` so the Mac doesn't sleep |

---

## 9. For maintainers

### Folder layout

```
Amira/
├── README.md               ← this guide
├── pdf_pipeline/
│   ├── run_pipeline.py     ← command-line entry point; builds each results row
│   ├── fields.py           ← extraction rules (titles, dates, depths, wells, …)
│   ├── sections.py         ← finds and skips bibliography / citation text
│   ├── extract.py          ← reads PDF text, metadata, encryption type (pypdf)
│   ├── ocr.py              ← runs OCRmyPDF on selected pages, in tmp/
│   ├── schema.py           ← the 46 result columns and status codes
│   ├── config.py           ← loads and checks dictionaries.json
│   ├── export.py           ← UTF-8 BOM CSV / Excel writing
│   ├── safety.py           ← write-location and no-overwrite guards
│   ├── config/dictionaries.json
│   └── tests/              ← tests on synthetic (made-up) PDFs
├── output/                 ← results (one folder per run)
├── logs/                   ← run logs
└── tmp/                    ← OCR working files (cleaned up automatically)
```

The Python environment is one level up, in `../.venv` (Python 3.12).

### Requirements

- **Homebrew tools:** `tesseract`, `tesseract-lang`, `ghostscript`, `ocrmypdf`
- **Python packages:** see `../requirements.txt` (`pypdf`, `pandas`, `openpyxl`, `python-dateutil`, `reportlab` for tests)

### Running the tests

```bash
cd ~/Documents/GitHub/NoramiraSazuki/Amira
../.venv/bin/python -m unittest discover -s pdf_pipeline/tests -v
```

The tests generate made-up PDFs in `tmp/` (searchable, scanned, encrypted, corrupt, a log header,
a completion report, and a journal article with a bibliography), run the full pipeline,
check every column, and delete the test files afterwards. They never touch real data.

---

## 10. Dulang PDF Extraction (local browser app)

A local web app for choosing documents (or page ranges) from the inventory, extracting text page by page
in the background, reviewing each page next to the original, and exporting the results. The command-line
tool above still works exactly as before. The app is separate from it and keeps its own data.

### Start and stop

```bash
cd ~/Documents/GitHub/NoramiraSazuki/Amira
./start_pilot.command
```

(Or double-click `start_pilot.command` in Finder.) The browser opens **http://127.0.0.1:8765/**.
Keep the Terminal window open. The Mac stays awake while the app is running.

- **Stop:** press Ctrl+C in that window, or click **Shut down** in the app. Unfinished work resumes the next time you start it.
- Closing the browser does **not** stop processing. Open the address again to check progress.
- First start: the app creates `~/AmiraPilotData/settings.toml`. Edit the paths and options there, then restart.

### Where data is kept

Everything the app writes goes into `~/AmiraPilotData/`, **outside** this git folder:
`pilot.sqlite3` (catalogue, jobs, reviews), `results/` (one JSON file per page and method),
`cache/previews/`, `tmp/` (OCR page images, deleted after each page by default), `exports/`, `logs/`.
The inventory workbook and the PDFs are only read, never changed.

### Workflow

1. **Catalogue:** documents come from the `File Inventory` sheet, filtered to `Sub-folder = Seismic Report`.
   Inventory paths are mapped to the NAS folder. Rows that cannot be matched exactly are shown as
   MISSING / AMBIGUOUS / OUTSIDE_SOURCE and cannot be selected. Files on disk that the inventory does
   not list are shown separately.
   The **Title** column comes from column H of sheet `DOC-Seismic Report` in `Dulang_DIRt_v1.xlsx`
   (set under `[titles]` in `settings.toml`). Rows are matched by the path in its `File Location`
   column, not by file name. Blank cells stay blank, and a path listed twice gets no title. This title is
   catalogue information from that workbook, not text extracted from the PDF. Exports show it as
   "Catalogue Title".
2. Select documents, optionally type page ranges (`1-20, 45`, using the PDF's own page numbers),
   choose a mode, and click **Queue extraction job**. Modes that are not implemented yet are greyed out.
3. **Jobs:** shows progress, pages reused, failed pages and the ETA. You can **Pause**, **Resume**,
   **Cancel**, **Retry failed** and **Export results** here.
4. **Review:** the rendered page sits next to each method's text, along with the routing decision and
   its evidence, and the automatic checks. You can retry a page with another method, and mark it
   approved / corrected / rejected. Corrections are saved as new revisions, and the original
   extraction is never changed.

### How pages are routed

For each page the app measures the amount of text, the text area, the image area, unreadable glyphs,
rotation and page size. It then applies the rules in `pilot/config/routing.json`. A scanned page with a
small digital header still goes to OCR. The digital text layer is always kept as well, so nothing found
digitally is lost. Every result records the method, engine version and settings it came from. Results
are reused only when the file's SHA-256 and those settings are unchanged.

### Status words: what they mean and what they don't

- **Processing** (done / failed) means an engine produced output. It does **not** mean the output is correct.
- **Automatic checks** (empty, low OCR confidence, garbled text, the text layer has more text than OCR)
  point to pages worth looking at. A page with no flags is still not verified.
- **Review** (approved / corrected / rejected) is the only status that means a person checked the page.

### Known limitations (pilot)

- Table extraction and local multimodal comparison are **not implemented yet**. Tables come out as plain text.
- Tesseract can lose minus signs and similar marks in scanned tables (seen on synthetic test pages). Check numbers by eye.
- Reading order is the PDF's content order, which is not verified for multi-column pages.
- Figures and seismic sections are kept as page images with their captions. They are not digitised.

### Running real data, and sharing errors

Run the app yourself, in your own Terminal, **not** through Claude Code. Use the **Copy sanitised report**
button on a job when you need help. It contains only IDs, statuses and error types, with no file names
or document text. The log files in `~/AmiraPilotData/logs/` follow the same rule.

Processing needs no internet connection, only access to the NAS. The pilot packages (`requirements-pilot.txt`,
about 25 MB) were downloaded once at install time.

### Tests (synthetic files only)

```bash
cd ~/Documents/GitHub/NoramiraSazuki/Amira
../.venv/bin/python -m unittest discover -s pilot/tests -t . -v
RUN_SLOW=1 ../.venv/bin/python -m unittest pilot.tests.test_pilot.LongDocumentTest   # 4,579-page document
```
