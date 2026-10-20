# Dulang PDF Extraction: Setup and User Guide

Dulang PDF Extraction is a small app that runs **on your own Mac** and opens in your web browser. It helps you:

1. **Browse** every PDF in the Dulang OSS file inventory, as a folder tree like Finder.
2. **Extract text** from selected PDFs or page ranges, in the background, one job at a time.
3. **Review** each page next to the extracted text, then correct or approve it.
4. **Download** the results as text, CSV, Excel or JSON.

All processing happens on your Mac. Documents are never uploaded to any website, cloud service or AI service.

> **Status: pilot.** Text extraction (digital text plus OCR for scanned pages) works. **Table extraction and
> AI-based (multimodal) extraction are not built yet.** Those options are greyed out in the app.

---

## Contents

1. [What you need](#1-what-you-need)
2. [One-time installation](#2-one-time-installation)
3. [First start and settings](#3-first-start-and-settings)
4. [Starting and stopping](#4-starting-and-stopping)
5. [Using the app](#5-using-the-app)
6. [What the statuses mean](#6-what-the-statuses-mean)
7. [Where your data is kept](#7-where-your-data-is-kept)
8. [Privacy and safety](#8-privacy-and-safety)
9. [Troubleshooting](#9-troubleshooting)
10. [Known limitations](#10-known-limitations)

---

## 1. What you need

### Computer
- **A Mac with Apple silicon** (M1 or newer). It was built and tested on an M4 MacBook Pro with 16 GB memory.
- **At least 20 GB of free disk space.** Jobs pause automatically if free space drops below 15 GB.
- An internet connection **only during installation**. Processing works offline, but needs access to the NAS.

### Access
You need your own access to all three of these:

| What | Why | Example location |
|---|---|---|
| **The Dulang NAS share**, mounted in Finder | The PDFs are read from here (read-only) | `/Volumes/homes/public/DulangOSS` |
| **The file inventory workbook** | The list of files the catalogue shows | `Dulang_OSS_All_Batches_File_Inventory.xlsx` |
| **The DIRt mapping workbook** (optional) | Fills the **Title** column (column H of each sheet) | `Dulang_DIRt_v1.xlsx` |

The workbooks are **not** included in the GitHub repository, because they are confidential. Ask the data owner for access.

### Software
| Software | Purpose | Approx. size |
|---|---|---|
| Homebrew | Installs the tools below | ~0.5 GB |
| Python 3.12 | Runs the app | ~0.1 GB |
| Tesseract OCR + language data | Reads scanned pages (English + Malay) | ~0.7 GB (all languages) |
| Python packages (see `requirements.txt` and `Amira/requirements-pilot.txt`) | PDF reading, web app, Excel export | ~0.2 GB |
| Git | Downloads the code | included with Xcode Command Line Tools |

No AI model is needed for the current version.

---

## 2. One-time installation

Open **Terminal** (Applications → Utilities → Terminal) and run these steps one at a time.

**Step 1: install Homebrew** (skip this if `brew --version` already works):
```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```
Follow the instructions it prints at the end, which add `brew` to your PATH.

**Step 2: install Python, Tesseract and Git:**
```bash
brew install python@3.12 tesseract tesseract-lang git
```
`tesseract-lang` contains all OCR languages, about 650 MB. The app uses English and Malay (`eng+msa`).

**Step 3: download the code:**
```bash
mkdir -p ~/Documents/GitHub
cd ~/Documents/GitHub
git clone https://github.com/NoramiraSazuki/NoramiraSazuki.git
cd NoramiraSazuki
```

**Step 4: create the Python environment and install the packages:**
```bash
/opt/homebrew/bin/python3.12 -m venv .venv
.venv/bin/pip install -r requirements.txt -r Amira/requirements-pilot.txt
```

**Step 5: check that it works** (uses made-up test files only; about 1 minute):
```bash
cd Amira
../.venv/bin/python -m unittest discover -s pilot/tests -t .
```
The last line should say `OK`.

**Step 6: mount the NAS.** In Finder, choose **Go → Connect to Server…** and connect to the Dulang NAS. Check
that the batch folders (`20260925`, `20261006`, …) are visible at `/Volumes/homes/public/DulangOSS`.

---

## 3. First start and settings

Start the app once (see [section 4](#4-starting-and-stopping)). On the first start it creates a settings file:

```
~/AmiraPilotData/settings.toml
```

**Open it in TextEdit or VS Code and fix the paths for your Mac.** The file is created with the original
author's paths, which will not exist on your Mac. Change these lines:

```toml
[paths]
inventory_workbook = "/path/to/your/Dulang_OSS_All_Batches_File_Inventory.xlsx"
source_root = "/Volumes/homes/public/DulangOSS"

[titles]
title_workbook = "/path/to/your/Dulang_DIRt_v1.xlsx"   # leave as is if you don't have it; titles stay blank
```

> **Tip:** to get a file's full path, select it in Finder, hold **Option**, then right-click and choose **Copy "…" as Pathname**.

Save the file, then restart the app. You can leave the other settings as they are. The most useful ones are:

| Setting | Default | Meaning |
|---|---|---|
| `batch_pages` | 20 | Pages processed between saved checkpoints |
| `ocr_lang` | `eng+msa` | OCR languages |
| `min_free_gb` | 15 | Jobs pause below this much free disk space |
| `large_job_pages` | 1000 | Jobs bigger than this need a second click to confirm |
| `scan_folder_timeout` | 60 | Seconds before a stalled NAS folder is skipped and reported |
| `port` | 8765 | The app's address: `http://127.0.0.1:8765/` |

---

## 4. Starting and stopping

**Start:**
```bash
cd ~/Documents/GitHub/NoramiraSazuki/Amira
./start_pilot.command
```
You can also double-click `start_pilot.command` in Finder. Your browser opens **http://127.0.0.1:8765/**.

- **Keep the Terminal window open.** While the app runs, it also stops the Mac from sleeping.
- **Closing the browser does not stop processing.** Open the address again to check progress.
- **Stop:** press **Ctrl+C** in the Terminal window, or click **Shut down** in the app. Unfinished jobs
  continue from where they stopped the next time you start the app.

---

## 5. Using the app

### Catalogue tab
- **Filling the catalogue:** the inventory's PDFs appear within seconds, grouped by **batch folder** (`20260925`, `20261006`, …).
- **NAS check:** in the background, the app then checks the NAS **folder by folder**. Each file changes from
  **NOT_CHECKED** to **OK** (found) or **MISSING** (not found). A slow NAS can make this take several minutes.
- **Browsing:** use the **Batch** dropdown, the search box (file name, title or folder) and the status filter
  to narrow the list. The list is split into pages; choose 50, 100, 200 or 500 rows per page.
- **Selecting:**
  - Tick individual PDFs, or tick a **folder** to select everything inside it.
  - The checkbox in the column header selects everything matching the current filter, on all pages.
  - Only **OK** files can be selected.
- **Page ranges:** optionally type them in **Pages to process**, using the PDF's own page numbers
  (`1-20, 45`). Leave it blank for the whole document.
- **Starting a job:** choose the extraction mode at the bottom and click **Queue extraction job**. Jobs over 1,000 pages
  ask you to click again to confirm, and show a rough time estimate.

### Jobs tab
- Shows progress, pages done, failed pages, pages reused and an estimated time left.
- **Pause**, **Resume**, **Cancel**, **Retry failed**: finished pages are always kept.
- **Export results** writes a folder under `~/AmiraPilotData/exports/`.
- **Download results (.zip)** downloads the whole job.
- **Copy sanitised report** copies a short status summary with IDs and error types only. It contains no file names
  or document text, so it is safe to share when asking for help.

### Review tab
- **Side by side:** the original page (left) next to the extracted text from each method (right).
- **Routing decision:** why the app chose digital text or OCR for the page, plus any automatic warnings.
- **Retry this page:** try another method; the earlier result is kept.
- **Human review:** mark the page **Approved**, **Corrected** (type the corrected text) or **Rejected**. Every change is
  saved as a new revision, and the original extraction is never overwritten.
- **Download extracted pages:** choose a page range (blank = current page, `all` = whole document) and download
  as **Text (.txt)**, **Pages table (.csv)** or **Full detail (.json)**.

### How the app decides how to read each page
For every page it measures the amount of digital text, how much of the page is an image, unreadable
characters, rotation and page size. Pages with a good digital text layer are read directly. Scanned or
image-heavy pages go to OCR. The digital text is always kept as well, so nothing is lost. The rules are in
`Amira/pilot/config/routing.json`.

---

## 6. What the statuses mean

### NAS badge (top of the page)
- **● NAS – Connected** (green): the PDF folder can be reached.
- **● NAS – Not connected** (red): mount the NAS in Finder. Jobs wait and continue when it is back.

### File statuses in the catalogue
| Status | Meaning | Can be processed? |
|---|---|---|
| **OK** | Found on the NAS at exactly the inventory's path | Yes |
| **NOT_CHECKED** | Listed in the inventory; NAS check not finished yet | Not yet |
| **MISSING** | Not on the NAS (confirmed by a second look-up) | No |
| **FOLDER_UNREADABLE** | The NAS did not answer for this folder; not the same as missing | Not until re-checked |
| **AMBIGUOUS** | More than one possible match (e.g. listed twice, or only a different upper/lower case exists) | No; the app never guesses |

### Processing, checks and review are separate
- **Processing done** means a method produced output. **It does not mean the text is correct.**
- **Automatic checks** (empty page, low OCR confidence, garbled text, "digital text has more than OCR") point to
  pages worth looking at.
- **Review: approved/corrected** is the only status that means a person checked the page.

---

## 7. Where your data is kept

Everything the app creates is stored **outside the code folder**, in:

```
~/AmiraPilotData/
├── settings.toml     your settings
├── pilot.sqlite3     catalogue, jobs, reviews
├── results/          extracted text, one file per page and method
├── cache/previews/   page images for the Review tab
├── exports/          exported job folders and .zip files
├── backups/          automatic database backups
├── logs/             logs (IDs and statuses only, no document text)
└── tmp/              temporary OCR images (deleted after each page)
```

The **inventory workbook, the DIRt workbook and the PDFs are only read, never changed.**

---

## 8. Privacy and safety

- **Everything runs locally.** OCR uses Tesseract on your Mac. No document content is sent to cloud, AI,
  analytics or telemetry services.
- **The app is only reachable from your own Mac** (`127.0.0.1`). Other computers on the network cannot open it.
- **Results are confidential.** Extracted text, exports and downloads contain company data. Treat them like
  the original PDFs, and do not upload or email them without approval.
- **Never commit data to Git.** The repository's `.gitignore` excludes data folders, but that is not access
  control. Keep data in `~/AmiraPilotData/`.
- **Passwords are never used.** Password-protected PDFs are reported as failed and are not opened.

---

## 9. Troubleshooting

| Problem | What to do |
|---|---|
| `Python environment not found` when starting | Repeat [installation step 4](#2-one-time-installation) |
| `Port 8765 is in use` | The app is already running. Open http://127.0.0.1:8765/, or stop the other copy |
| **NAS – Not connected** | Mount the NAS in Finder (Go → Connect to Server…), then wait a few seconds |
| Catalogue is empty or says "Inventory not loaded" | Check `inventory_workbook` in `settings.toml` points to the real file, then click **Reload inventory** |
| Many files show **FOLDER_UNREADABLE** | The NAS was slow or did not answer. Click **Reload inventory** later. Files already confirmed keep their OK status |
| Titles are all blank | Check `title_workbook` in `settings.toml`. Only rows with column H filled get a title |
| A job says **waiting_for_source** | The NAS dropped. It continues automatically when the NAS is back |
| A job paused with "free disk space below min_free_gb" | Free some disk space, then click **Resume** |
| The app stops responding | Press Ctrl+C in Terminal and start it again. Finished pages are kept |
| `Encrypted PDF requires a password` | The PDF is password-protected. Ask for an unprotected copy |

When asking for help, share the **Copy sanitised report** text from the Jobs tab. Do not share document
text, page images or exports.

---

## 10. Known limitations

- **Table extraction** and **local AI (multimodal) extraction** are not implemented yet.
- **Minus signs in scanned tables:** OCR can drop or misread them (for example `-12.5` read as `12.5`). Check numbers by eye.
- **Reading order** on multi-column pages follows the PDF's internal order and is not verified.
- **Figures, maps and seismic sections** are kept as page images with their captions. They are not converted to data.
- **Large scanned documents:** OCR speed has not been measured on them yet. Expect roughly 1–3 seconds per scanned page.
- **Slow NAS:** the NAS can be slow, so the first NAS check after starting may take several minutes.

For the older command-line tool (one metadata row per PDF), see `Amira/README.md`.
