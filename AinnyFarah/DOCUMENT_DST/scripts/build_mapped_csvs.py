#!/usr/bin/env python3
"""Write template-shaped CSVs containing ONLY 'Direct 1-to-1' mappings.

- Headers are copied verbatim from the template sheet (read-only).
- One output row per source record (source CSV + test); records are never merged.
- Template columns without a Direct 1-to-1 mapping are left blank (well, reservoir,
  dates etc. are 'needs review' / 'requires transformation' and are NOT filled).
- A provenance sidecar links each output row to its source CSV, PDF page, test,
  and the well/zone/date exactly as written in the report, without putting those
  unconfirmed values into the template columns.
"""
import csv
import hashlib
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path("/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/DOCUMENT_DST")
CSV_DIR = ROOT / "table_csv"
MAP_XLSX = ROOT / "attribute_mapping" / "DST_Template_Attribute_Mapping.xlsx"
OUT_DIR = ROOT / "mapped_csv"
TEMPLATE = Path("/Users/farahyusof/Library/CloudStorage/OneDrive-MRCHURROSF&BSDN.BHD/Ainny Farah/"
                "07_STE/13_DULANG/01_EDAFY_Asset_Templates/asset_combined_templates.xlsx")
DIRECT = "Direct 1-to-1"


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def read_csv(p):
    return list(csv.DictReader(open(p, encoding="utf-8-sig")))


def main():
    before = {p: sha(p) for p in [TEMPLATE, MAP_XLSX, *CSV_DIR.glob("*.csv")]}

    mapping = [r for r in load_workbook(MAP_XLSX, read_only=True).worksheets[0].iter_rows(min_row=2, values_only=True)
               if r[7] == DIRECT]
    # group: template sheet -> source csv -> {template attr: source attr}
    plan = {}
    for src_csv, _tbl, src_attr, _u, sheet, tattr, *_ in mapping:
        plan.setdefault(sheet, {}).setdefault(src_csv, {})[tattr] = src_attr

    twb = load_workbook(TEMPLATE, read_only=True, data_only=True)
    index = {r["CSV File"]: r for r in read_csv(CSV_DIR / "_table_index.csv")}
    OUT_DIR.mkdir(exist_ok=True)

    for sheet, sources in plan.items():
        header = [h for h in next(twb[sheet].iter_rows(min_row=1, max_row=1, values_only=True))]
        while header and header[-1] in (None, ""):
            header.pop()
        header = ["" if h is None else str(h) for h in header]
        for t in {t for m in sources.values() for t in m}:
            assert t in header, f"{t} not an exact header of {sheet}"

        rows, prov = [], []
        for src_csv, attrs in sorted(sources.items()):
            src_rows = read_csv(CSV_DIR / src_csv)
            # form-style CSV: one attribute per row ('Parameter' / 'Value')
            by_param = {r["Parameter"]: r for r in src_rows}
            out = {h: "" for h in header}
            pages = set()
            for tattr, sattr in attrs.items():
                r = by_param[sattr]
                out[tattr] = r["Value"]
                pages.add(r["Source Page"])
            rows.append([out[h] for h in header])

            idx = index[src_csv]
            # Context as written in the report (not loaded into template columns).
            t1 = {r["TEST NO."]: r for r in read_csv(CSV_DIR / "T01_Table1_Production_Test_Summary.csv")}
            d13 = {r["Field (as printed)"]: r["Value (as written)"] for r in
                   read_csv(CSV_DIR / "T15_D13_Test1_Production_Test_Summary.csv")}
            prov.append({
                "Output Row": len(rows),
                "Source CSV": src_csv,
                "Source Table (as printed)": idx["Table Title (as printed)"],
                "Source PDF Pages": ", ".join(sorted(pages)),
                "DST Test": idx["Test"],
                "Well (as written, not loaded)": d13.get("Well", ""),
                "Zone (as written, not loaded)": t1["1"]["ZONE"],
                "Perforation Interval (as written)": t1["1"]["INTERVAL M, RKB"],
                "Test Date (as written D-13, MM/DD/YY)": d13.get("Date", ""),
                "Populated Template Columns": "; ".join(f"{t} <- {s}" for t, s in attrs.items()),
                "Blank Template Columns (no Direct 1-to-1 match)": "; ".join(h for h in header if h not in attrs),
                "Note": ("Table 3 (Horner) and Table 4 (type-curve) are separate analyses of the same Test 1 record; "
                         "kept as separate rows, not merged. KH differs (1975.068 vs 2179.739 md·ft) — choose one before upload."),
            })

        name = sheet.lower().replace(" ", "_")
        with open(OUT_DIR / f"{name}_mapped.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(header)
            w.writerows(rows)
        with open(OUT_DIR / f"{name}_mapped_provenance.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(prov[0]))
            w.writeheader()
            w.writerows(prov)
        print(f"{sheet}: {len(rows)} rows -> {name}_mapped.csv (+ provenance)")

    twb.close()
    assert all(sha(p) == h for p, h in before.items()), "an input file changed"
    print("Sheets with no Direct 1-to-1 match (no CSV written):",
          [s for s in twb.sheetnames if s not in plan])
    print("inputs unchanged (sha256 verified)")


if __name__ == "__main__":
    main()
