#!/usr/bin/env python3
"""Write template-shaped CSVs populated ONLY from 'Requires transformation' mappings.

Rules
- Headers copied verbatim from the template (read-only).
- Only template columns whose mapping status is 'Requires transformation' are filled;
  'Possible match — needs review' columns (Well, Reservoir, WC, Datum ...) stay blank.
- One record per (source CSV, DST test, date/zone). Competing sources and different
  tests are never merged; Test 2 and Test 2A on 12-Apr-1990 stay separate rows.
- Every output row has a provenance row (source CSV, PDF pages, test, raw values,
  exact transformation applied).

Transformations (as specified in DST_Template_Attribute_Mapping.xlsx)
- psia -> psig: psig = psia - 14.696 (ASSUMED sea-level atmosphere; confirm).
- '<value> @ <depth>' text: number before '@' used; depth kept in provenance only.
- D-13 dates: MM/DD/YY -> YYYY-MM-DD (04/09/90 -> 1990-04-09); 'APRIL 13, 90' -> 1990-04-13.
- D-5 / D-7 timestamps: sheet date (DD-MM-YY as written) + hhmm; date rolls forward when
  the clock passes midnight on a sheet whose header date was not updated.
- Daily well-test values: arithmetic mean of numeric readings per (source, test, date);
  ditto marks '↓', '-', blanks are NOT counted as values. Not time-weighted.
- Gross = oil + water for readings where both are numeric at the same timestamp.
- Table 1 'K, md': number before '(' -> md; bracketed metres -> ft (x 3.28084);
  Test 1A 'K̄, md 45, 99, 84' split by zone in printed order (E-34/36, E-27, E-26).
"""
import csv
import hashlib
import re
from collections import OrderedDict
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path("/Users/farahyusof/Documents/GitHub/NoramiraSazuki/AinnyFarah/DOCUMENT_DST")
CSV_DIR = ROOT / "table_csv"
MAP_XLSX = ROOT / "attribute_mapping" / "DST_Template_Attribute_Mapping.xlsx"
OUT_DIR = ROOT / "mapped_csv"
TEMPLATE = Path("/Users/farahyusof/Library/CloudStorage/OneDrive-MRCHURROSF&BSDN.BHD/Ainny Farah/"
                "07_STE/13_DULANG/01_EDAFY_Asset_Templates/asset_combined_templates.xlsx")
STATUS = "Requires transformation"
ATM = 14.696
M_TO_FT = 3.28084

D13 = {"1": "T15_D13_Test1_Production_Test_Summary.csv", "1A": "T22_D13_Test1A_Production_Test_Summary.csv",
       "2": "T30_D13_Test2_Production_Test_Summary.csv", "2A": "T37_D13_Test2A_Production_Test_Summary.csv"}
D5 = {"1": "T09_D5_Test1_Production_Test_Data_Sheet.csv", "1A": "T17_D5_Test1A_Production_Test_Data_Sheet.csv",
      "2": "T25_D5_Test2_Production_Test_Data_Sheet.csv", "2A": "T32_D5_Test2A_Production_Test_Data_Sheet.csv"}
D7 = {"1": "T10_D7_Test1_Oil_Rate_Calculations.csv", "1A": "T18_D7_Test1A_Oil_Rate_Calculations.csv",
      "2": "T26_D7_Test2_Oil_Rate_Calculations.csv", "2A": "T33_D7_Test2A_Oil_Rate_Calculations.csv"}
T1 = "T01_Table1_Production_Test_Summary.csv"


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def rd(name):
    return list(csv.DictReader(open(CSV_DIR / name, encoding="utf-8-sig")))


def num(s):
    """Leading number of a form value such as '1364 (Last 6 hrs)' or '1546.9 @ 1420'."""
    m = re.match(r"^\s*(-?\d[\d,]*\.?\d*)", s or "")
    return float(m.group(1).replace(",", "")) if m else None


def reading(s):
    """Strict numeric reading for time-series cells: '348.5', '348.5 *', '3150 *'.
    Ditto marks, '-', margin notes ('5 BBLS IN 3 MINUTES...') and values the writer
    marked '(?)' return None."""
    m = re.fullmatch(r"\s*(-?\d[\d,]*\.?\d*)\s*\*?\s*", s or "")
    return float(m.group(1).replace(",", "")) if m else None


# Table 1 treats Test 1A as two tests: 1A (a) on 1/2" choke and 1A (b) high-rate test on 1"
# choke, which started with the choke change at 05:51 on 10-Apr-1990 (D-5 p.72).
SPLIT_1A = datetime(1990, 4, 10, 5, 51)


def subtest(test, dt):
    if test == "1A":
        return "1A (b)" if dt >= SPLIT_1A else "1A (a)"
    return test


def d13_label(test):
    return "1A (a)" if test == "1A" else test   # D-13 Test 1A form is titled '1A (a) / 2215-0551 (1/2" CHOKE)'


def at_value(s):
    """'1546.9 @ 1420' -> (1546.9, '1420')."""
    v = num(s.split("@")[0]) if s else None
    depth = s.split("@", 1)[1].strip() if s and "@" in s else ""
    return v, depth


class _Form(dict):
    """D-13 rows by label; also matches labels with a bracketed qualifier
    (Test 1: 'Initial reservoir pressure (from Buildup)')."""
    def __missing__(self, key):
        for k, v in self.items():
            if k.startswith(key + " ("):
                return v
        raise KeyError(key)


def d13_form(test):
    return _Form({r["Field (as printed)"]: r for r in rd(D13[test])})


def d13_date(test):
    raw = d13_form(test)["Date"]["Value (as written)"]
    if re.match(r"\d\d/\d\d/\d\d$", raw):
        return datetime.strptime(raw, "%m/%d/%y").date(), raw, "MM/DD/YY -> YYYY-MM-DD"
    return datetime.strptime(raw.replace(",", ""), "%B %d %y").date(), raw, "'Month DD, YY' -> YYYY-MM-DD"


def sheet_date(raw):
    s = raw.upper().replace("TH", "").replace("(8/9)", "8")
    s = re.split(r"\s*/\s*(?=\d+-APRIL)|\s+/\s+", s)[0].strip()
    for fmt in ("%d-%m-%y", "%d/%m/%y", "%d-%B-%y", "%d-%B-%Y", "%d-%b-%Y", "%d-%b-%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    m = re.match(r"(\d+)-(\d+) APR 90", s)          # '12-13 APR 90' -> first day
    if m:
        return datetime(1990, 4, int(m.group(1))).date()
    raise ValueError(raw)


def hhmm(raw):
    d = re.sub(r"\D", "", raw or "")
    if len(d) in (3, 4):
        h, m = int(d[:-2]), int(d[-2:])
        if h <= 24 and m < 60:
            return h, m
    return None


def timestamps(rows, date_col, time_col):
    """Yield (row, datetime) with midnight roll-over when the sheet date was not updated."""
    prev = None
    for r in rows:
        t = hhmm(r[time_col])
        if t is None:
            continue
        base = sheet_date(r[date_col])
        dt = datetime.combine(base, datetime.min.time()) + timedelta(hours=t[0], minutes=t[1])
        if prev and dt < prev - timedelta(hours=1):
            dt += timedelta(days=1)
        prev = dt
        yield r, dt


def template_header(twb, sheet, row=1):
    h = list(next(twb[sheet].iter_rows(min_row=row, max_row=row, values_only=True)))
    while h and h[-1] in (None, ""):
        h.pop()
    return ["" if v is None else str(v).strip() for v in h]


def write(name, header, rows, prov, units=None):
    OUT_DIR.mkdir(exist_ok=True)
    with open(OUT_DIR / f"{name}_transformed.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        if units:
            w.writerow(units)
        w.writerows([[r.get(h, "") for h in header] for r in rows])
    with open(OUT_DIR / f"{name}_transformed_provenance.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(prov[0]))
        w.writeheader()
        for i, p in enumerate(prov, 1):
            w.writerow({**p, "Output Row": i})
    print(f"{name}_transformed.csv: {len(rows)} rows")


def prov_row(src, pages, test, populated, raw, tx, note=""):
    return OrderedDict([("Output Row", ""), ("Source CSV", src), ("Source PDF Pages", pages), ("DST Test", test),
                        ("Well (as written, not loaded)", d13_form("1")["Well"]["Value (as written)"]
                         if test.startswith("1") else d13_form("2")["Well"]["Value (as written)"]),
                        ("Populated Template Columns", populated), ("Raw Source Values", raw),
                        ("Transformation Applied", tx), ("Note", note)])


# --------------------------------------------------------------------------- per template sheet
def bhp(twb):
    header = template_header(twb, "BHP")
    rows, prov = [], []
    for test in D13:
        f = d13_form(test)
        raw = f["Final flowing pressure"]["Value (as written)"]
        v, depth = at_value(raw)
        if v is None:
            continue
        dt, draw, dtx = d13_date(test)
        rows.append({"Survey Date": dt.isoformat(), "Pwf - BHFP (psig)": f"{v - ATM:.3f}"})
        prov.append(prov_row(D13[test], f["Final flowing pressure"]["Source Page"], d13_label(test),
                             "Survey Date <- D-13 Date; Pwf - BHFP (psig) <- D-13 item 17 Final flowing pressure",
                             f"Date '{draw}'; Pwf '{raw}'", f"{dtx}; {v} psia - {ATM} = {v - ATM:.3f} psig",
                             f"Gauge depth '{depth}' not loaded (no template field). Well and Datum Pressure left blank (needs review)."))
    t1 = rd(T1)
    for r in t1:
        v, depth = at_value(r["FBHP, PSIA"])
        if v is None:
            continue
        test = r["TEST NO."].split()[0]
        dt, draw, dtx = d13_date(test)
        rows.append({"Survey Date": dt.isoformat(), "Pwf - BHFP (psig)": f"{v - ATM:.3f}"})
        prov.append(prov_row(T1, r["Source Page"], r["TEST NO."],
                             "Survey Date <- D-13 Date of the same test; Pwf - BHFP (psig) <- Table 1 FBHP, PSIA",
                             f"FBHP '{r['FBHP, PSIA']}'; date from {D13[test]} '{draw}'",
                             f"{dtx}; {v} psia - {ATM} = {v - ATM:.3f} psig",
                             f"Competing source with D-13 item 17 for the same test — kept as a separate row. Depth '{depth}' not loaded."))
    write("bhp", header, rows, prov)


def pressure_profile(twb):
    header = template_header(twb, "Pressure Profile")
    rows, prov = [], []
    for test in D13:
        f = d13_form(test)
        raw = f["Initial reservoir pressure"]["Value (as written)"]
        v, depth = at_value(raw)
        dt, draw, dtx = d13_date(test)
        if v is None:
            continue
        rows.append({"Date": dt.isoformat(), "Reservoir Pressure (psig)": f"{v - ATM:.3f}"})
        prov.append(prov_row(D13[test], f["Initial reservoir pressure"]["Source Page"], d13_label(test),
                             "Date <- D-13 Date; Reservoir Pressure (psig) <- D-13 item 16 Initial reservoir pressure",
                             f"Date '{draw}'; Pi '{raw}'", f"{dtx}; {v} psia - {ATM} = {v - ATM:.3f} psig",
                             f"Reference depth '{depth}' not loaded (no template field). Well left blank (needs review)."))
    for r in rd(T1):
        v, depth = at_value(r["Pi, PSIA"])
        if v is None:
            continue
        test = r["TEST NO."].split()[0]
        dt, draw, dtx = d13_date(test)
        rows.append({"Date": dt.isoformat(), "Reservoir Pressure (psig)": f"{v - ATM:.3f}"})
        prov.append(prov_row(T1, r["Source Page"], r["TEST NO."],
                             "Date <- D-13 Date of the same test; Reservoir Pressure (psig) <- Table 1 Pi, PSIA",
                             f"Pi '{r['Pi, PSIA']}'; date from {D13[test]} '{draw}'",
                             f"{dtx}; {v} psia - {ATM} = {v - ATM:.3f} psig",
                             f"Competing source with D-13 item 16 — separate row. Depth '{depth}' not loaded."))
    write("pressure_profile", header, rows, prov)


def well_reservoir_props(twb):
    header = template_header(twb, "Well Reservoir Props")
    rows, prov = [], []
    for r in rd(T1):
        raw, zone, test = r["K, md"], r["ZONE"], r["TEST NO."]
        if not raw or raw == "-":
            continue
        multi = re.match(r"K̄, md\s+([\d.,\s]+)$", raw)
        if multi:
            ks = [k.strip() for k in multi.group(1).split(",")]
            zones = [z.strip() for z in zone.split("+")]
            assert len(ks) == len(zones)
            for k, z in zip(ks, zones):
                rows.append({"Permeability (md)": k})
                prov.append(prov_row(T1, r["Source Page"], test, "Permeability (md) <- Table 1 K, md (split by zone)",
                                     f"K '{raw}'; ZONE '{zone}'", f"value {k} assigned to zone {z} by printed order",
                                     f"Zone {z} (as written) not loaded — Reservoir needs review. Net pay not given for this test. "
                                     "D-13 p.82 remarks agree: E-34/36 45 md, E-27 99 md, E-26 84 md."))
            continue
        m = re.match(r"^\s*(\d+\.?\d*)\*?\s*\(([\d.]+)\s*m\)", raw)
        k, pay_m = m.group(1), float(m.group(2))
        rows.append({"Permeability (md)": k, "Net Pay (ft)": f"{pay_m * M_TO_FT:.2f}"})
        star = "*" in raw
        prov.append(prov_row(T1, r["Source Page"], test,
                             "Permeability (md) <- Table 1 K, md; Net Pay (ft) <- bracketed metres in K, md",
                             f"K '{raw}'; ZONE '{zone}'",
                             f"K = {k} md; net pay {pay_m} m x {M_TO_FT} = {pay_m * M_TO_FT:.2f} ft",
                             (f"Zone {zone} not loaded (Reservoir needs review). "
                              + ("'*' = average permeability based on PSS flow calculation (Table 1 footnote)." if star else ""))))
    write("well_reservoir_props", header, rows, prov)


def well_test_monitoring(twb):
    # Template layout: row 6 'No','Date' ; row 8 block labels ; row 9 block units; well name in block header (row 7).
    r6 = template_header(twb, "Well Test Monitoring", 6)
    r8 = template_header(twb, "Well Test Monitoring", 8)
    r9 = list(next(twb["Well Test Monitoring"].iter_rows(min_row=9, max_row=9, max_col=12, values_only=True)))
    header = [r6[0], r6[1]] + r8[2:12]
    units = ["", ""] + ["" if u is None else str(u) for u in r9[2:12]]
    rows, prov = [], []

    def daily(src, test, date_col, time_col, oil_col, water_col=None):
        groups = OrderedDict()
        seen = {}
        for r, dt in timestamps(rd(src), date_col, time_col):
            o = reading(r[oil_col])
            # Same timestamp recorded twice with a value (e.g. Test 2A 06:30 on p.115 and the
            # separate one-row sheet p.116): keep the first, report the duplicate.
            if o is not None and dt in seen:
                groups[(subtest(test, dt), dt.date())]["dup"].append(
                    f"{dt:%H:%M} p.{r['Source Page']} ({r[oil_col]}) duplicate of p.{seen[dt]}")
                continue
            if o is not None:
                seen[dt] = r["Source Page"]
            w = reading(r[water_col]) if water_col else None
            skipped = [r[c] for c in (oil_col, water_col) if c and r[c] and reading(r[c]) is None
                       and r[c].strip() not in ("↓", "-", '"')]
            key = (subtest(test, dt), dt.date())
            g = groups.setdefault(key, {"oil": [], "water": [], "gross": [], "pages": set(), "star": 0, "skip": [], "dup": []})
            g["skip"] += skipped
            if o is None and w is None:
                continue
            g["pages"].add(int(r["Source Page"]))
            if o is not None:
                g["oil"].append(o)
                g["star"] += "*" in r[oil_col]
            if w is not None:
                g["water"].append(w)
            if o is not None and w is not None:
                g["gross"].append(o + w)
        return groups

    mean = lambda xs: f"{sum(xs) / len(xs):.2f}" if xs else ""
    for test in D5:
        for src, oc, wc, label in [(D5[test], "10 RATES OIL STB/D", "11 RATES WATER B/D", "D-5"),
                                   (D7[test], "12 CORRECTED RATE STB/D", None, "D-7")]:
            for (sub, day), g in daily(src, test, "Sheet Date (as written)", "1 DATE TIME", oc, wc).items():
                if not g["oil"] and not g["water"]:
                    continue
                row = {header[1]: day.isoformat(), "Oil": mean(g["oil"])}
                if label == "D-5":
                    row["Water"] = mean(g["water"])
                    row["Gross"] = mean(g["gross"])
                rows.append(row)
                pages = sorted(g["pages"])
                prov.append(prov_row(src, f"{pages[0]}-{pages[-1]}" if len(pages) > 1 else str(pages[0]), sub,
                                     "Date <- sheet date + time; Oil <- " + oc + (
                                         "; Water <- " + wc + "; Gross <- oil + water" if wc else ""),
                                     f"{len(g['oil'])} oil readings, {len(g['water'])} water readings, "
                                     f"{len(g['gross'])} paired oil+water readings",
                                     "daily arithmetic mean of numeric readings (STB/D -> BOPD 1:1; bbl/d -> BWPD 1:1); "
                                     "Gross = mean of (oil + water) over paired readings only",
                                     (f"{g['star']} oil value(s) carry '*' (estimated / meter suspect). " if g["star"] else "")
                                     + (f"Excluded non-numeric cells: {'; '.join(x[:40] for x in g['skip'])}. " if g["skip"] else "")
                                     + (f"Excluded duplicate timestamp: {'; '.join(g['dup'])}. " if g["dup"] else "")
                                     + ("Gross blank: no paired water readings. " if label == "D-5" and not g["gross"] else "")
                                     + "Well block header and WC/Freq/Ampere/Voltage/Pi/Tm/RH left blank."))
        f = d13_form(test)
        raw = f["Average rate"]["Value (as written)"]
        v = num(raw)
        dt, draw, dtx = d13_date(test)
        rows.append({header[1]: dt.isoformat(), "Oil": f"{v:g}"})
        prov.append(prov_row(D13[test], f["Average rate"]["Source Page"], d13_label(test),
                             "Date <- D-13 Date; Oil <- D-13 item 4 Average rate",
                             f"Avg rate '{raw}'; date '{draw}'", f"{dtx}; STB/D -> BOPD 1:1; bracketed note stripped",
                             "Test-average rate (not a daily average) — competes with D-5/D-7 daily rows; separate row."))
    write("well_test_monitoring", header, rows, prov, units)


def main():
    before = {p: sha(p) for p in [TEMPLATE, MAP_XLSX, *CSV_DIR.glob("*.csv")]}
    # Confirm every column we fill is a 'Requires transformation' mapping in the workbook.
    m = [r for r in load_workbook(MAP_XLSX, read_only=True).worksheets[0].iter_rows(min_row=2, values_only=True)]
    allowed = {(r[0], r[4], r[5]) for r in m if r[7] == STATUS}
    for need in [(D13["1"], "BHP", "Pwf - BHFP (psig)"), (T1, "BHP", "Pwf - BHFP (psig)"),
                 (D13["1"], "BHP", "Survey Date"), (D13["1"], "Pressure Profile", "Reservoir Pressure (psig)"),
                 (T1, "Pressure Profile", "Reservoir Pressure (psig)"), (D13["1"], "Pressure Profile", "Date"),
                 (T1, "Well Reservoir Props", "Permeability (md)"), (T1, "Well Reservoir Props", "Net Pay (ft)"),
                 (D5["1"], "Well Test Monitoring", "Oil (BOPD)"), (D7["1"], "Well Test Monitoring", "Oil (BOPD)"),
                 (D13["1"], "Well Test Monitoring", "Oil (BOPD)"), (D5["1"], "Well Test Monitoring", "Water (BWPD)"),
                 (D5["1"], "Well Test Monitoring", "Gross (BFPD)"), (D5["1"], "Well Test Monitoring", "Date")]:
        assert need in allowed, f"not a '{STATUS}' mapping: {need}"

    twb = load_workbook(TEMPLATE, read_only=True, data_only=True)
    bhp(twb)
    pressure_profile(twb)
    well_reservoir_props(twb)
    well_test_monitoring(twb)
    twb.close()
    assert all(sha(p) == h for p, h in before.items()), "an input file changed"
    print("inputs unchanged (sha256 verified)")


if __name__ == "__main__":
    main()
