#!/usr/bin/env python3
"""Standardise the first seven columns of the Well Intervention upload sheets into the template order
ORIGINAL WELL NAME | FIELD | PLATFORM | WELL | STRING TYPE | RESERVOIR | DATE, keeping every other column after them.

Source (read only): all_data_standardisation/Well-Interventions/Well_Intervention_*.xlsx
Output:             all_data_standardisation/01_UPDATED/WELL-INTERVENTIONS/<same name>  (+ backup_original/, _reports/)

Columns are moved by meaning: old A-E stay A-E, RESERVOIR is new F, Start Date (old H) -> G, then Entity Name, Well Type,
Objective, Job Summary, Detail Job, Well Completion Schematic (, Source Reference) in their original order.
  ORIGINAL WELL NAME  column A as is
  FIELD               DULANG
  PLATFORM            DULANG-<letter> from the WELL identifier (two-well names: from the name if both wells agree)
  WELL                existing Well value: DULG-/DULANG- -> DULANG-, leading zeros dropped from the main number and every
                      numeric suffix (DULG-A001W01 -> DULANG-A1W1, DULANG-D12ST1 kept); blank for names covering two wells
  STRING TYPE         kept only where stated: in the original name (L / S / S&L ...), on Well Intervention History rows
                      (WellString suffix), or C rows built by rule b (source string column); otherwise blank
  RESERVOIR           no reservoir field in these reports -> blank
  DATE                Start Date, unchanged
Package-level edit: only Intervention_Upload's sheet XML, sharedStrings (appended) and the workbook's filter name change;
every other part is copied byte for byte. The result is validated cell by cell against the source.
"""
import collections
import html
import re
import shutil
import sys
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string as cidx, get_column_letter as cl

BASE = Path(__file__).resolve().parent / "all_data_standardisation"
SRC = BASE / "Well-Interventions"
OUT = BASE / "01_UPDATED" / "WELL-INTERVENTIONS"
BACKUP = OUT / "backup_original"
REPORTS = OUT / "_reports"
SHEET = "Intervention_Upload"
OLD_HDR = ["Original Well Name", "Field", "Platform", "Well", "String Type", "Entity Name", "Well Type", "Start Date"]
NEW_HDR = ["ORIGINAL WELL NAME", "FIELD", "PLATFORM", "WELL", "STRING TYPE", "RESERVOIR", "DATE"]
WELL_RE = re.compile(r"(?:DULG|DULANG)-([A-D])(\d+)((?:[A-Z]+\d*)*)")
STATED_RE = re.compile(r"\d\s*(S\s*&\s*L|S\s*/\s*L|L\s*&\s*S|L|S)(?![A-Z0-9])")
TWO_WELLS_RE = re.compile(r"&\s*[A-D]-?\d")


def sst_texts(xml):
    return [html.unescape("".join(t.decode() for t in re.findall(rb"<t[^>]*>(.*?)</t>", it, re.S)))
            for it in re.findall(rb"<si>(.*?)</si>", xml, re.S)]


def colmap(ncols):
    """old 1-based column -> new 1-based column (new column 6 = RESERVOIR)"""
    m = {1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 8, 7: 9, 8: 7}
    m.update({c: c + 1 for c in range(9, ncols + 1)})
    return m


def std_well(v):
    m = WELL_RE.fullmatch(v)
    if not m:
        return None
    sfx = re.sub(r"([A-Z]+)(\d*)", lambda s: s.group(1) + (str(int(s.group(2))) if s.group(2) else ""), m.group(3))
    return f"DULANG-{m.group(1)}{int(m.group(2))}{sfx}"


def stated_strings(name):
    """string types written in the original name, or None"""
    m = STATED_RE.search(name.upper())
    if not m:
        return None
    t = re.sub(r"\s", "", m.group(1))
    return {"LS", "SS"} if len(t) > 1 else {"L": {"LS"}, "S": {"SS"}}[t]


def decide(r, plat):
    """values for the seven template columns of one data row + review notes"""
    orig, well, st, ent = r[0], r[3], r[4], r[5]
    src_ref = r[12] if len(r) > 12 else None
    notes = []
    o = orig.strip()
    two = bool(TWO_WELLS_RE.search(o.upper()))
    # WELL / PLATFORM
    if two:
        new_well = None
        letters = set(re.findall(r"(?:DULANG\s*|&\s*)([A-D])-?\d", o.upper()))
        platform = f"DULANG-{letters.pop()}" if len(letters) == 1 else None
        notes.append(("WELL blank", "original name covers two wells"))
    else:
        new_well = std_well(well) if isinstance(well, str) else None
        if new_well is None:
            platform = None
            notes.append(("WELL blank", f"Well value {well!r} not in a recognised format"))
        else:
            platform = f"DULANG-{new_well[7]}"
            if new_well[7] != plat:
                notes.append(("review", f"WELL platform {new_well[7]} differs from the file's platform {plat}"))
            mo = re.search(r"(?:DULANG[\s-]*)([A-D])?\s*-?\s*0*(\d+)", o.upper())
            if not mo:
                notes.append(("review", "no well number in original name"))
            else:
                if mo.group(1) is None:
                    notes.append(("review", f"original name has no platform letter; {new_well[7]} taken from the WELL value"))
                elif mo.group(1) != new_well[7]:
                    notes.append(("review", f"platform letter in original name ({mo.group(1)}) differs from WELL"))
                if int(mo.group(2)) != int(WELL_RE.fullmatch(well).group(2)):
                    notes.append(("review", "well number in original name differs from WELL"))
    # STRING TYPE
    new_st = None
    if st is not None:
        named = None if two else stated_strings(o)
        if src_ref is not None:
            wih = str(src_ref).startswith("Well Intervention History")
            rule = re.search(r"String Type rule (\w)", str(src_ref))
            rule = rule.group(1) if rule else None
        else:
            wih = bool(re.match(r"DULANG-[A-D]\d", o))
            rule = None
        if named is not None:
            if st in named:
                new_st, basis = st, "stated in original name"
            else:
                notes.append(("STRING blank", f"{st} conflicts with original name"))
        elif wih:
            if isinstance(ent, str) and ent.endswith("-" + st):
                new_st, basis = st, "WIH WellString"
            else:
                notes.append(("STRING blank", f"{st}: WIH row but Entity Name {ent!r} does not carry it"))
        elif rule == "b":
            new_st, basis = st, "rule b (source string column)"
        else:
            how = f"rule {rule} (inferred)" if rule else "basis not verifiable"
            notes.append(("STRING blank", f"{st} removed: {how}"))
    return dict(field="DULANG", platform=platform, well=new_well, string=new_st), notes


CELL_RE = re.compile(rb'<c r="([A-Z]+)(\d+)"((?:\s+[\w:]+="[^"]*")*)\s*(?:/>|>(.*?)</c>)', re.S)
ROW_RE = re.compile(rb'(<row r="(\d+)"[^>]*?)(?:/>|>(.*?)</row>)', re.S)


def remap_range(ref, m):
    """'B1:L3941' -> range covering the mapped columns"""
    def one(a):
        mm = re.fullmatch(r"(\$?)([A-Z]+)(\$?)(\d+)", a)
        return mm.group(1), cidx(mm.group(2)), mm.group(3), mm.group(4)
    parts = ref.split(":")
    if len(parts) == 1:
        d1, c, d2, r = one(parts[0])
        return f"{d1}{cl(m[c])}{d2}{r}"
    (a1, c1, a2, r1), (b1, c2, b2, r2) = one(parts[0]), one(parts[1])
    cols = [m[c] for c in range(c1, c2 + 1)]
    return f"{a1}{cl(min(cols))}{a2}{r1}:{b1}{cl(max(cols))}{b2}{r2}"


def build(src, dst):
    plat = re.search(r"Dulang_([A-D])_", src.name).group(1)
    wb = openpyxl.load_workbook(src, read_only=True)
    rows = list(wb[SHEET].iter_rows(values_only=True))
    hdr = list(rows[0])
    assert hdr[:8] == OLD_HDR, hdr
    ncols = len(hdr)
    m = colmap(ncols)
    decisions, review = {}, []
    for i, r in enumerate(rows[1:], 2):
        if all(v is None for v in r):
            continue
        d, notes = decide(r, plat)
        decisions[i] = d
        for kind, why in notes:
            review.append((src.name, i, r[0], r[3], r[4], kind, why))

    zin = zipfile.ZipFile(src)
    sheet_part = "xl/worksheets/sheet1.xml"
    assert re.search(r'<sheet name="%s" sheetId="\d+" r:id="rId1"' % SHEET, zin.read("xl/workbook.xml").decode())
    sst_xml = zin.read("xl/sharedStrings.xml")
    nsst = len(sst_texts(sst_xml))
    added = {}

    def sidx(text):
        if text not in added:
            added[text] = nsst + len(added)
        return added[text]

    def cell(col, row, attrs, inner=None, text=None, blank=False):
        a = re.sub(r'\st="[^"]+"', "", attrs) if (text is not None or blank) else attrs
        ref = f"{cl(col)}{row}"
        if text is not None:
            return f'<c r="{ref}"{a} t="s"><v>{sidx(text)}</v></c>'
        if blank or inner is None:
            return f'<c r="{ref}"{a}/>'
        return f'<c r="{ref}"{a}>{inner}</c>'

    def row_sub(mo):
        head, rn, body = mo.group(1).decode(), int(mo.group(2)), (mo.group(3) or b"")
        head = re.sub(r'spans="1:\d+"', f'spans="1:{ncols + 1}"', head)
        cells = {cidx(c.group(1).decode()): (c.group(3).decode(), None if c.group(4) is None else c.group(4).decode())
                 for c in CELL_RE.finditer(body)}
        assert all(int(c.group(2)) == rn for c in CELL_RE.finditer(body))
        out = {}
        for oc, (attrs, inner) in cells.items():
            out[m[oc]] = cell(m[oc], rn, attrs, inner)
        if rn == 1:
            for c, name in enumerate(NEW_HDR, 1):
                src_c = {6: 5, 7: 8}.get(c, c)
                out[c] = cell(c, rn, cells[src_c][0], text=name)
        elif rn in decisions:
            d = decisions[rn]
            for c, key in ((2, "field"), (3, "platform"), (4, "well"), (5, "string")):
                attrs = cells.get(c, ("", None))[0]
                out[c] = cell(c, rn, attrs, text=d[key]) if d[key] is not None else cell(c, rn, attrs, blank=True)
            out[6] = cell(6, rn, cells.get(5, ("", None))[0], blank=True)
        body = "".join(out[c] for c in sorted(out))
        return f"{head}>{body}</row>".encode()

    xml = zin.read(sheet_part)
    pre, rest = xml.split(b"<sheetData>", 1)
    data, post = rest.split(b"</sheetData>", 1)
    data = ROW_RE.sub(row_sub, data)
    pre, post = pre.decode(), post.decode()
    # dimension, column widths
    pre = re.sub(r'<dimension ref="([^"]+)"/>', lambda x: f'<dimension ref="{remap_range(x.group(1), m)}"/>', pre)
    cols = {}
    for c in re.findall(r"<col [^>]*/>", pre):
        lo, hi = int(re.search(r'min="(\d+)"', c).group(1)), int(re.search(r'max="(\d+)"', c).group(1))
        for k in range(lo, hi + 1):
            cols[m[k]] = re.sub(r'min="\d+" max="\d+"', f'min="{m[k]}" max="{m[k]}"', c)
    cols[6] = '<col min="6" max="6" width="12" customWidth="1"/>'
    pre = re.sub(r"<cols>.*?</cols>", "<cols>" + "".join(cols[k] for k in sorted(cols)) + "</cols>", pre, flags=re.S)
    # filter, sort, validation refs
    af_start = None

    def ref_sub(x):
        return f'{x.group(1)}="{" ".join(remap_range(p, m) for p in x.group(2).split())}"'
    af = re.search(r'<autoFilter ref="([A-Z]+)\d', post)
    af_start = cidx(af.group(1)) if af else None
    post = re.sub(r'\b(ref|sqref)="([^"]+)"', ref_sub, post)
    if af:
        new_start = cidx(re.search(r'<autoFilter ref="([A-Z]+)\d', post).group(1))
        post = re.sub(r'<filterColumn colId="(\d+)"',
                      lambda x: f'<filterColumn colId="{m[af_start + int(x.group(1))] - new_start}"', post)
    new_sheet = (pre + "<sheetData>").encode() + data + ("</sheetData>" + post).encode()

    wbx = zin.read("xl/workbook.xml").decode()
    wbx = re.sub(r"(%s!)(\$[A-Z]+\$\d+:\$[A-Z]+\$\d+)" % SHEET, lambda x: x.group(1) + remap_range(x.group(2), m), wbx)
    new_sst = sst_xml.replace(b"</sst>", b"".join(b"<si><t>%s</t></si>" % html.escape(t, quote=False).encode()
                                                for t in added) + b"</sst>")
    new_sst = re.sub(rb'uniqueCount="\d+"', b'uniqueCount="%d"' % (nsst + len(added)), new_sst, count=1)
    parts = {sheet_part: new_sheet, "xl/sharedStrings.xml": new_sst, "xl/workbook.xml": wbx.encode()}

    tmp = dst.with_name(dst.stem + ".partial.xlsx")
    with open(tmp, "xb") as fh, zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zi = zipfile.ZipInfo(info.filename, info.date_time)
            zi.compress_type, zi.external_attr = zipfile.ZIP_DEFLATED, info.external_attr
            zout.writestr(zi, parts.get(info.filename, zin.read(info.filename)))

    validate(src, tmp, m, decisions)
    tmp.rename(dst)
    return decisions, review, len(added)


def validate(src, tmp, m, decisions):
    zs, zt = zipfile.ZipFile(src), zipfile.ZipFile(tmp)
    assert zs.namelist() == zt.namelist()
    for n in zs.namelist():
        if n not in ("xl/worksheets/sheet1.xml", "xl/sharedStrings.xml", "xl/workbook.xml"):
            assert zs.read(n) == zt.read(n), n
    a, b = openpyxl.load_workbook(src), openpyxl.load_workbook(tmp)
    assert a.sheetnames == b.sheetnames
    for name in a.sheetnames:
        if name == SHEET:
            continue
        assert [[c.value for c in r] for r in a[name].iter_rows()] == [[c.value for c in r] for r in b[name].iter_rows()]
    wa, wb_ = a[SHEET], b[SHEET]
    assert wa.max_row == wb_.max_row and wb_.max_column == wa.max_column + 1
    for r in range(1, wa.max_row + 1):
        ra, rb = wa.row_dimensions[r], wb_.row_dimensions[r]
        assert (ra.hidden, ra.height) == (rb.hidden, rb.height), r
        for oc in range(1, wa.max_column + 1):
            x, y = wa.cell(r, oc), wb_.cell(r, m[oc])
            assert x.style_id == y.style_id and x.number_format == y.number_format, (r, oc)
            if r == 1:
                continue
            if oc in (1, 6, 7) or oc >= 8:
                assert x.value == y.value, (r, oc, x.value, y.value)
        f = wb_.cell(r, 6)
        assert f.value == (NEW_HDR[5] if r == 1 else None) and f.style_id == wa.cell(r, 5).style_id, r
        if r == 1:
            assert [wb_.cell(1, c).value for c in range(1, 8)] == NEW_HDR
        elif r in decisions:
            d = decisions[r]
            assert [wb_.cell(r, c).value for c in (2, 3, 4, 5)] == [d["field"], d["platform"], d["well"], d["string"]]
    for k, dim in wa.column_dimensions.items():
        if dim.width:
            assert wb_.column_dimensions[cl(m[cidx(k)])].width == dim.width, k
    dva = [(d.formula1, str(d.sqref)) for d in wa.data_validations.dataValidation]
    dvb = [(d.formula1, str(d.sqref)) for d in wb_.data_validations.dataValidation]
    assert [f for f, _ in dva] == [f for f, _ in dvb] and all(s.startswith("I2:I") for _, s in dvb), dvb
    zs.close()
    zt.close()


def main():
    files = sorted(p for p in SRC.glob("Well_Intervention_*.xlsx") if not p.name.startswith("~$"))
    OUT.mkdir(parents=True, exist_ok=True)
    for p in files:
        for d in (OUT / p.name, BACKUP / p.name):
            if d.exists():
                sys.exit(f"STOP: {d} already exists - not overwritten")
    BACKUP.mkdir(exist_ok=True)
    REPORTS.mkdir(exist_ok=True)
    all_review, summary = [], []
    for p in files:
        shutil.copy2(p, BACKUP / p.name)
        decisions, review, nadd = build(BACKUP / p.name, OUT / p.name)
        all_review += review
        kinds = collections.Counter(k for *_, k, _ in review)
        st = collections.Counter(d["string"] for d in decisions.values())
        summary.append((p.name, len(decisions), dict(kinds), dict(st),
                        sum(1 for d in decisions.values() if d["well"] is None),
                        len({d["well"] for d in decisions.values() if d["well"]})))
        print(f"== {p.name}: {len(decisions)} rows; review {dict(kinds)}; STRING TYPE {dict(st)}")
    out = openpyxl.Workbook()
    ws = out.active
    ws.title = "Unresolved"
    ws.append(["File", "Row", "Original Well Name", "Well (source)", "String Type (source)", "Issue", "Detail"])
    for r in all_review:
        ws.append(list(r))
    ws2 = out.create_sheet("Summary")
    ws2.append(["File", "Data rows", "Issues by kind", "STRING TYPE after update", "WELL blank", "Distinct WELL"])
    for s in summary:
        ws2.append([s[0], s[1], str(s[2]), str(s[3]), s[4], s[5]])
    rp = REPORTS / "well_intervention_standardisation_review.xlsx"
    out.save(rp)
    print("review list:", rp)


if __name__ == "__main__":
    main()
