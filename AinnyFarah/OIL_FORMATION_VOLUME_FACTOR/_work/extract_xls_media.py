"""Extract embedded pictures (MSODRAWINGGROUP blips) and text-box text (TXO) from a BIFF8 .xls, per sheet."""
import struct, sys, zlib, olefile, json
from pathlib import Path
src, out = Path(sys.argv[1]), Path(sys.argv[2])
data = olefile.OleFileIO(str(src)).openstream("Workbook").read()
recs, i = [], 0
while i + 4 <= len(data):
    t, n = struct.unpack_from("<HH", data, i); recs.append((t, data[i+4:i+4+n])); i += 4 + n
# sheet names from BOUNDSHEET
sheets = []
for t, b in recs:
    if t == 0x85:
        ln, fl = b[6], b[7]
        sheets.append(b[8:8+ln*(2 if fl&1 else 1)].decode("utf-16le" if fl&1 else "latin-1"))
grp, cur, k = b"", None, 0
texts, sheet_idx, bof = [], -1, 0
j = 0
while j < len(recs):
    t, b = recs[j]
    if t == 0x809:
        bof += 1; sheet_idx = bof - 2
    if t == 0xEB:
        grp += b
        while j + 1 < len(recs) and recs[j+1][0] == 0x3C:
            j += 1; grp += recs[j][1]
    if t == 0x1B6:  # TXO: text in following CONTINUE records
        cch = struct.unpack_from("<H", b, 10)[0]
        s, got = "", 0
        jj = j + 1
        while jj < len(recs) and recs[jj][0] == 0x3C and got < cch:
            c = recs[jj][1]; fl = c[0]; body = c[1:]
            part = body.decode("utf-16le", "replace") if fl & 1 else body.decode("latin-1")
            s += part; got += len(part); jj += 1
        texts.append(dict(sheet=sheets[sheet_idx] if 0 <= sheet_idx < len(sheets) else sheet_idx, text=s[:cch]))
    j += 1
# blips inside BStore: walk OfficeArt records
def walk(buf, off, end, found):
    while off + 8 <= end:
        verinst, rt, ln = struct.unpack_from("<HHI", buf, off)
        ver = verinst & 0xF
        if ver == 0xF:
            walk(buf, off + 8, off + 8 + ln, found)
        elif 0xF018 <= rt <= 0xF117:
            found.append((rt, verinst >> 4, buf[off+8:off+8+ln]))
        elif rt == 0xF007:  # FBSE: 36-byte header then embedded blip
            sub = off + 8 + 36
            if sub < off + 8 + ln: walk(buf, sub, off + 8 + ln, found)
        off += 8 + ln
found = []
walk(grp, 0, len(grp), found)
n = 0
for rt, inst, body in found:
    n += 1
    if rt in (0xF01D, 0xF01E, 0xF01F, 0xF020, 0xF021, 0xF029, 0xF02A):  # JPEG, PNG, DIB, TIFF
        hdr = 17 + (16 if inst in (0x46B, 0x6E1, 0x6E3, 0x6E5, 0x7A9, 0x6E1) else 0)
        raw = body[hdr:]
        for sig, ext in ((b"\x89PNG", "png"), (b"\xff\xd8", "jpg")):
            p = raw.find(sig)
            if 0 <= p < 64:
                (out / f"blip{n:02d}.{ext}").write_bytes(raw[p:]); break
        else:
            (out / f"blip{n:02d}.bin").write_bytes(raw)
    elif rt in (0xF01A, 0xF01B, 0xF01C):  # EMF, WMF, PICT (zlib, 34-byte metafile header after UID(s))
        for hdr in (50, 66):
            try:
                raw = zlib.decompress(body[hdr:]); break
            except Exception:
                raw = None
        ext = {0xF01A: "emf", 0xF01B: "wmf", 0xF01C: "pict"}[rt]
        (out / f"blip{n:02d}.{ext}").write_bytes(raw if raw else body)
print("sheets", sheets)
print("blips", [(hex(rt), len(b)) for rt, _, b in found])
(out / "textboxes.json").write_text(json.dumps(texts, indent=1, ensure_ascii=False))
for t in texts: print(t)
