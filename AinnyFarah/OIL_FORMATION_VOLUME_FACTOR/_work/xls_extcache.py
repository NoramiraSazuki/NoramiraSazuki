"""Decode BIFF8 external-reference caches: SUPBOOK (0x1AE) -> XCT (0x59) sheet blocks -> CRN (0x5A) cached cells."""
import olefile, struct, sys, json
d = olefile.OleFileIO(sys.argv[1]).openstream("Workbook").read()
recs = []; i = 0
while i + 4 <= len(d):
    t, n = struct.unpack_from("<HH", d, i); recs.append((t, d[i+4:i+4+n])); i += 4 + n
def ustr(b, o, cch=None, lenbytes=2):
    if cch is None:
        cch = struct.unpack_from("<H" if lenbytes == 2 else "<B", b, o)[0]; o += lenbytes
    fl = b[o]; o += 1
    if fl & 1: s = b[o:o+2*cch].decode("utf-16le", "replace"); o += 2*cch
    else: s = b[o:o+cch].decode("latin-1"); o += cch
    return s, o
out = []; book = None; sheets = []; cur = None; in_globals = True; bof = 0
for k, (t, b) in enumerate(recs):
    if t == 0x809:
        bof += 1
    if t == 0x0A and bof == 1: break  # end of globals
    if t == 0x1AE:
        ctab, cch = struct.unpack_from("<HH", b, 0)
        if cch in (0x401, 0x3A01): book = dict(name="<self/addin>", sheets=[], cells=[]); out.append(book); continue
        name, o = ustr(b, 2)
        sh = []
        for _ in range(ctab):
            s, o = ustr(b, o); sh.append(s)
        book = dict(name=name, sheets=sh, cells=[]); out.append(book)
    elif t == 0x59 and book is not None:
        ncrn, itab = struct.unpack_from("<hH", b, 0); cur = book["sheets"][itab] if itab < len(book["sheets"]) else itab
    elif t == 0x5A and book is not None:
        colL, colF, row = struct.unpack_from("<BBH", b, 0); o = 4
        for c in range(colF, colL + 1):
            g = b[o]; o += 1
            if g == 1: v = struct.unpack_from("<d", b, o)[0]; o += 8
            elif g == 2: v, o = ustr(b, o)
            elif g in (4, 16): v = ("BOOL" if g == 4 else "ERR", b[o]); o += 8
            else: v = None; o += 8
            if v not in (None, ""): book["cells"].append(dict(sheet=cur, row=row+1, col=c+1, value=v))
json.dump(out, open(sys.argv[2], "w"), indent=0, ensure_ascii=False, default=str)
for bk in out: print(repr(bk["name"])[:160], len(bk["sheets"]), "sheets,", len(bk["cells"]), "cached cells", bk["sheets"][:12])
