#!/usr/bin/env python3
"""Generate 150 synthetic PDF test files (10 bytes to >1 GB) with a manifest, locally (29 Oct 2026).

Usage:  generate_synthetic_pdfs.py            plan, check disk space, generate (resumable), validate, write manifest.csv
        generate_synthetic_pdfs.py --plan     show the plan and storage requirement only
        generate_synthetic_pdfs.py --validate re-validate existing files and rewrite manifest.csv only

Design
  * A small streaming PDF writer (no third-party PDF library) writes objects sequentially to disk, hashes while writing,
    and builds the xref table, so multi-GB files never sit in memory. Every target size is hit to the byte: the last
    object written is either a text string in the page content (small files) or an embedded synthetic attachment
    (CSV production data or SEG-Y-like binary traces) whose length is solved exactly.
  * Content: fictional oil & gas themes; text pages with tables, scanned-style image pages (Pillow-rendered, noisy,
    slightly rotated, Flate), charts, and large uncompressed seismic-style RGB pages. All valid files carry
    "SYNTHETIC DATA — FOR TESTING ONLY" in the page content (text/mixed) or rendered in the image + document metadata (image).
  * Files below the smallest possible marked valid PDF are written as intentional invalid files (.pdf extension).
  * Resumable: files are written as <name>.partial and renamed when complete; _generation_state.json records finished
    files (size + SHA-256). Existing final files that are not recorded are never overwritten (the run stops).
  * Validation uses pypdf (strict mode): page count, marker text / metadata, attachments; invalid files must fail to parse.
"""
import argparse
import csv
import hashlib
import json
import math
import random
import shutil
import sys
import time
import zlib
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps

OUT = Path(__file__).resolve().parent
STATE = OUT / "_generation_state.json"
MANIFEST = OUT / "manifest.csv"
MARK = "SYNTHETIC DATA — FOR TESTING ONLY"
SEED = 20261029
FONT_PATH = "/System/Library/Fonts/Supplemental/Arial.ttf"
W, H = 595, 842  # A4 portrait (pt)

RANGES = [  # label, low (incl), high (excl; last range inclusive), count
    ("10-999 B", 10, 1_000, 10),
    ("1 KB to <100 KB", 1_000, 100_000, 25),
    ("100 KB to <1 MB", 100_000, 1_000_000, 30),
    ("1 MB to <10 MB", 1_000_000, 10_000_000, 35),
    ("10 MB to <100 MB", 10_000_000, 100_000_000, 30),
    ("100 MB to <500 MB", 100_000_000, 500_000_000, 15),
    ("500 MB to <1 GB", 500_000_000, 1_000_000_000, 3),
    ("1 GB to 1.2 GB", 1_000_000_000, 1_200_000_000, 2),
]
THEMES = ["Well_Summary", "Production_Report", "Pressure_Survey", "Seismic_Inventory", "Completion_Report",
          "Drilling_Daily_Report", "Core_Analysis", "Well_Test_Report"]
FIELDS = ["Aster Field", "Borealis Field", "Cobalt Ridge Field", "Driftwood Field", "Emberline Field", "Fjordhaven Field"]
OPERATOR = "Example Petroleum Co. (fictional)"


def in_range(size, lo, hi, last):
    return lo <= size <= hi if last else lo <= size < hi


# ------------------------------------------------------------------ streaming PDF writer
END = b"\nendstream\nendobj\n"


class TooBig(Exception):
    pass


class Writer:
    def __init__(self, fh):
        self.f, self.pos, self.off, self.next, self.h = fh, 0, {}, 1, hashlib.sha256()
        self._w(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")

    def _w(self, b):
        self.f.write(b)
        self.h.update(b)
        self.pos += len(b)

    def alloc(self):
        self.next += 1
        return self.next - 1

    def obj(self, n, body):
        self.off[n] = self.pos
        self._w(b"%d 0 obj\n" % n + body + b"\nendobj\n")

    def stream(self, n, inner, data):
        self.off[n] = self.pos
        self._w(b"%d 0 obj\n<< %s /Length %d >>\nstream\n" % (n, inner, len(data)) + data + END)

    def xref_len(self):
        n = self.next  # objects 0..next-1
        return len(b"xref\n0 %d\n" % n) + 20 * n

    def trailer(self, xref_pos, root, info, idhex):
        return (b"trailer\n<< /Size %d /Root %d 0 R" % (self.next, root) + (b" /Info %d 0 R" % info if info else b"")
                + b" /ID [<%s><%s>] >>\nstartxref\n%d\n%%%%EOF\n" % (idhex, idhex, xref_pos))

    def solve(self, head_fn, min_len, target, root, info, idhex):
        """data length L so that the final object + xref + trailer end exactly at target bytes"""
        xl = self.xref_len()
        for pad in range(4):
            L = max(min_len, target - self.pos - xl - 200)
            for _ in range(60):
                h = head_fn(L, pad)
                xp = self.pos + len(h) + L + len(END)
                tot = xp + xl + len(self.trailer(xp, root, info, idhex))
                if tot == target:
                    return L, pad
                L += target - tot
                if L < min_len:
                    raise TooBig
        raise RuntimeError("size did not converge")

    def finish(self, num, head, chunks, L, root, info, idhex):
        self.off[num] = self.pos
        self._w(head)
        n = 0
        for c in chunks:
            n += len(c)
            self._w(c)
        assert n == L, (n, L)
        self._w(END)
        xp = self.pos
        self._w(b"xref\n0 %d\n0000000000 65535 f \n" % self.next)
        self._w(b"".join(b"%010d 00000 n \n" % self.off[i] for i in range(1, self.next)))
        self._w(self.trailer(xp, root, info, idhex))


def esc(s):
    b = s.encode("cp1252", "replace")
    return b.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)")


def tx(font, size, x, y, s, rgb=None):
    col = (b"%.2f %.2f %.2f rg " % rgb) if rgb else b""
    return col + b"BT /%s %d Tf %.1f %.1f Td (%s) Tj ET\n" % (font.encode(), size, x, y, esc(s)) + (b"0 g\n" if rgb else b"")


def utf16hex(s):
    return b"<FEFF" + s.encode("utf-16-be").hex().upper().encode() + b">"


FONTS = {"F1": "Helvetica", "F2": "Helvetica-Bold", "F3": "Courier"}


def font_dict(name):
    return b"<< /Type /Font /Subtype /Type1 /BaseFont /%s /Encoding /WinAnsiEncoding >>" % name.encode()


# ------------------------------------------------------------------ synthetic content
def doc_meta(rng, theme, idx):
    field = rng.choice(FIELDS)
    well = f"{field.split()[0][:3].upper()}-{rng.randint(1, 48):02d}"
    return dict(theme=theme, field=field, well=well, idx=idx,
                date=f"20{rng.randint(18, 26)}-{rng.randint(1, 12):02d}-{rng.randint(1, 28):02d}",
                title=f"{theme.replace('_', ' ')} - {well}, {field}")


def table_spec(theme, rng):
    r = rng
    specs = {
        "Well_Summary": (["Well", "Spud Date", "TD (ft MD)", "Status", "Lift"],
                         lambda i: [f"WX-{i + 1:02d}", f"20{r.randint(5, 24)}-{r.randint(1, 12):02d}", f"{r.randint(5200, 12800):,}",
                                    r.choice(["Producing", "Shut-in", "Suspended", "Injecting"]), r.choice(["ESP", "Gas lift", "Natural"])]),
        "Production_Report": (["Date", "Oil (bbl/d)", "Gas (Mscf/d)", "Water (bbl/d)", "WC (%)"],
                              lambda i: [f"2025-{(i % 12) + 1:02d}-01", f"{r.uniform(80, 2400):.1f}", f"{r.uniform(50, 3000):.1f}",
                                         f"{r.uniform(10, 1800):.1f}", f"{r.uniform(1, 92):.1f}"]),
        "Pressure_Survey": (["Depth (ft TVD)", "Pressure (psia)", "Temp (degF)", "Gradient (psi/ft)"],
                            lambda i: [f"{1000 + 250 * i:,}", f"{900 + 108 * i + r.uniform(-5, 5):.1f}", f"{95 + 1.6 * i:.1f}", f"{r.uniform(0.30, 0.46):.3f}"]),
        "Seismic_Inventory": (["Survey ID", "Type", "Year", "Area (km2)", "Format", "Size (GB)"],
                              lambda i: [f"SV-{r.randint(100, 999)}", r.choice(["2D", "3D", "4D"]), str(r.randint(1992, 2025)),
                                         f"{r.uniform(20, 2400):.0f}", r.choice(["SEG-Y", "SEG-D", "ZGY"]), f"{r.uniform(0.2, 900):.1f}"]),
        "Completion_Report": (["Item", "Depth (ft MD)", "OD (in)", "ID (in)", "Remarks"],
                              lambda i: [r.choice(["Packer", "Sliding sleeve", "Nipple", "Flow coupling", "Mule shoe", "Screen"]),
                                         f"{r.uniform(600, 9800):.1f}", f"{r.choice([3.5, 4.5, 5.5, 7.0, 8.4]):.3f}",
                                         f"{r.choice([2.75, 2.81, 2.99, 3.81, 3.96]):.3f}", r.choice(["as run", "set", "closed", "open"])]),
        "Drilling_Daily_Report": (["Time", "Activity", "Depth (ft)", "Mud Wt (ppg)"],
                                  lambda i: [f"{(i * 2) % 24:02d}:00", r.choice(["Drilling ahead", "Circulating", "POOH", "RIH", "Survey", "Wiper trip"]),
                                             f"{3000 + 120 * i:,}", f"{r.uniform(8.6, 11.8):.1f}"]),
        "Core_Analysis": (["Sample", "Depth (ft)", "Porosity (%)", "Perm (mD)", "Grain dens."],
                          lambda i: [f"C-{i + 1:03d}", f"{7200 + 1.5 * i:.1f}", f"{r.uniform(8, 31):.1f}", f"{10 ** r.uniform(-1, 3.4):.2f}", f"{r.uniform(2.62, 2.71):.3f}"]),
        "Well_Test_Report": (["Period", "Choke (/64)", "Rate (bbl/d)", "FTHP (psi)", "GOR (scf/bbl)"],
                             lambda i: [f"P{i + 1}", str(r.choice([16, 24, 32, 40, 48])), f"{r.uniform(150, 3200):.0f}",
                                        f"{r.uniform(120, 1450):.0f}", f"{r.uniform(200, 2600):.0f}"]),
    }
    return specs[theme]


SENTENCES = [
    "This document is a fictional example prepared for software testing of {field}.",
    "Values for well {well} are randomly generated and do not describe any real asset.",
    "Reporting date {date}. Depths are measured depth below a fictional rotary table.",
    "No operator, client or personal information is contained in this report.",
    "Tables and charts are synthetic and intended for ingestion and parser testing.",
    "Reference: {operator}. Internal test identifier SYNTH-{idx:03d}.",
]


def text_page_ops(m, rng, page_no, npages_hint, limit, y_top=H - 30, with_table=True):
    """content-stream operators, adding optional elements while the total stays within limit bytes"""
    ops = [tx("F2", 9, 40, y_top, MARK, (0.8, 0, 0))]
    if sum(map(len, ops)) > limit:
        return None
    opt = [tx("F2", 13, 40, y_top - 22, m["title"]),
           tx("F1", 7, 40, 18, f"{MARK}   |   SYNTH-{m['idx']:03d}   |   page {page_no}" + (f" of {npages_hint}" if npages_hint else ""))]
    y = y_top - 44
    for s in SENTENCES:
        opt.append(tx("F1", 9, 40, y, s.format(operator=OPERATOR, **m)))
        y -= 13
    if with_table:
        cols, row = table_spec(m["theme"], rng)
        cw = (W - 80) / len(cols)
        y -= 10
        opt.append(b"0.85 g 40 %.1f %.1f 14 re f 0 g\n" % (y - 4, W - 80) + b"".join(tx("F2", 7, 43 + j * cw, y, c) for j, c in enumerate(cols)))
        i = 0
        while y > 60:
            y -= 14
            cells = row(i + page_no * 40)
            opt.append(b"0.6 G 0.4 w 40 %.1f %.1f 14 re S 0 G\n" % (y - 4, W - 80) + b"".join(tx("F3" if j else "F1", 7, 43 + j * cw, y, c) for j, c in enumerate(cells)))
            i += 1
    size = sum(map(len, ops))
    for o in opt:
        if size + len(o) > limit:
            break
        ops.append(o)
        size += len(o)
    return b"".join(ops)


def font(px):
    return ImageFont.truetype(FONT_PATH, max(6, int(px)))


def scanned_gray(m, rng, w, h, page_no):
    img = Image.new("L", (w, h), 242)
    d = ImageDraw.Draw(img)
    s = w / 850
    d.text((40 * s, 30 * s), MARK, fill=40, font=font(22 * s))
    d.text((40 * s, 70 * s), m["title"], fill=20, font=font(26 * s))
    y = 120 * s
    for t in SENTENCES:
        d.text((40 * s, y), t.format(operator=OPERATOR, **m), fill=30, font=font(15 * s))
        y += 24 * s
    cols, row = table_spec(m["theme"], rng)
    cw = (w - 80 * s) / len(cols)
    y += 16 * s
    i = 0
    while y < h - 80 * s:
        cells = cols if i == 0 else row(i + page_no * 40)
        d.rectangle([40 * s, y, w - 40 * s, y + 24 * s], outline=90)
        for j, c in enumerate(cells):
            d.text((44 * s + j * cw, y + 4 * s), c, fill=25, font=font(13 * s))
        y += 24 * s
        i += 1
    d.text((40 * s, h - 40 * s), f"{MARK} - scanned page {page_no}", fill=40, font=font(13 * s))
    img = img.rotate(rng.uniform(-1.2, 1.2), resample=Image.BICUBIC, fillcolor=242)
    noise = Image.effect_noise((w, h), rng.uniform(8, 22))
    return Image.blend(img, noise, rng.uniform(0.06, 0.14))


def chart_rgb(m, rng, w, h):
    img = Image.new("RGB", (w, h), (255, 255, 255))
    d = ImageDraw.Draw(img)
    x0, y0, x1, y1 = int(w * 0.07), int(h * 0.09), int(w * 0.98), int(h * 0.89)
    for k in range(6):
        yy = y0 + (y1 - y0) * k / 5
        d.line([x0, yy, x1, yy], fill=(225, 225, 225))
    d.rectangle([x0, y0, x1, y1], outline=(60, 60, 60))
    for colr in ((20, 90, 200), (200, 60, 30), (30, 150, 60)):
        q0, dcl = rng.uniform(0.5, 0.95), rng.uniform(0.01, 0.08)
        pts = [(x0 + (x1 - x0) * t / 120, y1 - (y1 - y0) * q0 * math.exp(-dcl * t) * (1 + rng.uniform(-0.03, 0.03))) for t in range(121)]
        d.line(pts, fill=colr, width=2)
    d.text((x0, 2), f"{m['theme'].replace('_', ' ')} trend - {m['well']} (synthetic)", fill=(0, 0, 0), font=font(w / 56))
    d.text((x0, int(h * 0.92)), MARK, fill=(200, 0, 0), font=font(w / 64))
    return img


def seismic_rgb(m, rng, w, h, page_no):
    g = Image.effect_noise((max(8, w // 60), h), 70).resize((w, h), Image.BICUBIC)
    g = g.rotate(rng.uniform(-7, 7), resample=Image.BICUBIC, fillcolor=128)
    g = Image.blend(g, Image.effect_noise((w, h), 30), 0.22)
    img = ImageOps.colorize(g, black=(25, 45, 190), white=(195, 25, 25), mid=(248, 248, 248))
    d = ImageDraw.Draw(img)
    for x in range(0, w, max(50, w // 20)):
        d.line([x, 0, x, 12], fill=(0, 0, 0))
    d.rectangle([0, 0, w - 1, h - 1], outline=(0, 0, 0))
    d.text((20, 20), f"{MARK}  |  inline {1000 + page_no}  |  {m['field']} synthetic seismic", fill=(0, 0, 0), font=font(max(12, w // 70)))
    return img


def img_obj(w_, img, flate):
    raw = img.tobytes()
    cs = b"/DeviceGray" if img.mode == "L" else b"/DeviceRGB"
    inner = b"/Type /XObject /Subtype /Image /Width %d /Height %d /ColorSpace %s /BitsPerComponent 8" % (img.width, img.height, cs)
    if flate:
        return inner + b" /Filter /FlateDecode", zlib.compress(raw, 6)
    return inner, raw


# ------------------------------------------------------------------ attachments
def csv_chunks(rng, m, L):
    out, n = [], 0
    head = (f"# {MARK}\n# fictional production export for {m['well']}, {m['field']}\n"
            "date,well,oil_bbl_d,gas_mscf_d,water_bbl_d,wc_pct,choke_64,fthp_psi\n").encode()
    buf = head
    day = 0
    while n < L:
        lines = []
        for _ in range(4000):
            day += 1
            lines.append(f"{2000 + day // 365}-{(day // 30) % 12 + 1:02d}-{day % 28 + 1:02d},{m['well']},{rng.uniform(50, 2600):.2f},"
                         f"{rng.uniform(40, 3200):.2f},{rng.uniform(5, 2100):.2f},{rng.uniform(0.5, 95):.2f},{rng.choice((16, 24, 32, 40, 48))},{rng.uniform(90, 1500):.1f}\n")
        buf += "".join(lines).encode()
        take = buf[:L - n]
        n += len(take)
        yield take
        buf = b""


def bin_chunks(rng, m, L):
    head = (f"{MARK}\nSYNTHETIC SEG-Y-LIKE TRACE DATA (random IEEE float bytes), survey SV-{m['idx']:03d}, {m['field']} - fictional\n").encode()
    head = head.ljust(3200, b" ")[:L]
    yield head
    n = len(head)
    while n < L:
        k = min(8 << 20, L - n)
        yield rng.randbytes(k)
        n += k


# ------------------------------------------------------------------ file builders
def write_invalid(path, spec):
    size, kind = spec["target"], spec["invalid_kind"]
    rng = random.Random(SEED + spec["idx"])
    if kind == "truncated header":
        data = (b"%PDF-1.7\n%" + b"SYNTHETIC INVALID TEST FILE " * 40)[:size]
    elif kind == "random bytes":
        data = rng.randbytes(size)
    elif kind == "plain text":
        data = ("SYNTHETIC DATA - FOR TESTING ONLY - INTENTIONALLY INVALID PDF (plain text). " * 20).encode()[:size]
    else:  # "incomplete object"
        data = (b"%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n2 0 obj\n<< /Type /Pages /Kids [" + b" " * 2000)[:size]
    path.write_bytes(data)
    return dict(pages=None, sha=hashlib.sha256(data).hexdigest(), remarks=f"INTENTIONALLY INVALID ({kind}); {spec['invalid_reason']}")


def build_single_page(f, spec, m, limit):
    """one page, text filler in the last content stream (small files)"""
    T, compact = spec["target"], spec["target"] < 4000
    w = Writer(f)
    idhex = hashlib.md5(spec["name"].encode()).hexdigest().upper().encode()
    root, pages = w.alloc(), w.alloc()
    info = None if compact else w.alloc()
    rng = random.Random(SEED * 7 + spec["idx"])
    if compact:
        res = b"<< /Font << /F1 %s >> >>" % font_dict("Helvetica")
        base = text_page_ops(m, rng, 1, 1, limit, with_table=False)
        base = base and base.replace(b"/F2 ", b"/F1 ").replace(b"/F3 ", b"/F1 ")
    else:
        fnum = {k: w.alloc() for k in FONTS}
        for k, v in FONTS.items():
            w.obj(fnum[k], font_dict(v))
        res = b"<< /Font << %s >> >>" % b" ".join(b"/%s %d 0 R" % (k.encode(), fnum[k]) for k in FONTS)
        base = text_page_ops(m, rng, 1, 1, limit)
    if base is None:
        raise TooBig
    content, page = w.alloc(), w.alloc()
    w.obj(page, b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %d %d] /Resources %s /Contents %d 0 R >>" % (pages, W, H, res, content))
    w.obj(pages, b"<< /Type /Pages /Kids [%d 0 R] /Count 1 >>" % page)
    w.obj(root, b"<< /Type /Catalog /Pages %d 0 R >>" % pages)
    if info:
        w.obj(info, info_dict(m))
    pre, suf = b"BT /F1 2 Tf 6 4 Td (", b") Tj ET\n"
    fixed = len(base) + len(pre) + len(suf)
    head = lambda L, pad: b"%d 0 obj\n<<%s /Length %d >>\nstream\n" % (content, b" " * pad, L)
    L, pad = w.solve(head, fixed, T, root, info, idhex)
    filler = (b"SYNTHETIC TEST CONTENT " * (L // 20 + 2))[:L - fixed]
    w.finish(content, head(L, pad), [base, pre + filler + suf], L, root, info, idhex)
    return w, dict(pages=1, remarks=f"1 text page ({'compact, inline font' if compact else 'with fonts + metadata'}); content-stream text filler {L - fixed} B")


def info_dict(m):
    return (b"<< /Title " + utf16hex(m["title"]) + b" /Subject " + utf16hex(MARK) + b" /Keywords " + utf16hex("synthetic; test data; fictional; " + m["theme"])
            + b" /Author " + utf16hex("Synthetic test-data generator") + b" /Producer (generate_synthetic_pdfs.py) >>")


def build_multi(f, spec, m, frac):
    T, typ, sc = spec["target"], spec["pdf_type"], spec["range_idx"]
    rng = random.Random(SEED * 11 + spec["idx"])
    w = Writer(f)
    idhex = hashlib.md5(spec["name"].encode()).hexdigest().upper().encode()
    root, pages_n, info = w.alloc(), w.alloc(), w.alloc()
    fnum = {k: w.alloc() for k in FONTS}
    for k, v in FONTS.items():
        w.obj(fnum[k], font_dict(v))
    fonts = b"/Font << %s >>" % b" ".join(b"/%s %d 0 R" % (k.encode(), fnum[k]) for k in FONTS)
    budget = T * frac
    kids, comp = [], {"text": 0, "scanned": 0, "chart": 0, "seismic": 0}
    cap = {0: 1, 1: 6, 2: 30, 3: 200, 4: 60, 5: 80, 6: 90, 7: 100}[sc]
    scan_dims = {1: (260, 368), 2: (520, 736), 3: (850, 1202)}
    raw_dims = {4: (1200, 900), 5: (2000, 1500), 6: (2000, 1500), 7: (2000, 1500)}

    def room():
        return budget - w.pos - 1500 - 40 * w.next

    def add_page(content, xobj=None, box=(W, H)):
        c, p = w.alloc(), w.alloc()
        w.stream(c, b"", content)
        xo = b" /XObject << /Im1 %d 0 R >>" % xobj if xobj else b""
        w.obj(p, b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %d %d] /Resources << %s%s >> /Contents %d 0 R >>" % (pages_n, box[0], box[1], fonts, xo, c))
        kids.append(p)

    while len(kids) < cap:
        pno = len(kids) + 1
        kind = "text" if typ == "text" else ("seismic" if sc >= 4 else "scanned") if typ == "image" else \
            ("text" if pno % 2 == 1 else ("seismic" if sc >= 4 else "chart"))
        r = room()
        if kind == "text":
            lim = min(r - 400, 6000) if typ != "mixed" or comp["chart"] + comp["seismic"] else min(r * 0.4, 6000)
            ops = text_page_ops(m, rng, pno, None, lim)
            if ops is None or r < 800:
                break
            add_page(ops)
        elif kind == "seismic":
            iw, ih = raw_dims[sc]
            if iw * ih * 3 + 2000 > r:
                break
            img = seismic_rgb(m, rng, iw, ih, pno)
            inner, data = img_obj(w, img, False)
            n = w.alloc()
            w.stream(n, inner, data)
            if typ == "image":
                add_page(b"q 842 0 0 595 0 0 cm /Im1 Do Q\n", n, (842, 595))
            else:
                hdr = text_page_ops(m, rng, pno, None, 900, with_table=False)
                add_page(hdr + b"q 515 0 0 386 40 300 cm /Im1 Do Q\n", n)
        else:
            if kind == "scanned":
                iw, ih = scan_dims[sc]
                data = None
                for _ in range(8):
                    img = scanned_gray(m, rng, iw, ih, pno)
                    inner, data = img_obj(w, img, True)
                    if len(data) + 1500 <= r:
                        break
                    iw, ih, data = int(iw * 0.75), int(ih * 0.75), None
                if data is None:
                    break
                n = w.alloc()
                w.stream(n, inner, data)
                add_page(b"q %d 0 0 %d 0 0 cm /Im1 Do Q\n" % (W, H), n)
            else:
                cw_, data = 900, None
                while cw_ >= 120:
                    img = chart_rgb(m, rng, cw_, cw_ // 2)
                    inner, data = img_obj(w, img, True)
                    if len(data) + 2500 <= r:
                        break
                    cw_, data = int(cw_ * 0.7), None
                if data is None:
                    break
                n = w.alloc()
                w.stream(n, inner, data)
                hdr = text_page_ops(m, rng, pno, None, 1200, with_table=False)
                add_page(hdr + b"q 515 0 0 257 40 420 cm /Im1 Do Q\n", n)
        comp[kind] += 1
    if not kids:
        raise TooBig
    if typ == "mixed" and not comp["chart"] + comp["seismic"]:
        raise RuntimeError(f"{spec['name']}: mixed document without an image page")
    w.obj(pages_n, b"<< /Type /Pages /Kids [%s] /Count %d >>" % (b" ".join(b"%d 0 R" % k for k in kids), len(kids)))
    fs, ef = w.alloc(), w.alloc()
    binary = sc >= 4
    fname = f"synthetic_traces_SV-{spec['idx']:03d}.bin" if binary else f"synthetic_production_{m['well']}.csv"
    w.obj(root, b"<< /Type /Catalog /Pages %d 0 R /Names << /EmbeddedFiles << /Names [%s %d 0 R] >> >> >>" % (pages_n, utf16hex(fname), fs))
    w.obj(info, info_dict(m))
    w.obj(fs, b"<< /Type /Filespec /F (%s) /UF %s /EF << /F %d 0 R >> /Desc %s >>" % (fname.encode(), utf16hex(fname), ef, utf16hex(MARK + " - embedded attachment")))
    sub = b"/application#2Foctet-stream" if binary else b"/text#2Fcsv"
    head = lambda L, pad: b"%d 0 obj\n<< /Type /EmbeddedFile /Subtype %s%s /Length %d >>\nstream\n" % (ef, sub, b" " * pad, L)
    L, pad = w.solve(head, 1, T, root, info, idhex)
    chunks = bin_chunks(rng, m, L) if binary else csv_chunks(rng, m, L)
    w.finish(ef, head(L, pad), chunks, L, root, info, idhex)
    parts = ", ".join(f"{v} {k} page{'s' if v != 1 else ''}" for k, v in comp.items() if v)
    return w, dict(pages=len(kids), remarks=f"{parts}; embedded attachment {fname} {L:,} B ({L / T:.1%} of file)")


def generate(spec, path):
    m = doc_meta(random.Random(SEED + spec["idx"]), spec["theme"], spec["idx"])
    partial = path.with_name(path.name + ".partial")
    T = spec["target"]
    tries = ([("single", lim) for lim in single_limits(T)] if T < 8000 else
             [("multi", max(0.05, spec["frac"] - 0.08 * a)) for a in range(12)])
    for mode, arg in tries:
        with open(partial, "wb") as f:
            try:
                w, info = build_single_page(f, spec, m, limit=arg) if mode == "single" else build_multi(f, spec, m, frac=arg)
            except TooBig:
                continue
        assert w.pos == spec["target"] == partial.stat().st_size, (w.pos, spec["target"])
        partial.rename(path)
        return dict(info, sha=w.h.hexdigest())
    partial.unlink(missing_ok=True)
    raise RuntimeError(f"could not fit content into {spec['target']} bytes")


# ------------------------------------------------------------------ plan
def single_limits(T):
    """content-size limits tried (largest first) for single-page files"""
    over, step = (150, 20) if T < 4000 else (1500, 120)
    return list(range(T - over, 59, -step))[:40] or [60]


def min_valid_size():
    import io
    m = doc_meta(random.Random(1), "Well_Summary", 0)
    for T in range(300, 1000):
        for lim in single_limits(T):
            try:
                build_single_page(io.BytesIO(), dict(target=T, name=f"SYNTH_000_{m['theme']}_valid.pdf", idx=0), m, limit=lim)
                return T
            except TooBig:
                continue
    return None


def make_plan():
    rng = random.Random(SEED)
    minv = min_valid_size()
    specs = []
    types_cycle = {1: ["text", "text", "image", "text", "mixed"], 2: ["text", "image", "mixed"], 3: ["text", "image", "mixed", "mixed", "image"],
                   4: ["image", "mixed"], 5: ["mixed", "image"], 6: ["image", "mixed", "image"], 7: ["mixed", "image"]}
    for ri, (label, lo, hi, n) in enumerate(RANGES):
        last = ri == len(RANGES) - 1
        if ri == 0:
            n_inv = 5 if minv else n
            sizes = [int(lo * ((minv - 1) / lo) ** ((i + rng.random()) / n_inv)) for i in range(n_inv)]
            sizes += [int(minv + (hi - 1 - minv) * (i + rng.random()) / (n - n_inv)) for i in range(n - n_inv)]
        else:
            lo2 = lo + 1 if last else lo
            sizes = [int(lo2 * (hi / lo2) ** ((i + rng.random()) / n)) for i in range(n)]
        for j, s in enumerate(sorted(sizes)):
            s = min(max(s, lo + (1 if last else 0)), hi if last else hi - 1)
            valid = not (ri == 0 and s < minv)
            typ = "intentionally invalid" if not valid else ("text" if ri == 0 else types_cycle[ri][j % len(types_cycle[ri])])
            if typ == "image" and s < 20_000:
                typ = "text"
            if typ == "mixed" and s < 15_000:
                typ = "text"
            specs.append(dict(range=label, range_idx=ri, lo=lo, hi=hi, last=last, target=s, valid=valid, pdf_type=typ,
                              frac=round(rng.uniform(0.55, 0.85), 3)))
    for i, sp in enumerate(specs, 1):
        sp["idx"] = i
        if sp["valid"]:
            sp["theme"] = THEMES[(i * 3) % len(THEMES)]
            sp["name"] = f"SYNTH_{i:03d}_{sp['theme']}_valid.pdf"
        else:
            sp["theme"] = "Tiny_File"
            sp["invalid_kind"] = ["truncated header", "random bytes", "plain text", "incomplete object"][i % 4]
            sp["invalid_reason"] = f"{sp['target']} B is below the smallest valid marked PDF this generator can write ({minv} B)"
            sp["name"] = f"SYNTH_{i:03d}_Tiny_File_invalid.pdf"
    return specs, minv


# ------------------------------------------------------------------ validation + manifest
def sha_file(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        while (b := f.read(16 << 20)):
            h.update(b)
    return h.hexdigest()


def validate(specs, state):
    import pypdf
    rows, fails = [], []
    present = sorted(p.name for p in OUT.glob("*.pdf"))
    for sp in specs:
        p = OUT / sp["name"]
        rec = state["done"].get(sp["name"], {})
        if not p.exists():
            fails.append(f"{sp['name']}: missing")
            continue
        size = p.stat().st_size
        sha = sha_file(p)
        notes, ok, pages = [], True, None
        if sha != rec.get("sha"):
            ok = False; notes.append("checksum differs from generation record")
        if not in_range(size, sp["lo"], sp["hi"], sp["last"]) or size != sp["target"]:
            ok = False; notes.append(f"size {size} outside range/target")
        if sp["valid"]:
            try:
                rd = pypdf.PdfReader(str(p), strict=True)
                pages = len(rd.pages)
                if pages != rec.get("pages"):
                    ok = False; notes.append(f"page count {pages} != {rec.get('pages')}")
                if sp["pdf_type"] == "image":
                    xo = rd.pages[0]["/Resources"]["/XObject"]
                    has_img = any(xo[k].get("/Subtype") == "/Image" for k in xo)
                    mark = MARK in (rd.metadata.subject or "")
                    ok &= has_img and mark
                    detail = f"image page; marker in metadata{' + rendered in image' if mark else ' MISSING'}"
                else:
                    txt = rd.pages[0].extract_text() or ""
                    mark = MARK in txt
                    ok &= mark
                    detail = "marker text found on page 1" if mark else "marker text NOT found"
                if sp["target"] >= 8000:
                    att = rd.attachments
                    ok &= len(att) == 1
                    detail += f"; {len(att)} attachment"
                res = f"{'PASS' if ok else 'FAIL'}: opened with pypdf {pypdf.__version__} (strict), {pages} page(s); {detail}" + (f"; {'; '.join(notes)}" if notes else "")
            except Exception as e:  # noqa: BLE001
                ok = False
                res = f"FAIL: pypdf error {type(e).__name__}: {str(e)[:120]}"
        else:
            try:
                pypdf.PdfReader(str(p), strict=True)
                ok = False
                res = "FAIL: intentionally invalid file parsed as a PDF"
            except Exception as e:  # noqa: BLE001
                res = f"PASS (invalid as intended): pypdf {type(e).__name__}" + (f"; {'; '.join(notes)}" if notes else "")
                ok = ok and not notes
        if not ok:
            fails.append(f"{sp['name']}: {res}")
        rows.append({"Filename": sp["name"], "Content category": sp["theme"].replace("_", " "), "PDF type": sp["pdf_type"], "Size range": sp["range"],
                     "Target size (bytes)": sp["target"], "Actual size (bytes)": size, "Actual size (MB)": f"{size / 1e6:.6f}",
                     "Actual size (GB)": f"{size / 1e9:.9f}", "Page count": pages if pages is not None else "", "Expected validity": "valid" if sp["valid"] else "invalid (intentional)",
                     "Validation result": res, "SHA-256": sha, "Remarks": rec.get("remarks", "")})
        print(f"  validated {sp['name']:48} {res[:90]}", flush=True)
    names = [r["Filename"] for r in rows]
    shas = [r["SHA-256"] for r in rows]
    checks = [
        (f"exactly 150 .pdf files in {OUT.name}", len(present) == 150 and set(present) == {s["name"] for s in specs}),
        ("every file within its assigned size range", all(in_range(r["Actual size (bytes)"], s["lo"], s["hi"], s["last"]) for r, s in zip(rows, specs))),
        ("at least one file > 1,000,000,000 bytes", any(r["Actual size (bytes)"] > 1_000_000_000 for r in rows)),
        ("every expected-valid file opens with a PDF parser", all(r["Validation result"].startswith("PASS") for r in rows if r["Expected validity"] == "valid")),
        ("intentional invalid files recorded and rejected by the parser", all(r["Validation result"].startswith("PASS (invalid") for r in rows if r["Expected validity"] != "valid")),
        ("filenames unique", len(set(names)) == len(names)),
        ("SHA-256 checksums unique", len(set(shas)) == len(shas)),
    ]
    with open(MANIFEST, "w", newline="", encoding="utf-8") as f:
        wr = csv.DictWriter(f, fieldnames=list(rows[0]))
        wr.writeheader(); wr.writerows(rows)
    return rows, checks, fails


def summary(specs, rows, checks, fails):
    print("\nFile counts by size range")
    for label, lo, hi, n in RANGES:
        rs = [r for r in rows if r["Size range"] == label]
        tot = sum(r["Actual size (bytes)"] for r in rs)
        inv = sum(1 for r in rs if r["Expected validity"] != "valid")
        print(f"  {label:20} {len(rs):4d} files (planned {n}){f', {inv} intentionally invalid' if inv else ''}   {tot:>16,} B")
    total = sum(r["Actual size (bytes)"] for r in rows)
    print(f"  {'TOTAL':20} {len(rows):4d} files                      {total:>16,} B  ({total / 1e9:.3f} GB)")
    print("  by type: " + ", ".join(f"{t} {sum(1 for r in rows if r['PDF type'] == t)}" for t in ("text", "image", "mixed", "intentionally invalid")))
    print("\nChecks")
    for c, ok in checks:
        print(f"  {'PASS' if ok else 'FAIL'} - {c}")
    print(f"\nValidation failures: {len(fails)}")
    for x in fails:
        print("  " + x)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--plan", action="store_true")
    ap.add_argument("--validate", action="store_true")
    a = ap.parse_args()
    state = json.loads(STATE.read_text()) if STATE.exists() else {}
    if "plan" not in state:
        specs, minv = make_plan()
        state = {"plan": specs, "min_valid_size": minv, "done": {}}
    specs = state["plan"]
    total = sum(s["target"] for s in specs)
    todo = [s for s in specs if s["name"] not in state["done"]]
    need = sum(s["target"] for s in todo)
    free = shutil.disk_usage(OUT).free
    print(f"Plan: {len(specs)} files, total {total:,} B ({total / 1e9:.3f} GB); smallest valid marked PDF = {state['min_valid_size']} B")
    print(f"Remaining to generate: {len(todo)} files, {need:,} B ({need / 1e9:.3f} GB); free disk space {free:,} B ({free / 1e9:.1f} GB)")
    if a.plan:
        for s in specs:
            print(f"  {s['name']:48} {s['pdf_type']:22} {s['target']:>15,} B  [{s['range']}]")
        return
    if not a.validate:
        margin = 2_000_000_000
        if need + margin > free:
            sys.exit(f"STOP: insufficient disk space — need {need:,} B + {margin:,} B safety margin, only {free:,} B free. Nothing generated.")
        for s in todo:
            p = OUT / s["name"]
            if p.exists():
                sys.exit(f"STOP: {p.name} already exists but is not recorded as generated by this script — not overwriting.")
        STATE.write_text(json.dumps(state, indent=1))
        for k, s in enumerate(specs, 1):
            p = OUT / s["name"]
            if s["name"] in state["done"]:
                print(f"[{k:3d}/150] {s['name']:48} target {s['target']:>15,} B  actual {p.stat().st_size:>15,} B  SKIPPED (already complete)", flush=True)
                continue
            t0 = time.time()
            rec = write_invalid(p, s) if not s["valid"] else generate(s, p)
            state["done"][s["name"]] = dict(rec, size=p.stat().st_size)
            STATE.write_text(json.dumps(state, indent=1))
            print(f"[{k:3d}/150] {s['name']:48} target {s['target']:>15,} B  actual {p.stat().st_size:>15,} B  "
                  f"{'OK' if p.stat().st_size == s['target'] else 'SIZE MISMATCH'}  ({time.time() - t0:.1f}s)", flush=True)
    print("\nValidating ...")
    rows, checks, fails = validate(specs, state)
    summary(specs, rows, checks, fails)
    print(f"\nmanifest: {MANIFEST}")


if __name__ == "__main__":
    main()
