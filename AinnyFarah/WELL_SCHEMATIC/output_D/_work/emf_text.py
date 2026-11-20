#!/usr/bin/env python3
"""Extract native text records (EMR_EXTTEXTOUTW = 84, EMR_EXTTEXTOUTA = 83) from EMF files with their reference points.

Usage: emf_text.py file.emf > out.tsv   (columns: x, y, text) — positions are logical units of the EMF.
EMF+ (GDI+ comment records) DrawString text is also decoded when present (EmfPlusDrawString, type 0x401C).
"""
import struct
import sys


def records(b):
    off = 0
    while off + 8 <= len(b):
        typ, size = struct.unpack_from("<II", b, off)
        if size < 8:
            break
        yield typ, b[off:off + size]
        off += size
        if typ == 14:  # EMR_EOF
            break


def exttext(rec, wide):
    # EMR_EXTTEXTOUT: type,size, rclBounds(16), iGraphicsMode(4), exScale(4), eyScale(4), EMRTEXT
    base = 8 + 16 + 12
    x, y, nchars, offstr = struct.unpack_from("<iiII", rec, base)
    if wide:
        s = rec[offstr:offstr + 2 * nchars].decode("utf-16-le", "replace")
    else:
        s = rec[offstr:offstr + nchars].decode("latin-1", "replace")
    return x, y, s


def emfplus_strings(rec):
    # EMR_COMMENT (70) with 'EMF+' identifier
    out = []
    if len(rec) < 16 or rec[12:16] != b"EMF+":
        return out
    off = 16
    while off + 12 <= len(rec):
        typ, flags, size, dsize = struct.unpack_from("<HHII", rec, off)
        if size < 12:
            break
        if typ == 0x401C:  # DrawString
            d = rec[off + 12: off + 12 + dsize]
            brush, fmt, length = struct.unpack_from("<III", d, 0)
            x, y, w, h = struct.unpack_from("<ffff", d, 12)
            s = d[28:28 + 2 * length].decode("utf-16-le", "replace")
            out.append((round(x, 1), round(y, 1), s))
        off += size
    return out


def main(path):
    b = open(path, "rb").read()
    n = 0
    for typ, rec in records(b):
        if typ in (83, 84):
            x, y, s = exttext(rec, typ == 84)
            if s.strip():
                print(f"{x}\t{y}\t{s}")
                n += 1
        elif typ == 70:
            for x, y, s in emfplus_strings(rec):
                if s.strip():
                    print(f"{x}\t{y}\t{s}")
                    n += 1
    print(f"# {n} text records", file=sys.stderr)


if __name__ == "__main__":
    main(sys.argv[1])
