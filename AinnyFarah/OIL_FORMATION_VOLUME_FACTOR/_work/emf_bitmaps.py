"""List EMF record types and save STRETCHDIBITS/BITBLT bitmaps as PNG."""
import struct, sys, io, collections
from pathlib import Path
from PIL import Image
for f in sys.argv[1:]:
    b = Path(f).read_bytes(); i = 0; c = collections.Counter(); n = 0
    while i + 8 <= len(b):
        t, sz = struct.unpack_from("<II", b, i)
        if sz < 8: break
        c[t] += 1
        if t in (81, 80, 76):  # STRETCHDIBITS, SETDIBITSTODEVICE, BITBLT
            off = {81: (48, 52, 56, 60), 80: (48, 52, 56, 60), 76: (84, 88, 92, 96)}[t]
            obmi, cbmi, obits, cbits = struct.unpack_from("<IIII", b, i + off[0])
            if cbmi and cbits:
                bmi = b[i+obmi:i+obmi+cbmi]; bits = b[i+obits:i+obits+cbits]
                fh = b"BM" + struct.pack("<IHHI", 14 + len(bmi) + len(bits), 0, 0, 14 + len(bmi))
                n += 1
                out = Path(f).with_name(Path(f).stem + f"_bmp{n}.png")
                Image.open(io.BytesIO(fh + bmi + bits)).save(out)
                print(out, Image.open(out).size)
        i += sz
    print(f, dict(c))
