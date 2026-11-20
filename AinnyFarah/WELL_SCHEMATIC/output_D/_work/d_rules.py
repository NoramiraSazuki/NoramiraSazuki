"""Shared rules for the Dulang C extraction: type classification, unit conversion, text parsing (imported by the build script)."""
import re

FT_PER_M = 3.28084
SRC_NAME = "Dulang D Well Schematic.pptx"


def ft(m):
    return None if m is None else round(m * FT_PER_M, 3)


def frac_in(s):
    """'2 7/8"' / '3-1/2"' / '9 5/8""' / '13-3/8"' / '26"' / '3 ½"' -> inches (float) or None"""
    s = s.replace("½", " 1/2").replace("¼", " 1/4")
    m = re.match(r'\s*(\d+)(?:[\s-]+(\d+)\s*/\s*(\d+))?\s*"', s)
    if not m:
        return None
    v = float(m.group(1))
    if m.group(2):
        v += float(m.group(2)) / float(m.group(3))
    return round(v, 4)


# (regex, Type_Key, closest-type note or None, unresolved?)
TYPE_RULES = [
    (r"INSERT\s*VALVE|\bMLO\b|-MLO", "wrsv", "insert (wireline-retrievable) safety valve set in the SCSSV/TRSV profile -> WRSV is the nearest type", False),
    (r"LINER\s*(HANGER|HGR)", "casing-hanger", "liner hanger -> Casing Hanger is the nearest type", False),
    (r"FLOAT\s*SHOE", "pup-joint", "UNRESOLVED: float shoe has no Type_Key (recorded in floatEquipment)", True),
    (r"CEMENT\s*RETAINER", "bridge-plug", "cement retainer -> Bridge Plug is the nearest type", False),
    (r"MIRAGE\s*PLUG", "mechanical-plug", "Mirage (pressure-activated barrier) plug -> Mechanical Plug is the nearest type", False),
    (r"PACK\s*-?\s*OFF", "packer", "tubing pack-off (straddle across a leak) -> Packer is the nearest type", False),
    (r"TUBING\s*PATCH|X-?SPAN", "pup-joint", "UNRESOLVED: tubing patch has no Type_Key; Pup Joint used as the nearest short tubular", True),
    (r"TUBING\s*STOP", "tubing-anchor", "wireline tubing stop -> Tubing Anchor is the nearest type", False),
    (r"NO-?\s?GO\s*(NIPPLE|SEAL\s*BORE)|\bXN\b", "nipple (no-go)", None, False),
    (r"M/SHOE", "mule-shoe", None, False),
    (r"PIN\s*X\s*.*BOX|BOX\s*X\s*.*PIN", "crossover", "pin x box sub joining connections -> Crossover is the nearest type", False),
    (r"\bCSW\b|CSW-", "sliding-sleeve", None, False),
    (r"\bPDG\b|DOWNHOLE\s*GAUGE", "pup-joint", "UNRESOLVED: permanent downhole gauge (carrier) has no Type_Key; Pup Joint used as placeholder", True),
    (r"Y-?\s?BLOCK", "crossover", "Y-block (string junction) -> Crossover is the nearest type", False),
    (r"ZONE\s*MASTER", "sliding-sleeve", "Zone Master (isolation sleeve in a nipple profile) -> Sliding Sleeve is the nearest type", False),
    (r"\bDIAL\b", "spm", "'DIAL' injection-valve mandrel in the gas-lift string -> SPM (Gas Lift) is the nearest type", False),
    (r"AUTOFILL", "circulating-valve", "autofill sub -> Circulating Valve is the nearest type", False),
    (r"SHEAR\s*SUB", "pup-joint", "UNRESOLVED: keyed safety shear sub has no Type_Key; Pup Joint used as the nearest short tubular", True),
    (r"EXTENSION|ENTENSION", "pbr", "packer extension -> PBR is the nearest type", False),
    (r"POLISHED\s*BORE|\bPBR\b", "pbr", None, False),
    (r"WIRE\s*WRAP|SCRE?E?NS?\b|SCRENS", "screen", None, False),
    (r"BLANK\s*PIPE|FLUSH\s*JOINT|EUE\s*TUBING", "tubing", "blank pipe / flush joint inside the completion -> Tubing is the nearest type", False),
    (r"\bCOLLAR\b", "pup-joint", "UNRESOLVED: casing/screen collar has no Type_Key; Pup Joint used as the nearest short tubular", True),
    (r"\bBHR\b", "packer", "'BHR' (hydraulic-set retrievable packer) -> Packer", False),
    (r"SWELL\s*PACKER", "packer", None, False),
    (r"END\s*OF\s*(TUBING|TBG)", "eot", None, False),
    (r"X-?\s?OVER|\bXO\b|CROSSOVER", "crossover", None, False),
    (r"SEAL\s*BORE\s*WT\s*NO-GO|SEAL\s*BORE", "pbr", "sealbore extension -> PBR is the nearest type", False),
    (r"PTD\s*NIPPLE", "landing-nipple", "'PTD' nipple profile -> Landing Nipple is the nearest type", False),
] + [    (r"STORM\s*CHOKE", "storm-choke", None, False),
    (r"INSERT\s*VALVE|MLO\s*IV|\bIV\b|FXE", "wrsv", "insert (wireline-retrievable) safety valve set in the SCSSV/TRSV profile -> WRSV is the nearest type", False),
    (r"BRIDGE\s*PLUG", "bridge-plug", None, False),
    (r"BULL\s*PLUG", "mechanical-plug", "bull plug (end-of-string plug) -> Mechanical Plug is the nearest type", False),
    (r"PLUG", "mechanical-plug", "wireline-set lock/plug in a nipple or SSD profile -> Mechanical Plug is the nearest type", False),
    (r"SCSSV|SSSV.*SCSSV", "scssv", None, False),
    (r"TRSSSV|TRSV", "trsv", None, False),
    (r"FLOW\s*COUPLING", "flow-coupling", None, False),
    (r"NO-?\s?GO\s*NIPPLE", "nipple (no-go)", None, False),
    (r"XN-?\s?NIPPLE", "nipple (no-go)", "XN nipple (X profile with no-go shoulder; make not stated) -> No-Go Nipple is the nearest type", False),
    (r"\bX-?\s?NIPPLE", "nipple (x)", None, False),
    (r"\bR-?\s?NIPPLE", "nipple (r)", None, False),
    (r"WP-?\s?1|WP1", "landing-nipple", "'WP-1 NIPPLE' (nipple profile, no no-go stated) -> Landing Nipple is the nearest type", False),
    (r"PTD\s*NIPPLE|NIPPLE", "landing-nipple", "nipple profile without X/R/no-go designation -> Landing Nipple is the nearest type", False),
    (r"ORIFICE", "spm", "row label names the gas-lift orifice set in a side-pocket mandrel (same MIN ID / MAX OD as the SPM rows) -> SPM (Gas Lift) is the nearest type", False),
    (r"\bSPM\b|MANDREL", "spm", None, False),
    (r"SSD|SLIDING|SIDE\s*DOOR", "sliding-sleeve", None, False),
    (r"PKR|PACKER", "packer", None, False),
    (r"TELESCOPIC", "expansion-joint", "telescopic joint -> Expansion Joint is the nearest type", False),
    (r"BLAST\s*JOINT", "blast-joint", None, False),
    (r"SEAL\s*BORE|SEALBORE|SEAL\s*EXT", "pbr", "sealbore extension -> PBR is the nearest type", False),
    (r"SEAL\s*ASS|SEAL\s*UNIT|MSN\s*SEAL", "seal-assembly", None, False),
    (r"MULE\s*SHOE|MULESHOE", "mule-shoe", None, False),
    (r"ENTRY\s*GUIDE|\bWEG\b", "mule-shoe", "wireline entry guide (WEG) at the tubing end -> Mule Shoe is the nearest type", False),
    (r"END\s*OF\s*TUBING|\bEOT\b", "eot", None, False),
    (r"3\s*WAY\s*SUB|CROSS\s*OVER|X-?OVER|ADAPTER", "crossover", "sub joining tubulars of different sizes -> Crossover is the nearest type", False),
    (r"NO-?GO\s*LOCATOR|LOCATOR", "st-slot-locator", "no-go locator -> ST Slot Locator is the nearest type", False),
    (r"SHEAR\s*JOINT|CLOSING\s*SLEEVE", "sliding-sleeve", "closing sleeve / shear joint of the gravel-pack/VTA assembly -> Sliding Sleeve is the nearest type", False),
    (r"POP\b|PUMP\s*OUT", "mechanical-plug", "pump-out plug assembly -> Mechanical Plug is the nearest type", False),
    (r"HANGER", "tubing-hanger", None, False),
    (r"CENTRALI[ZS]ER", "pup-joint", "UNRESOLVED: tubing centralizer has no Type_Key (template 'centralizers' section is for casing centralizers); Pup Joint used as the nearest short tubular", True),
    (r"GUN\s*RELEASE|DEBRIS\s*SUB|BOTTOM\s*NOSE", "pup-joint", "UNRESOLVED: TCP-gun accessory sub has no Type_Key; Pup Joint used as the nearest short tubular", True),
    (r"TCP\s*GUN", "pup-joint", "UNRESOLVED: perforating gun left in hole has no Type_Key; Pup Joint used as placeholder (not a perforation interval)", True),
    (r"PUP", "pup-joint", None, False),
]

ANNOT = re.compile(r"(?i)HUD|LEAK|FLUID\s*LEVEL|\bFL\b|RELEASE\s*@|RELEASING|CONVERTED|IDLE|IDENTIFY|MPLT|PLT\b|SGS|FGS|"
                   r"PACK\s*-?\s*OFF|PACKOFF|RETRIEVED|RECOVER|DROID|TOP OF CEMENT|CEMENT|TBG PATCH|\(\d|\d{1,2}/\d{1,2}/\d{2,4}|"
                   r"1st Packer|CONTROL\s*LINE|Dummy off|STIFF")


def classify(name):
    t = name.upper()
    for rx, key, note, unresolved in TYPE_RULES:
        if re.search(rx, t):
            return key, note, unresolved
    return None, None, False


MONTHS = {m: i + 1 for i, m in enumerate(["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"])}


def iso_date(text):
    """returns (iso, note) — iso None if not parseable"""
    t = text.upper().replace(",", " ").replace(".", " ")
    m = re.search(r"(\d{1,2})\s*(?:ST|ND|RD|TH)?\s+([A-Z]{3})[A-Z]*\s+(\d{2,4})", t)
    if m:
        d, mon, y = int(m.group(1)), MONTHS.get(m.group(2)), m.group(3)
        if mon:
            note = ""
            if len(y) == 2:
                y = "19" + y if int(y) > 50 else "20" + y
                note = "two-digit year expanded"
            return f"{int(y):04d}-{mon:02d}-{d:02d}", note
    m = re.search(r"([A-Z]{3})[A-Z]*\s*-?\s*(\d{2,4})\b", t)
    if m and MONTHS.get(m.group(1)):
        y = m.group(2)
        note = "month and year only"
        if len(y) == 2:
            y, note = ("20" + y), note + "; two-digit year expanded"
        return f"{int(y):04d}-{MONTHS[m.group(1)]:02d}", note
    return None, "date not parseable"


PERF_RX = re.compile(r"(?<![\d/])(\d{3,5}(?:\.\d+)?)\s*m?\s*[-–]\s*(\d{3,5}(?:\.\d+)?)\s*(?:m\b|m-|m_|mMD|m\s*MD|m\s*$|\s*m)", re.I)


def parse_zone_box(text):
    """-> list of (zone, top_m, bottom_m, line) and pressure/other lines. Zone = sand label as printed (after 'SAND :')."""
    perfs, other = [], []
    zone = None
    for raw in text.split("\n"):
        line = raw.strip()
        if not line:
            continue
        zm = re.match(r"(?i)SAND\s*:\s*(.+?)\s*(\(|$)", line)
        if zm:
            zone = re.sub(r"\s+SAND$", "", zm.group(1).strip(), flags=re.I)
            continue
        if re.search(r"(?i)\bPSI|\bpsi|Pr\s*=|k\s*=|Echometer", line):
            other.append(line)
            continue
        sub = None
        sm = re.search(r"(?i)(?:PERFT?\s*:?\s*)?\b(E-?\s?\d+(?:/\d+)?)\s*(?:PERFT\b|=)", line)
        if sm:
            sub = sm.group(1).replace(" ", "")
        found = False
        for m in PERF_RX.finditer(line):
            perfs.append((sub or zone, float(m.group(1)), float(m.group(2)), line))
            found = True
        if not found:
            other.append(line)
    return perfs, other
