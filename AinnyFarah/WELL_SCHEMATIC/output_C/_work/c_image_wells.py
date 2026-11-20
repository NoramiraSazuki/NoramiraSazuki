"""Image-based slides (3, 16, 22, 25, 29, 31): schematic is a picture, values transcribed from EMF native text (slide 22)
or tesseract OCR (x2/x3 upscale, ocr/*.tsv) confirmed by visual inspection of zoomed crops (ocr/*_check.png).

Each well: header values -> wellInfo; rows -> components. 'where' strings give slide, media file and the image position
(pixel row in the source image) or EMF text-record position. Depth datum is kept as printed per value.
"""
import re

from c_rules import FT_PER_M, classify, frac_in, ft

LABELS = {3: "WELL NAME: DULANG C-03ST1 (image3.jpeg)", 16: "WELL C-16 (slide title; image13.png)",
          22: "WELL NAME: DULANG C-22 ST2 (image17.emf)", 25: "WELL NAME: DULANG C-25 (image21.png)",
          29: "WELL NAME: DULANG C-29 (image23.png)", 31: "WELL: DULANG C-31 (image26.emf)"}

# ---------------------------------------------------------------- per-well transcription
# row: (name, top_m, bot_m, ID, OD, y, extra) ; ID/OD from explicit 'MIN ID IN' / 'MAX OD IN' (or 'ID' / 'OD') columns
# extra keys: string, status, idrule (size in inches from the label, no ID/OD designation), typ/closest override, note

UPPER_ACCESSORIES = "Tubing String (OIL PRODUCER) column"

C3 = dict(
    media="image3.jpeg", method="OCR (tesseract psm 11, x3) + visual check", datum="m MDTHF (depth column header 'DEPTH m MDTHF')",
    revision="PROPOSED COMPLETION SCHEMATIC (Vinod - Dec 2012)",
    info=dict(completionDate=("December 2012", 139), rig=("KM-1", 153), maxDeviation=("66.75 DEG @ 2108 M MDDF / 1335 M TVDDF", 186),
              tubing=('3-1/2" 13 Cr 9.2 PPF & 2-7/8" 13Cr 6.4 PPF Inner String; 6-1/8" Open Hole wt 125 micron 4-1/2" 12.6 PPF PGAM Screen', 248),
              completionFluid=("(LOWER COMPLETION) 10.2 ppg Sodium Formate (treated with fluid loss); (UPPER COMPLETION) 8.5 ppg KCL", 236),
              packerFluid=("8.5 ppg Drill water + KCl based fluid + 0.1% vol. Bacteride + 2% vol. CONQOR 303A Corrosion Inhibitor + 1% vol. O2 scavenger", 268),
              completionType=("OIL PRODUCER (Tubing String header); single string", 348)),
    td=(2108, "MD DF", 168, "TD : 2108 M MD DF"),
    xmas=("Conventional Solar Alert 5K Wellhead & Single F-F Trim Christmas Tree", 153),
    rows=[("FLOW COUPLING", 222, None, 2.880, 3.920, 411, {}), ("TRSSSV", 224, None, 2.813, 5.030, 444, {"note": "OCR read 'RSSSV'; visual 'TRSSSV'"}),
          ("FLOW COUPLING", 225, None, 2.880, 3.920, 473, {}),
          ('3-1/2" SPM', 398, None, 2.880, 6.000, 599, {"idrule": None}), ('3-1/2" SPM', 700, None, 2.880, 6.000, 628, {}),
          ('3-1/2" SPM', 994, None, 2.880, 6.000, 657, {}), ('3-1/2" SPM', 1273, None, 2.880, 6.000, 686, {}),
          ('3-1/2" Orifice', 1535, None, 2.880, 6.000, 728, {}),
          ('3-1/2" SSD', 1775, None, 2.813, 4.550, 777, {"status": "SSD-1 Closed"}),
          ('7" x 3-1/2" AHC Packer', 1780, None, 2.900, 5.960, 823, {}),
          ('3-1/2" XN-Nipple', None, None, 2.666, 3.893, 890, {}), ('3-1/2" No-Go Locator', None, None, None, None, 908, {}),
          ('3-1/2" Self Aligning Muleshoe', 1842, None, 3.880, 6.000, 1012,
           {"note": "CONFLICT: MIN ID 3.880 / MAX OD 6.000 printed on the muleshoe row equal the VTA-packer values of the same "
                    "design on C-22/C-29 (rows may be offset by one); kept as printed", "conflict": True}),
          ('4-1/2" x 7" VTA Packer', 1844, None, 3.894, 4.967, 1070,
           {"note": "CONFLICT: MIN ID 3.894 / MAX OD 4.967 equal the sealbore-extension values on C-22/C-29; kept as printed", "conflict": True}),
          ('3.88" ID Bore NO Seal Ext', None, None, ("explicit", 3.88), None, 1084, {}),
          ('3 Way Sub : 4-1/2"x 4-1/2"x 2-7/8" Box X Pin X Pin', None, None, None, None, 1170, {}),
          ("3 Meter Oil Swell Packer", 1887, 1890, None, None, 1322, {"string": None}),
          ('3.22" SLIM SSD Inner String', 1896, None, None, None, 1339, {"idrule": 3.22, "status": "SSD-2 Open"}),
          ("3 Meter Oil Swell Packer", 1914, 1917, None, None, 1423, {}),
          ('3.25" Seal Bore Ext with Seal Unit', 1919, None, None, None, 1453, {"idrule": 3.25}),
          ("No-Go Locator Inner String", None, None, None, None, 1469, {}),
          ("3 Meter Oil Swell Packer", 1935, 1938, None, None, 1495, {}),
          ('3.22" SLIM SSD Inner String', 1946, None, None, None, 1507, {"idrule": 3.22, "status": "SSD-3 Closed"}),
          ('3.25" Seal Bore Ext with Seal Unit', 1956, None, None, None, 1633, {"idrule": 3.25}),
          ("3 Meter Oil Swell Packer", 1961, 1964, None, None, 1647, {}),
          ('3.22" SLIM SSD Inner String', 1968, None, None, None, 1659, {"idrule": 3.22, "status": "SSD-4 Closed"}),
          ('3.25" Seal Bore Ext with Seal Unit', 1971, None, None, None, 1738, {"idrule": 3.25}),
          ("3 Meter Oil Swell Packer", 1976, 1979, None, None, 1776, {}),
          ('3.22" SLIM SSD Inner String', 1990, None, None, None, 1809, {"idrule": 3.22, "status": "SSD-5 Closed"}),
          ("3 Meter Oil Swell Packer", 2022, 2025, None, None, 2010, {}),
          ('3.25" Seal Bore Ext with Seal Unit', 2027, None, None, None, 2030, {"idrule": 3.25}),
          ('3.22" SLIM XN-Nipple Inner String', 2029, None, None, None, 2046, {"idrule": 3.22}),
          ('3.22" Self Aligning Muleshoe Inner String', 2035, None, None, None, 2064, {"idrule": 3.22, "end": "inner"})],
    shoe=dict(name="SHOE 1874 / 1246 TVD DF", m=1874, tvd=1246, y=1150, datum_note="datum not printed for 1874 (in the m MDTHF depth column); TVD 1246 m DF",
              est=True),
    float_shoe=("Double Valve Float Shoe", 2040, 2148),
    tubing=[('3-1/2" 13 Cr 9.2 PPF', 3.5, 9.2, "13 Cr", 248, "upper"), ('2-7/8" 13Cr 6.4 PPF Inner String', 2.875, 6.4, "13Cr", 248, "inner")],
    screen=('6-1/8" Open Hole wt 125 micron 4-1/2" 12.6 PPF PGAM Screen', 4.5, 12.6, 268),
    zones=[("E2", 1891, 1902, "MD: 1911-1922 mDF (1891 - 1902 m MDTHF)", 1379, "MDTHF"),
           ("E6", 1938, 1943, "E6 (1938 -1943 m MDTHF) MD: 1958-1963 mDF", 1534, "MDTHF"),
           ("E7", 1949, 1954, "MD: 1969-1974 mDF (1949 - 1954 m MDTHF)", 1612, "MDTHF"),
           ("E10-11", 1967, 1969, "MD: 1987-1989 mDF (1967 - 1969 m MDTHF)", 1720, "MDTHF"),
           ("E10-11", 1981, 2020, "MD: 2001-2040 mDF (1981 - 2020 m MDTHF)", 1880, "MDTHF")],
    interventions=[("FXE IV – 22/09/2024", None, None, 8, "slide text box (no depth)"),
                   ("5/4/18 : set PXN plug", None, None, 2, "slide text box (no depth)")],
)

C16 = dict(
    media="image13.png", method="OCR (tesseract psm 11, x3) + visual check (low-resolution 671x969 image)", datum="m MD RKB (depth column header 'DEPTH m MD RKB')",
    revision="Vinod - Jan 2013",
    info=dict(completionDate=("December 2012", 7), rig=("KM-1", 17), maxDeviation=("32.90 DEG @ 911 M MDDF / 837.22 M TVD", 48),
              tubing=('4 1/2" CS L-80 Lined wt GRE wt 25 Cr Completion Accessories RiceWarp Lower Completion', 88),
              completionFluid=("8.5 ppg KCL", 88),
              packerFluid=("8.5 ppg Drill water + KCl based fluid + 0.1% vol. Bacteride + 2% vol. CONQOR 303A Corrosion Inhibitor + 1% vol. O2 scavenger", 99),
              completionType=("WAG Injector (Tubing STRING header '( WAG Injector)')", 133)),
    td=(1621, "MD DF", 28, "TD : 1621 M MD DF"),
    xmas=("Conventional Solar Alert 5K Wellhead & Single H-H Trim Christmas Tree", 68),
    rows=[("FLOW COUPLING", 242, None, None, None, 177, {"status": "FLOW COUPLING - DDLAA BOX - TOP & BOTTOM"}),
          ("TRSSSV", 243, None, None, None, 194, {}),
          ("FLOW COUPLING", 245, None, None, None, 211, {"status": "FLOW COUPLING - DDLAA BOX - TOP & BOTTOM"}),
          ("SSD", 1350, None, None, None, 468, {"status": "SSD - DDLAA BOX - TOP & BOTTOM"}),
          ('THT 9-5/8" x 3-1/2" Packer', 1378, None, None, None, 520,
           {"status": '7" x 4 1/2" AHC PACKER DDLAA BOX - TOP & BOTTOM',
            "note": "CONFLICT: equipment column '9-5/8\" x 3-1/2\" Packer' vs remarks '7\" x 4 1/2\" AHC PACKER' — both recorded, sizes not assigned", "conflict": True}),
          ("SSD", 1383, None, None, None, 565, {"status": "SSD - DDLAA BOX - TOP & BOTTOM"}),
          ('THT 9-5/8" x 3-1/2" Packer', 1470, None, None, None, 659,
           {"status": '7" x 4 1/2" AHC PACKER DDLAA BOX - TOP & BOTTOM', "note": "CONFLICT: equipment vs remarks packer size (as above)", "conflict": True}),
          ("SSD", 1475, None, None, None, 680, {"status": "Zone Master - DDLAA BOX - TOP & BOTTOM", "note": "OCR read 'sso'; visual 'SSD'"}),
          ('THT 9-5/8" x 3-1/2" Packer', 1502, None, None, None, 806,
           {"status": '7" x 4 1/2" AHC PACKER DDLAA BOX - TOP & BOTTOM', "note": "CONFLICT: equipment vs remarks packer size (as above)", "conflict": True}),
          ("SSD", 1508, None, None, None, 839, {"status": "Zone Master - DDLAA BOX - TOP & BOTTOM"}),
          ("BULL PLUG - DDLAA BOX - TOP & BOTTOM", 1533, None, None, None, 884,
           {"note": "name from the Remarks column on the 1533 row (equipment cell holds the E12-13 zone box); slide note 'Bull Plug @ 1516m-MDTHF (30/10/2022)' "
                    "gives the same plug on the MDTHF datum — diagram value kept, both recorded", "end": "upper"})],
    casing=[('9-58" 40 PPF L-80 CS (depth: TD)', 9.625, 40.0, "L-80", None, 17,
             "size printed '9-58\"' read as 9-5/8\"; DEPTH column says 'TD' (shoe depth not given as a number)")],
    pbtd=(1587, "M MD DF", 38, "PBTD : 1587 M MD DF (digits partly clipped in the image; bottom label '587 M MD D' clipped)"),
    tubing=[('4 1/2" CS L-80 Lined wt GRE', 4.5, None, "L-80", 88, "upper")],
    zones=[("LD1", 1370, 1375, "LD1 1370-1375 M MDTHF", 590, "MDTHF"), ("LD1", 1376, 1379, "1376-1379 M MD THF", 608, "MDTHF"),
           ("E6-8 / E10-11", 1456, 1471, "1456-1471 M MDTHF", 722, "MDTHF"), ("E6-8 / E10-11", 1473, 1475, "1473-1475 M MDTHF", 732, "MDTHF"),
           ("E6-8 / E10-11", 1476, 1480, "1476-1480 M MDTHF", 742, "MDTHF"),
           ("E12-13", 1490, 1500, "MD: 1490-1500 m MDTHF", 870, "MDTHF"), ("E12-13", 1502, 1505, "MD: 1502-1505 m MDTHF", 877, "MDTHF")],
    interventions=[("Storm Choke Jammed Open  (30/7/2022)", None, None, 7, "slide text box (no depth)")],
)

C22_ROWS_EMF = "image17.emf native text record"
C22 = dict(
    media="image17.emf", method="EMF native text (EMR_EXTTEXTOUTW records) + visual check", datum="m MDTHF (depth column header 'DEPTH m MDTHF')",
    revision="ACTUAL COMPLETION SCHEMATIC", emf=True,
    info=dict(completionDate=("November-12", 155), rig=("KM-1", 183), maxDeviation=("69.21 deg @ 729.21 m MD/ 597.43 m TVD", 237),
              tubing=('3-1/2" 13 Cr 9.2 PPF & 2-7/8" 13Cr 6.4 PPF Inner String; 6-1/8" Open Hole wt 125 micron 4-1/2" 12.6 PPF PGAM Screen', 347),
              completionFluid=("(LOWER COMPLETION) 10.2 ppg Sodium Formate (treated with fluid loss); (UPPER COMPLETION) 8.5 ppg KCL", 320),
              packerFluid=("8.5 ppg Drill water + KCl based fluid + 0.1% vol. Bacteride + 2% vol. CONQOR 303A Corrosion Inhibitor + 1% vol. O2 scavenger", 376),
              completionType=("OIL PRODUCER (Tubing String header); single string", 495)),
    td=(1663, "MDDF", 211, "TD : 1663 m MDDF (bottom label 'TD = 1663 m MDDF')"),
    xmas=(":Conventional Solar Alert 5K Wellhead & Single F-F Trim Christmas Tree", 184),
    rows=[("FLOW COUPLING", 222, None, 2.880, 3.920, 594, {}), ("TRSSSV", 224, None, 2.81, 5.030, 639, {}),
          ("FLOW COUPLING", 225, None, 2.880, 3.920, 688, {}),
          ('3-1/2" SPM', 284, None, 2.880, 6.000, 796, {"status": "Dummy 9/9/2023 (slide note)"}),
          ('3-1/2" SPM', 526, None, 2.880, 6.000, 877, {"status": "Dummy 9/9/2023 (slide note)"}),
          ('3-1/2" SPM', 779, None, 2.880, 6.000, 958, {"status": "Dummy 9/9/2023 (slide note)"}),
          ('3-1/2" SPM', 1038, None, 2.880, 6.000, 1039, {"status": "GLV 9/9/2023 (slide note)"}),
          ('3-1/2" Orifice', 1302, None, 2.880, 6.000, 1120, {"status": "GLOV 9/9/2023 (slide note)"}),
          ('3-1/2" SSD', 1475, None, 2.813, 4.550, 1198, {"status": "CLOSED SSD#1 17/10/2016 (slide note)"}),
          ('7" x 3-1/2" AHC Packer', 1480, None, 2.900, 5.960, 1291, {}),
          ('3-1/2" XN-Nipple', 1505, None, 2.666, 3.893, 1391, {}), ('3-1/2" No-Go Locator', None, None, None, None, 1416, {}),
          ('3-1/2" Self Aligning Muleshoe', 1510, None, None, None, 1472, {"end": "upper"}),
          ('4-1/2" x 7" VTA Packer', 1512, None, 3.880, 6.000, 1499, {}),
          ('3.88" Seal Bore Ext', None, None, 3.894, 4.967, 1526, {}),
          ('3 Way Sub : 4-1/2"x 4-1/2"x 2-7/8" Box X Pin X Pin', None, None, None, None, 1584, {}),
          ("3 Meter Oil Swell Packer", 1546, 1549, None, None, 1687, {}),
          ('3.22" SLIM SSD Inner String', 1552, None, None, None, 1749, {"idrule": 3.22, "status": "Open SSD#2 08/03/2020 (slide note)"}),
          ('3.25" Seal Bore Ext', 1556, None, None, None, 1857, {"idrule": 3.25}),
          ("3 Meter Oil Swell Packer", 1566, 1569, None, None, 1916, {}),
          ('3.22" SLIM SSD Inner String', 1572.0, None, None, None, 2030, {"idrule": 3.22, "status": "CLOSED due to water SSD#3 17/10/2016 (slide note)"}),
          ('3.25" Seal Bore Ext', None, None, None, None, 2138, {"idrule": 3.25}),
          ("3 Meter Oil Swell Packer", 1574, 1577, None, None, 2192, {}),
          ("3 Meter Oil Swell Packer", 1582, 1585, None, None, 2273, {}),
          ('3.22" SLIM SSD Inner String', 1592, None, None, None, 2321, {"idrule": 3.22, "status": "OPEN SSD#4 (23/10/2016) (slide note)"}),
          ('3.25" Seal Bore Ext', 1596, None, None, None, 2483, {"idrule": 3.25}),
          ('3.22" SLIM XN-Nipple Inner String', 1598, None, None, None, 2589, {"idrule": 3.22}),
          ('3.22" Self Aligning Muleshoe Inner String', 1600, None, None, None, 2616, {"idrule": 3.22, "end": "inner"}),
          ("3 Meter Oil Swell Packer", 1600, 1603, None, None, 2695, {})],
    shoe=dict(name='7" SHOE 1562 m MDDF 1230 TVD DF', m=1562, tvd=1230, y=1584, datum_note="m MDDF as printed (not MDTHF); TVD 1230 m DF", est=True, size=7.0),
    float_shoe=("Double Valve Float Shoe", 1605, 2776),
    tubing=[('3-1/2" 13 Cr 9.2 PPF', 3.5, 9.2, "13 Cr", 347, "upper"), ('2-7/8" 13Cr 6.4 PPF Inner String', 2.875, 6.4, "13Cr", 347, "inner")],
    screen=('6-1/8" Open Hole wt 125 micron 4-1/2" 12.6 PPF PGAM Screen', 4.5, 12.6, 376),
    zones=[("E6", 1572, 1574, "E6 MD: 1572 - 1574 m DF", 1804, "DF"), ("E8", 1590, 1593, "E8 MD: 1590 - 1593 Mdf", 2003, "DF"),
           ("E10-11", 1610, 1614, "E10-11 MD: 1610 - 1614 mDF", 2402, "DF")],
    interventions=[('2.31” FXE Insert Valve @ 139m-MDTHF (24/9/2024)', 139, "MDTHF", 10, "slide text box"),
                   (' Set 2.31"PXN Plug @ 1602 mMDTHF (22/10/2016)', 1602, "MDTHF", 17,
                    "slide text box; second note 'TOP PXN Plug @ 1601m-THF (1/6/2017)' (#9) gives 1601 m -> CONFLICT recorded, 1602 m (set depth) kept")],
)

C25 = dict(
    media="image21.png", method="OCR (tesseract psm 11, x3) + visual check", datum="m MDTHF (depth columns '(m MDTHF) +/-')",
    revision="FINAL COMPLETION SCHEMATIC (AFTER WORKOVER), prepared SHSI 10/07/2011", dual=True,
    info=dict(completionDate=("5 JULY 2011", 106), rig=("HWU GAIT 2/ BARGE GD9", 121), maxDeviation=("Max Dev: 30.3 Degree @ 675.30 m MDTHF", 178),
              tubing=('2-7/8", 6.4 PPF, 13CR, K-FOX', 136),
              completionFluid=("8.6 PPG BRINE + 3% KCl", 190),
              packerFluid=("8.6 PPG BRINE + 3% KCl + 0.2% BIOCIDE + 2% OXYGEN SCAVENGER + 2% CORROSION INHIBITOR", 164),
              surfaceCasing=('13 3/8" SURF. 54.5 PPF K-55 @ 694.4 mMD-THF', 121),
              completionType=("Dulang C-25 FINAL COMPLETION SCHEMATIC (AFTER WORKOVER); SHORT STRING / LONG STRING", 66)),
    td=(1932.4, "MDTHF", 1422, "TD @ 1932.4 m-MDTHF"),
    xmas=('IKS 3-1/8" 5,000 PSI DUAL X-MAS TREE (SOLAR ALERT)', 149),
    rows=[("Flow Coupling", 138.26, None, 2.355, None, 326, {"string": "Short"}),
          ("TRSV Wellstar", 139.11, None, 2.313, None, 340, {"string": "Short"}),
          ("Flow Coupling", 140.40, None, 2.355, None, 354, {"string": "Short", "note": "OCR read '2,365'; visual 2.355"}),
          ("SPM #1", 711.24, None, 2.355, None, 383, {"string": "Short", "status": '8/64" GLV Ptro @ 996 psi'}),
          ("SPM #2", 1073.80, None, 2.355, None, 423, {"string": "Short", "status": '8/64" GLV Ptro @ 717 psi'}),
          ("SPM #3", 1177.86, None, 2.355, None, 466, {"string": "Short", "status": '8/64" GLOV'}),
          ("X-Nipple", 1188.12, None, 2.313, None, 508, {"string": "Short"}),
          ('9-5/8" BHD Packer', 1207.08, None, 2.380, None, 566,
           {"string": "Short & Long", "note": "INTERPRETED: listed in the short-string column, but the drawing shows both strings through this packer (dual) -> represented once as shared"}),
          ("XN-Nipple", 1209.75, None, 2.250, None, 609, {"string": "Short"}),
          ("POP Assembly", 1210.84, None, 2.275, None, 637, {"string": "Short"}),
          ("EOT", 1212.48, None, None, None, 665, {"string": "Short", "end": "Short"}),
          ("VTA Packer", 1222.11, None, 5.000, None, 724, {"lower": True}),
          ("Closing Sleeve Assy", 1225.29, None, 5.000, None, 738, {"lower": True}),
          ("Shear Joint Assy", 1231.29, None, 4.872, None, 766, {"lower": True}),
          ("Sealbore Extension", 1254.21, None, 3.250, None, 880, {"lower": True, "note": "OCR read 1264.21; visual 1254.21 (sequence 1231.29 < 1254.21 < 1255.04)"}),
          ("VTA Packer", 1255.04, None, 5.000, None, 905, {"lower": True}),
          ("Closing Sleeve Assy", 1258.32, None, 5.000, None, 922, {"lower": True}),
          ("Shear Joint Assy", 1264.23, None, 4.872, None, 951, {"lower": True}),
          ("AWD Packer (Top of Sump Packer)", 1304.00, None, 5.000, None, 1081, {"lower": True}),
          ("Flow Coupling", 147.4, None, 2.355, None, 354, {"string": "Long"}),
          ("TRSV Wellstar", 148.3, None, 2.313, None, 354, {"string": "Long"}),
          ("Flow Coupling", 149.5, None, 2.355, None, 366, {"string": "Long"}),
          ("SPM#1", 701.4, None, 2.355, None, 407, {"string": "Long", "status": '8/64" GLV Ptro @ 747 psi'}),
          ("SPM #2", 1064.7, None, 2.355, None, 448, {"string": "Long", "status": '8/64" GLV Ptro @ 740 psi'}),
          ("SPM #3", 1168.5, None, 2.355, None, 494, {"string": "Long", "status": '12/64" GLOV',
                                                       "note": "OCR read 4168.5; visual 1168.5 (last digits low resolution: 1166.5 also possible)", "unclear": True}),
          ("SSD #1", 1197.2, None, 2.313, None, 537, {"string": "Long"}),
          ("SSD #2", 1217.7, None, 2.313, None, 665, {"string": "Long"}),
          ("No-Go Locator", 1253.2, None, 2.380, None, 880, {"string": "Long"}),
          ('3.25" MSN Seal', 1253.7, None, 2.350, None, 894, {"string": "Long", "note": "label size 3.25\" not used: explicit MIN ID column value 2.350 recorded"}),
          ("ADAPTER", 1253.7, None, 2.419, None, 905, {"string": "Long"}),
          ("XN-Nipple", 1278.8, None, 2.250, None, 922, {"string": "Long"}),
          ("Self Aligning Mule Shoe", 1279.9, None, 2.441, None, 951, {"string": "Long"}),
          ("EOT", 1280.7, None, None, None, 979, {"string": "Long", "end": "Long"}),
          ('9-5/8" Bridge Plug #2', 1388.0, None, None, None, 1123, {"string": None}),
          ('9-5/8" Bridge Plug #1', 1528.5, None, None, None, 1151, {"string": None})],
    screens=[('5-1/2" WWS 0.012 GA', 5.5, 800, "beside LD1 SAND (long-string side)"), ('5-1/2" WWS 0.012 GA', 5.5, 1020, "beside E2 SAND (short-string side)")],
    cement_top=(1476.0, 1138, "Top of Cement"),
    casing=[('13 3/8" SURF. 54.5 PPF K-55 @ 694.4 mMD-THF', 13.375, 54.5, "K-55", 694.4, 121, None),
            ('9-5/8" 43.5 PPF, N-80 PROD, 10 PPF @1650 m-MDTHF', 9.625, None, "N-80", 1650.0, 136,
             "CONFLICT: two weights printed (43.5 PPF and 10 PPF) -> weight left empty; diagram '9-5/8\" Shoe 1650.0' agrees with the depth; grade read N-80 (OCR 'N-60')"),
            ('7" LINER N-80 @ 1930.4 m-MDTHF', 7.0, None, "N-80", 1930.4, 149, "diagram 'END OF 7\" LINER 1930.4 m-MDTHF' agrees")],
    pbtd=(1873.4, "MD-THF", 164, "PBTD: 1873.4 m MD-THF (7\" Bridge Plug)"),
    extra_pbtd_bp=('7" BRIDGE PLUG', 1873.4, 1294, "PBTD 7\" BRIDGE PLUG 1873.4 m-MDTHF"),
    tubing=[('2-7/8", 6.4 PPF, 13CR, K-FOX', 2.875, 6.4, "13CR", 136, "Short"), ('2-7/8", 6.4 PPF, 13CR, K-FOX', 2.875, 6.4, "13CR", 136, "Long")],
    zones=[("LD1", 1244.80, 1250.70, "LD1 SAND TOP PERF: 1244.80 m-MDTHF BOTTOM PERF: 1250.70 m-MDTHF", 820, "MDTHF"),
           ("E2", 1293.0, 1299.7, "E2 SAND TOP PERF: 1293 m-MDTHF BOTTOM PERF: 1299.7 m-MDTHF (6.7 meters interval)", 1020, "MDTHF")],
    zones_are_perfs=True,
    interventions=[("2.31” FXE Insert Valve @ 139m-WLTHF (15/10/2022)", 139, "WLTHF", 9, "slide text box (short-string side)", "Short"),
                   ("2.31” FXE Insert Valve @ 148m-WLTHF (9/9/2023)", 148, "WLTHF", 4, "slide text box (long-string side)", "Long")],
)

C29 = dict(
    media="image23.png", method="OCR (tesseract psm 11, x3) + visual check", datum="m MD RKB (depth column header 'DEPTH MD m RKB')",
    revision="PROPOSED COMPLETION SCHEMATIC (Vinod - Oct-12)",
    info=dict(completionDate=("October2012", 110), rig=("KM-1", 122), maxDeviation=("71.49 DEG @ 1027.95 M MD DF / 789.70 M TVD DF", 150),
              tubing=('3-1/2" 13 Cr 9.2 PPF & 2-7/8" 13Cr 6.4 PPF Inner String; 6-1/8" Open Hole wt 125 micron 4-1/2" 12.6 PPF PGAM Screen', 206),
              completionFluid=("(LOWER COMPLETION) 10.2 ppg Sodium Formate (treated with fluid loss); (UPPER COMPLETION) 8.5 ppg KCL", 195),
              packerFluid=("8.5 ppg Drill water + KCl based fluid + 0.1% vol. Bacteride + 2% vol. CONQOR 303A Corrosion Inhibitor + 1% vol. O2 scavenger", 223),
              completionType=("OIL PRODUCER (Tubing String header); single string", 292)),
    td=(2620, "DF", 136, "TD : 2620 Meter DF (bottom label 'TD = 2620 M DF')"),
    xmas=("Conventional Solar Alert 5K Wellhead & Single F-F Trim Christmas Tree", 123),
    rows=[("FLOW COUPLING", 242, None, 2.880, 3.920, 350, {}), ("TRSSSV", 245, None, 2.813, 5.030, 377, {}),
          ("FLOW COUPLING", 248, None, 2.880, 3.920, 405, {}),
          ('3-1/2" SPM', 480, None, 2.880, 6.000, 542, {"status": "Dummy 1/10/2020 (slide note)"}),
          ('3-1/2" SPM', 826, None, 2.880, 6.000, 587, {"status": "Dummy 1/10/2020 (slide note)"}),
          ('3-1/2" SPM', 1173, None, 2.880, 6.000, 633, {"status": "Dummy 1/10/2020 (slide note)"}),
          ('3-1/2" SPM', 1511, None, 2.880, 6.000, 663, {"status": "Dummy 1/10/2020 (slide note)"}),
          ('3-1/2" SPM', 1840, None, 2.880, 6.000, 692, {"status": "Dummy 1/10/2020 (slide note)"}),
          ('3-1/2" Orifice', 2159, None, 2.880, 6.000, 721, {"status": "GLOV 16/64 1/10/2020 (slide note)"}),
          ('3-1/2" SSD', 2288, None, 2.813, 4.550, 774, {}),
          ('7" x 3-1/2" AHC Packer', 2295, None, 2.900, 5.960, 820, {}),
          ('3-1/2" XN-Nipple', 2300, None, 2.666, 3.893, 870, {}), ('3-1/2" No-Go Locator', None, None, None, None, 885, {}),
          ('3-1/2" Self Aligning Muleshoe', 2326, None, None, None, 918, {"end": "upper"}),
          ('4-1/2" x 7" VTA Packer', 2327, None, 3.880, 6.000, 960, {}),
          ('3.88" Seal Bore Ext', None, None, 3.894, 4.967, 974, {}),
          ('3 Way Sub : 4-1/2"x 4-1/2"x 2-7/8" Box X Pin X Pin', None, None, None, None, 1033, {}),
          ('3.22" SLIM SSD Inner String', 2408, None, None, None, 1133, {"idrule": 3.22}),
          ("3 Meter Oil Swell Packer", 2400, 2403, None, None, 1163, {}),
          ('3.25" Seal Bore Ext', 2412, None, None, None, 1312, {"idrule": 3.25}),
          ("3 Meter Oil Swell Packer", 2416, 2419, None, None, 1356, {}),
          ('3.22" SLIM SSD Inner String', 2428, None, None, None, 1372, {"idrule": 3.22, "status": "OPEN SSD #02 (24/10/2017) (slide note)"}),
          ('3.25" Seal Bore Ext', 2431, None, None, None, 1473, {"idrule": 3.25}),
          ("3 Meter Oil Swell Packer", 2431, 2434, None, None, 1491, {}),
          ('3.22" SLIM SSD Inner String', 2449, None, None, None, 1578, {"idrule": 3.22, "status": "OPEN SSD #03 (24/10/2017) (slide note)"}),
          ("3 Meter Oil Swell Packer", 2458, 2461, None, None, 1652, {}),
          ('3.25" Seal Bore Ext', 2472, None, None, None, 1756, {"idrule": 3.25}),
          ("3 Meter Oil Swell Packer", 2469, 2472, None, None, 1802, {}),
          ('3.22" SLIM XN-Nipple Inner String', 2480, None, None, None, 1882, {"idrule": 3.22}),
          ('3.22" Self Aligning Muleshoe Inner String', 2482, None, None, None, 1897, {"idrule": 3.22, "end": "inner"})],
    shoe=dict(name='7" SHOE 2357 m MDDF 1268 TVD DF', m=2357, tvd=1268, y=1033, datum_note="m MDDF as printed (not MD RKB like the equipment column); TVD 1268 m DF", est=True, size=7.0),
    float_shoe=("Double Valve Float Shoe", 2487, 1959),
    tubing=[('3-1/2" 13 Cr 9.2 PPF', 3.5, 9.2, "13 Cr", 206, "upper"), ('2-7/8" 13Cr 6.4 PPF Inner String', 2.875, 6.4, "13Cr", 206, "inner")],
    screen=('6-1/8" Open Hole wt 125 micron 4-1/2" 12.6 PPF PGAM Screen', 4.5, 12.6, 223),
    zones=[("E7", 2405, 2410, "E 7 MD: 2405-2410 mDF", 1197, "DF"), ("E10-11 A", 2426, 2429, "E 10-11 A MD: 2426-2429 mDF", 1409, "DF"),
           ("E10-11 B", 2440, 2458, "E 10-11 B MD: 2440-2458 mDF", 1529, "DF"), ("E10-11 C", 2462, 2470, "E 10-11 C MD: 2462-2470 mDF", 1690, "DF")],
    interventions=[],
)

C31 = dict(
    media="image26.emf (embedded 1656x2339 bitmap)", method="OCR (tesseract psm 11, x2) + visual check of zoomed crops",
    datum="m RKB ('All Depths Base On RKB'; column 'DEPTH')",
    revision="HALLIBURTON COMPLETION REPORT, DATE 24/10/2012",
    info=dict(completionDate=("24/10/2012", 266), maxDeviation=("65.03deg @818m MMDF", 352),
              tubing=('4 1/2" 12.6ppf L80 JFE-FOX @ 2450.654 m', 284), completionFluid=("8.5ppg drill water + 3% KCl", 284),
              completionType=('4 1/2" SINGLE WAG INJECTOR', 248)),
    td=(2514, "MDDF", 1821, "TD @ 2514 m MDDF"),
    rows=[('Hanger-Single 13 5/8" X 4 1/2" KFOX', None, None, 3.905, 13.480, 463,
           {"note": "OD read 13.480 (last digit clipped in the image: '13.48')", "typ": "tubing-hanger"}),
          ("Flow Coupling", 246.12, None, 2.882, 4.916, 554, {}),
          ('3-1/2" TRSV,NE,5.20, 2.813,H2S, 5K, 25CR', 246.93, None, 2.813, 5.200, 574, {"note": "OCR read 246.33; visual 246.93"}),
          ("Flow Coupling", 248.43, None, 2.882, 4.916, 592, {}),
          ("Flow Coupling", 2313.91, None, 2.882, 4.916, 1013, {}),
          ('3-1/2" Sliding Side Door XD', 2314.75, None, 2.813, 4.550, 1032, {}),
          ("Flow Coupling", 2316.05, None, 2.882, 4.916, 1050, {"note": "depth not returned by OCR; visual 2316.05"}),
          ('7" AHC Packer (Chemical cut or Overpull 120kips to release)', 2321.43, None, 2.902, 5.930, 1198, {}),
          ("Flow Coupling", 2324.94, None, 2.882, 4.916, 1307, {}),
          ('PTD Nipple with 2.688" Zone Master', 2325.77, None, 2.688, 4.030, 1324, {}),
          ("Flow Coupling", 2326.36, None, 2.882, 4.916, 1362, {}),
          ('7" AHC Packer (Chemical cut or Overpull 120kips to release)', 2410.24, None, 2.902, 5.930, 1472, {"note": "OCR read 2440.24; visual 2410.24"}),
          ("Flow Coupling", 2413.76, None, 2.882, 4.916, 1600, {}),
          ('PTD Nipple with 2.562" Zone Master', 2414.57, None, 2.562, 3.884, 1618, {"note": "OCR read 2414.87; visual 2414.57"}),
          ("Flow Coupling", 2415.16, None, 2.882, 4.916, 1652, {}),
          ("Bull Plug", 2450.31, None, None, 4.916, 1728, {"note": "OCR read 2459.34; visual 2450.31"}),
          ("EOT", 2450.65, None, None, None, 1746, {"end": "upper", "note": "OCR read 2459.65; visual 2450.65"})],
    casing=[('CASING 7" 29ppf L80 VAM TOP', 7.0, 29.0, "L80", 2480.224, 248, "OCR read '28ppf'; visual 29ppf; DEPTH (M) column, datum RKB per report")],
    pbtd=(2482, "MDDF", 1802, "PBTD : 2482 m MDDF"),
    tubing=[('TUBING 4 1/2" 12.6ppf L80 JFE-FOX', 4.5, 12.6, "L80", 284, "upper", 2450.654)],
    zones=[("E10-11", 2329, 2342, "E10 - 11 : 2329 - 2342 m MDDF", 1802, "MDDF"), ("E10-11", 2353, 2395, ": 2353 - 2395 m MDDF", 1821, "MDDF"),
           ("E12-13", 2422, 2445, "E12 - 13 : 2422 - 2445 m MDDF", 1840, "MDDF")],
    interventions=[],
)

SPECS = {3: C3, 16: C16, 22: C22, 25: C25, 29: C29, 31: C31}
PRIMARY = {3: "MDTHF", 16: "RKB", 22: "MDTHF", 25: "MDTHF", 29: "RKB", 31: "RKB"}


def build(w, dump, add_comp):
    S = SPECS[w.slide]
    s = w.slide
    med = S["media"]
    prim = PRIMARY[s]
    w.info = {}
    w.datum = S["datum"]

    def where(y):
        if S.get("emf"):
            return f"slide {s} {med} text record y={y}"
        return f"slide {s} {med} pixel row ~{y}"

    for k, (v, y) in S["info"].items():
        val = v
        if k == "completionDate":
            from c_rules import iso_date
            iso, note = iso_date(v)
            if iso is None and re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", v):
                d, m, yy = v.split("/")
                iso, note = f"{yy}-{int(m):02d}-{int(d):02d}", "dd/mm/yyyy"
            w.info[k] = (iso or v, where(y), v, f"{S['method']}; ISO date {note}".strip())
        else:
            w.info[k] = (val, where(y), v, S["method"])
    if S.get("td"):
        m, datum, y, orig = S["td"]
        w.info["totalDepth"] = (ft(m), where(y), orig, f"{m} m x {FT_PER_M} = {ft(m)} ft; datum {datum} as printed (not shifted); {S['method']}")
        w.td = (m, datum)
    w.issue("Unclear", f"slide {s}", f"schematic is an embedded image ({med}); values transcribed by {S['method']}; revision: {S['revision']}")
    if "PROPOSED" in S["revision"]:
        w.issue("Conflict", f"slide {s}", f"image is a {S['revision']} — slide notes added later (status/HUD/plugs) are recorded separately; not merged into the proposed equipment list")

    def dtag(datum):
        return datum if datum else prim

    seq = []
    ends = {}
    for name, top, bot, idd, od, y, ex in S["rows"]:
        typ, closest, unres = classify(name)
        if ex.get("typ"):
            typ, closest, unres = ex["typ"], None, False
        if re.search(r"(?i)Swell Packer", name):
            typ, closest, unres = "packer", None, False
        notes, est, extra = [], [], []
        if ex.get("note"):
            (notes if ex.get("conflict") or ex.get("unclear") else extra).append(ex["note"])
            if ex.get("conflict"):
                w.issue("Conflict", where(y), f"'{name}': {ex['note']}")
            if ex.get("unclear"):
                w.issue("Unclear", where(y), f"'{name}': {ex['note']}")
                est.append("topMd")
        unit = f"{S['datum']}"
        if prim != "MDTHF":
            unit += " — NOT MDTHF; kept on the printed datum"
        topv = (top, where(y), f"{top}" if bot is None else f"{top}-{bot}", unit) if top is not None else None
        botv = (bot, where(y), f"{top}-{bot}", unit + "; range end printed with the item") if bot is not None else None
        iddv = None
        if isinstance(idd, tuple):
            iddv = (idd[1], where(y), name, "explicit 'ID' in the label")
        elif idd is not None:
            iddv = (idd, where(y), f"{idd:.3f}", "column 'MIN ID IN' = explicit ID")
        elif ex.get("idrule"):
            iddv = (ex["idrule"], where(y), name, "USER-DEFINED RULE: label size without ID/OD designation -> innerDiameter")
            est.append("innerDiameter")
            notes.append(f"innerDiameter: label size {ex['idrule']}\" has no ID/OD designation -> recorded as ID (user-defined rule)")
        odv = (od, where(y), f"{od:.3f}", "column 'MAX OD IN' / 'OD' = explicit OD") if od is not None else None
        string = ex.get("string", "Short" if False else None)
        if "string" not in ex and S.get("dual"):
            string = None
        if ex.get("lower"):
            notes.append("lower-completion (gravel-pack/screen) assembly item listed in the short-string column below the short-string EOT -> string not assigned")
        if ex.get("status"):
            extra.append(f"status/remark '{ex['status']}'")
        cid = w.cid(("" if not string or string == "Short & Long" else string[0].lower() + "-") + typ.replace(" (", "-").replace(")", ""))
        d = add_comp(w, cid=cid, typ=typ, name=name, name_src=(where(y), name, S["method"]), string=string, top=topv, bot=botv,
                     idd=iddv, od=odv, closest=closest, unresolved=unres, est=est, notes=notes, src_extra=extra + [f"depth datum: {S['datum']}"])
        if ex.get("end"):
            ends[ex["end"]] = d
        seq.append((string, d))
    # sequence check per string (image wells list items top-down)
    last = {}
    for string, d in seq:
        k = string or "-"
        if d["topMdM"] is None:
            continue
        p = last.get(k)
        if p and d["topMdM"] < p["topMdM"] and not re.search(r"Swell|Bridge", d["name"] + p["name"]):
            w.issue("Conflict", d["id"], f"depth {d['topMdM']} m shallower than the item listed above ({p['id']} {p['topMdM']} m) — kept as printed")
            d["source"] += f"; CONFLICT: out of depth sequence (item above {p['id']} at {p['topMdM']} m)"
        last[k] = d
    # casing / shoe
    for c in S.get("casing", []):
        name, size, wt, grade, bot, y, note = c
        typ = "liner" if "LINER" in name.upper() else "casing"
        notes = ["innerDiameter: header size has no ID/OD designation -> recorded as ID (user-defined rule)"]
        est = ["innerDiameter"]
        if note:
            if note.startswith("CONFLICT"):
                w.issue("Conflict", where(y), f"'{name}': {note}")
            notes.append(note)
        datum_note = S["datum"] if s != 25 else "m MDTHF as printed"
        add_comp(w, cid=w.cid("production-casing" if "PROD" in name.upper() or s in (16, 31) else ("liner" if typ == "liner" else "surface-casing")),
                 typ=typ, name=name, name_src=(where(y), name, S["method"]),
                 bot=(bot, where(y), name, datum_note) if bot is not None else None,
                 idd=(size, where(y), name, "USER-DEFINED RULE: size without ID/OD designation -> innerDiameter"),
                 weight=(wt, where(y), name) if wt else None, grade=(grade, where(y), name) if grade else None, est=est, notes=notes)
        if bot is None:
            w.issue("Missing", where(y), f"'{name}': shoe depth not given as a number")
    if S.get("shoe"):
        sh = S["shoe"]
        notes = [f"DATUM: {sh['datum_note']}"]
        est = ["bottomMd"]
        idd = None
        if sh.get("size"):
            idd = (sh["size"], where(sh["y"]), sh["name"], "USER-DEFINED RULE: size without ID/OD designation -> innerDiameter")
            est.append("innerDiameter")
            notes.append("innerDiameter: '7\"' has no ID/OD designation -> recorded as ID (user-defined rule)")
        else:
            notes.append("casing size not printed")
        d = add_comp(w, cid=w.cid("casing-shoe"), typ="casing", name=sh["name"], name_src=(where(sh["y"]), sh["name"], S["method"]),
                     bot=(sh["m"], where(sh["y"]), sh["name"], sh["datum_note"]), idd=idd, est=est, notes=notes)
        d["bottomTvd"], d["bottomTvdM"] = ft(sh["tvd"]), float(sh["tvd"])
        d["source"] += f"; bottomTvd {sh['tvd']} m DF x {FT_PER_M} = {ft(sh['tvd'])} ft (TVD DF as printed)"
        w.tr(f"components[{d['id']}].bottomTvd", d["bottomTvd"], where(sh["y"]), sh["name"], f"{sh['tvd']} m TVD DF x {FT_PER_M}")
    # tubing
    for t in S.get("tubing", []):
        name, size, wt, grade, y, part = t[:6]
        bot_given = t[6] if len(t) > 6 else None
        notes = ["innerDiameter: header size has no ID/OD designation -> recorded as ID (user-defined rule)"]
        est = ["innerDiameter"]
        top = None
        string = part if part in ("Short", "Long") else None
        if part in ("upper", "Short", "Long") and prim == "MDTHF":
            top = (0.0, where(y), "depth reference m MDTHF", "ASSUMED 0 m MDTHF")
            est.append("topMd")
            notes.append("ASSUMED: top = 0 m MDTHF (depth reference is the tubing hanger flange)")
        elif part in ("upper",):
            notes.append(f"top not set: depth reference is {prim}, not the tubing hanger (0 would not be supported)")
        else:
            notes.append("inner string top not printed")
        bot = None
        if bot_given is not None:
            bot = (bot_given, where(y), name, f"{S['datum']}; tubing DEPTH (M) column of the report")
        else:
            end = ends.get(part)
            if end is not None and end["topMdM"] is not None:
                bot = (end["topMdM"], f"component {end['id']}", end["name"], f"{S['datum']}; depth of the end-of-string item")
                est.append("bottomMd")
                notes.append(f"DERIVED: bottom = depth of the string-end item '{end['name']}' ({end['id']})")
            else:
                w.issue("Missing", where(y), f"tubing '{name}' ({part}): no end-of-string depth -> bottomMd empty")
        cid = w.cid(("" if not string else string[0].lower() + "-") + ("inner-string-tubing" if part == "inner" else "tubing"))
        add_comp(w, cid=cid, typ="tubing", name=name, name_src=(where(y), name, S["method"]), string=string, top=top, bot=bot,
                 idd=(size, where(y), name, "USER-DEFINED RULE: size without ID/OD designation -> innerDiameter"),
                 weight=(wt, where(y), name) if wt else None, grade=(grade, where(y), name) if grade else None, est=est, notes=notes)
    if S.get("screen"):
        name, size, wt, y = S["screen"]
        add_comp(w, cid=w.cid("sand-screen"), typ="sand-screen", name=name, name_src=(where(y), name, S["method"] + " (COMPLETION TUBING header)"),
                 idd=(size, where(y), name, "USER-DEFINED RULE: size without ID/OD designation -> innerDiameter"), weight=(wt, where(y), name),
                 est=["innerDiameter"], notes=["innerDiameter: '4-1/2\"' has no ID/OD designation -> recorded as ID (user-defined rule)",
                                                "screen interval not printed -> topMd/bottomMd empty (not derived from the zone list)"])
        w.issue("Missing", where(y), f"'{name}': screen top/bottom depths not printed")
    for name, size, y, note in S.get("screens", []):
        add_comp(w, cid=w.cid("sand-screen"), typ="sand-screen", name=name, name_src=(where(y), name, S["method"]),
                 idd=(size, where(y), name, "USER-DEFINED RULE: size without ID/OD designation -> innerDiameter"),
                 est=["innerDiameter"], notes=["innerDiameter: '5-1/2\"' has no ID/OD designation -> recorded as ID (user-defined rule)",
                                                f"position: {note}; depths not printed -> empty"])
    if S.get("xmas"):
        name, y = S["xmas"]
        add_comp(w, cid=w.cid("christmas-tree"), typ="christmas-tree", name=name, name_src=(where(y), name, S["method"]))
    if S.get("pbtd"):
        m, datum, y, orig = S["pbtd"]
        est = [] if datum in ("MD-THF", "MDTHF") and prim == "MDTHF" else ["topMd"]
        add_comp(w, cid=w.cid("pbtd"), typ="pbtd", name=orig, name_src=(where(y), orig, S["method"]), top=(m, where(y), orig, f"m {datum} as printed"),
                 est=est, notes=[] if not est else [f"DATUM: PBTD printed in {datum}, equipment datum {prim}; not shifted"])
    if S.get("extra_pbtd_bp"):
        name, m, y, orig = S["extra_pbtd_bp"]
        add_comp(w, cid=w.cid("bridge-plug"), typ="bridge-plug", name=name, name_src=(where(y), orig, S["method"]), top=(m, where(y), orig, "m-MDTHF as printed"),
                 src_extra=["note 'After Workover All completion in 7\" Liner Section was left in hole as it is.'; 'NOTE: SLB FISH HERE'"])
    if S.get("cement_top"):
        m, y, orig = S["cement_top"]
        w.plug_cement = {"name": orig, "topMd": ft(m), "bottomMd": None}
        w.tr("plugBackCement[0].topMd", ft(m), where(y), f"{m} {orig}", f"{m} m x {FT_PER_M}; bottom not printed (bridge plug #1 at 1528.5 m below) — not derived")
        w.issue("Unclear", where(y), f"'Top of Cement' {m} m-MDTHF on the long-string side between bridge plugs #2 (1388.0) and #1 (1528.5) -> plugBackCement topMd only; bottom left empty")
    for zone, top, bot, orig, y, datum in S.get("zones", []):
        est, notes = [], []
        if not S.get("zones_are_perfs"):
            est.append("type")
            closest = "reservoir zone interval opposite the open-hole screens / zonal isolation (no perforation printed); Perforation is the nearest interval type"
        else:
            closest = None
        if datum != prim:
            est += ["topMd", "bottomMd"]
            notes.append(f"DATUM: interval printed in {datum}; equipment depths are {prim}; not shifted")
        d = add_comp(w, cid=w.cid(f"zone-{zone}"), typ="perforation", name=f"{zone} {top}-{bot} m {datum}", zone=zone,
                     name_src=(where(y), orig, S["method"] + "; DERIVED name: zone + interval"),
                     top=(float(top), where(y), orig, f"m {datum} as printed"), bot=(float(bot), where(y), orig, f"m {datum} as printed"),
                     closest=None, est=[e for e in est if e != "type"], notes=notes)
        if "type" in est:
            d["estimated"] = True
            d["estimatedFields"] = ["type"] + d.get("estimatedFields", [])
            d["estimationNote"] = "type: " + closest + ("; " + d["estimationNote"] if d.get("estimationNote") else "")
            d["actualType"] = d["name"]
    for it in S.get("interventions", []):
        text, m, datum, obj, how = it[:5]
        string = it[5] if len(it) > 5 else None
        typ, closest, unres = classify(text)
        top, extra, est = None, [how], []
        if m is not None and datum == "MDTHF" and prim == "MDTHF":
            top = (float(m), f"slide {s} obj #{obj}", text, "m MDTHF as printed in the note")
        elif m is not None:
            extra.append(f"depth {m} m {datum} in the note is NOT mapped to topMd (datum {datum}, not {prim})")
            w.issue("Unclear", f"slide {s} obj #{obj}", f"'{text}': depth {m} m {datum} not mapped (datum differs)")
        if "CONFLICT" in how:
            w.issue("Conflict", f"slide {s} obj #{obj}", how)
        add_comp(w, cid=w.cid(("" if not string else string[0].lower() + "-") + typ.replace(" (", "-").replace(")", "")), typ=typ, name=text.strip(),
                 name_src=(f"slide {s} obj #{obj}", text, "native slide text (note over the image)"), string=string, top=top,
                 closest=closest, unresolved=unres, est=est, notes=["intervention recorded by note; status as of the note date"], src_extra=extra)
    if S.get("float_shoe"):
        name, m, y = S["float_shoe"]
        w.float_equipment = [{"casingId": None, "md": ft(m), "type": "float shoe", "name": name}]
        w.tr("floatEquipment[0].md", ft(m), where(y), f"{m} {name}", f"{m} m x {FT_PER_M} ({S['datum']}); keys from the Float_Equipment sheet (Casing_ID, MD, Type, Name)")
        w.issue("Unclear", where(y), f"'{name}' {m} m mapped to floatEquipment (no Type_Key for a float shoe); JSON key names follow the template workbook "
                f"sheet Float_Equipment because the JSON template's floatEquipment list is empty; casingId left null (shoe is on the open-hole screen/inner string)")
    # slide notes over the image (not schematic content)
    for o in dump["objects"]:
        t = (o.get("text") or "").strip()
        if not t or t == "Open":
            continue
        if any(t.strip() == (i[0]).strip() for i in S.get("interventions", [])):
            continue
        w.unmapped.append((f"slide {s} obj #{o['id']}", re.sub(r"\s+", " ", t), "slide note over the image (status / HUD / history) — not equipment"))
        if o.get("table"):
            pass
    for o in dump["objects"]:
        if o.get("table"):
            w.unmapped.append((f"slide {s} obj #{o['id']}", " | ".join(" / ".join(r) for r in o["table"]), "slide table (survey / job history)"))
    if s == 31:
        w.issue("Conflict", "slide 31 obj #65543/#65545", "slide notes label zones 'E 12 A/B' and 'E 14 B' (OPEN 15/9/16, 16/9/16) while the report lists E10-11 and "
                "E12-13 — not merged; notes listed as unmapped")
    if s == 16:
        w.issue("Conflict", "slide 16 obj #5", "note 'Bull Plug @ 1516m-MDTHF (30/10/2022)' vs diagram bull plug row 1533 (m MD RKB) — different datums/dates; diagram value kept, note recorded")


def post_checks(wells):
    for w, doc in wells:
        if getattr(w, "plug_cement", None):
            doc["plugBackCement"] = [w.plug_cement]
