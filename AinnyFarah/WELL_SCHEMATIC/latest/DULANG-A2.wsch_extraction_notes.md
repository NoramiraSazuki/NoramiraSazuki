# DULANG-A2.wsch.json — extraction notes

Input: `Dulang A Well Schematic updated.pptx` slide 3 (read only, SHA-256 a5ca9a3b0076122633de31730fc97b5b8e3185d813c93bdcfdb3ad56bef2861b)
Template: `TEST1-DULANG-A4ST1.wsch.json` (read only, SHA-256 e94406028505daab795651226afabc0dcce72083fd5facac8e2016e605a906fa)
Well names: `EDAFY-Data_Dulang_Well.xlsx` (read only, SHA-256 eb874ebc996350ae1d25f3cb1020d6668d7887418d1066788473d04030fa3dac)

## How the slide was read
- Slide 3 holds two screenshots of the Freelance diagram 'DULANG WELL COMPLETION DIAGRAM — WELL A-2: DUAL OIL PRODUCER WITH LONG STRING SELECTIVE (TYPE 6)' (image6.png top half, image7.png bottom half, title bar 'Dulang-A Updated March 2007'). Equipment names, depths and MIN IDs are a visual transcription of these screenshots (clear, fully legible print).
- Later updates are text boxes placed over the screenshots; they were read from the slide XML and matched to equipment by position: ['FB – S3', 'OPEN  / Oct 2016', 'CLOSE / Oct 2016', 'OPEN / June 2012', 'No clear GLV data', 'E2 - 1906.4 – 1910.4  / mdTHF / E7 -  1931 - 1944', 'E12/13 – 1990 – 1995 / 1998 - 2001', 'E14 – 2015– 2025', 'CLOSED (March 2015)', '10', '2.31” CSA Plug (Feb 2023)', '2.25” CA Plug @ No-Go', 'CLOSED', 'GLV 12/64” (Apr/08) / DV (Jun/06) / DV (Jun/06) / GLOV 12/64” (Apr/08)', 'Open'].

## Datum, units and conflicts
- Well name: slide 'WELL A-2' -> EDAFY 'DULANG-A2' (UWI 1000000721; EDAFY base depth 2150 m, TVD 1384.5 m, DF 30.1 m).
- Equipment depths are printed in **m MDTHF** (column headers 'DEPTH m MDTHF'; perforation box 'mdTHF'). Casing shoes and PBTD are printed in **m MDDF**.
- The stored datum is DF (as in the template). THF depths were shifted to DF with the header 'RTE TO TUBING HANGER : 14.13 m' (MDDF = MDTHF + 14.13). topMdM / bottomMdM hold the resulting MDDF metres; the printed MDTHF value is quoted in every source text.
- displayDatum is **THF** with dfToThf = 14.13 m = 46.358 ft, so the viewer shows the depths exactly as printed on the slide.
- Depths converted m -> ft (x 3.28084, 3 dp); diameters in inches.
- totalDepth: not printed on the slide -> null (EDAFY base depth 2150 m = 9-5/8" shoe depth).
- OD (MAX OD) is not printed on this diagram -> outerDiameter null throughout. Casing IDs are not printed -> null.
- No completion fluid printed -> '-'. X-MAS TREE 'INGRAM CACTUS, DUAL, 3M' stored as wellhead.

## Status / update information (from the overlay text boxes)
- Short string: 2.31” CSA Plug (Feb 2023) in the TRDP-4A SCSSV; SPMs 'No clear GLV data'; WPB-2DE SSD 1882.1 'CLOSED (March 2015)'; 'D' no-go nipple 'OPEN June 2012'.
- Long string GLVs (top to bottom): GLV 12/64” (Apr/08), DV (Jun/06), DV (Jun/06), GLOV 12/64” (Apr/08); WPB-2DE SSD 1881.6 'CLOSED'; 'WPB-2DE' SSD 1982.6 'OPEN Oct 2016'; 2.25” CA Plug @ No-Go 2012.6 with 'CLOSE Oct 2016'.
- Statuses and valve types are appended to the component name in brackets because the format has no status field.

## Information with no slot in the template format
- '1/4" CONTROL LINE' on both strings (no depth).
- 'FB – S3' (fault block label) and page number '10'.
- SAND labels: 'SAND : E-2 & E-7', 'SAND : E-12/13', 'SAND : E-14'; '9 5/8" CASING SHOE' drawn below the E-14 sand.
- 'DOGLEG (DEPTH) : 103.24, 201.0 - 229.0, 288.0 - 574.0, 812.0 - 870.0, 1523.7, 1581.3' (no survey MD/inc/azimuth, so surveys stays empty).
- Seal bore extension has no printed depth (topMd null).
- Top blast joint (1902.6 - 1914.8) has no MIN ID printed on its row.
- Each component carries 'string' (Short / Long / Short & Long) as in TEST1-DULANG-A4ST1.
