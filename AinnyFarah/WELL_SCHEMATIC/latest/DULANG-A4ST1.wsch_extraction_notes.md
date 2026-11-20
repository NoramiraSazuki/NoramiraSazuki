# DULANG-A4ST1.wsch.json — extraction notes

Input: `Dulang A04ST1 Well Schematic.xlsx` (read only, SHA-256 feac278f7db9c6737024e2ec3613cd5fb5070f5c159bf7d3bf7e9fc586f4900d)  
Template: `THF M DULANG-B30ST1-2-7.wsch 1.json` (read only, SHA-256 214c1e85160c4b581cae1406146fb538d6c9f3ec5413bafdba1d3a19c05c6257)  
Well names: `EDAFY-Data_Dulang_Well.xlsx` (read only, SHA-256 eb874ebc996350ae1d25f3cb1020d6668d7887418d1066788473d04030fa3dac)

## Information with no slot in the template format, conflicts and decisions

- Well name: input label 'WELL NAME: DULANG A004S01 OP4' (file 'Dulang A04ST1') -> EDAFY name 'DULANG-A4ST1' (UWI 1000023735; EDAFY base depth 3593 m, TVD 1505.61 m).
- Units: depths converted m -> ft (x 3.28084, 3 dp); topMdM / bottomMdM keep the printed metres; diameters in inches; source texts quote the original metres.
- TD conflict: header G10 '3700 m MDDF' vs diagram K190 3591 m MDDF (EDAFY base depth 3593 m). totalDepth uses the diagram value 3591 m = 11781.496 ft.
- PBTD conflict: header G11 '3675 m MDDF' vs diagram K187 3550 m MDDF (diagram used).
- Liner: header '2513 - 2591 m MDDF'; liner top matches diagram TOL 2513 m; bottom 2591 kept as printed (flagged; liner shoe is drawn at TD).
- Conductor and intermediate casing depths are 'xx m MDDF' placeholders -> bottomMd null.
- Location / slot: 'LOCATION / SLOT: DULANG A / 07' — no wellInfo field.
- Water depth: 77 m; drill floor height: 33 m-AMSL — no wellInfo field.
- DF to THF is not given in the input -> dfToThf null (enter it in the viewer to display in THF).
- Tubing joint calculations: AD13 =K71+S88 = 5018.123 m / AE13 419 jts (3-1/2"); AD14 =S163+12-S88 = 915.049 m / AE14 92 jts (2-7/8").
- Deviation column (degrees, J): [('J40', 42), ('J47', 79.2), ('J54', 80.7), ('J103', 74.5), ('J135', 74.4), ('J157', 74.4), ('J179', 74.4), ('J189', 76)] — no survey MD/azimuth, so surveys stays empty.
- Zones without an interval: 'E34B - BCO' (row 118, no depths); 'E51 (Future BCO)' with '1485.15 m - 1487.58 m TVD', '3517m - 3526 m MDDF (9 m)', pressure '2,220 psi' — future zone, not added as a perforation.
- Zone spacing notes: 'Distance 135m (E40A to E34A) ', 'Distance 130m', 'Distance 132 m'.
- Footer: 'Designed by: Hidayah / Mahirah' (Completion Engineer); 'Jagaan' (Production Technologist); 'Date: 6th May 2020', 'Rev: Final'.
- Pictures: logo and equipment symbol images (no data).
- Each component carries 'string' (Short / Long / Short & Long), as in the other Dulang JSON outputs; the B30ST1 template (single string) has no such key.
- Drawing text box at row 96: 'TOC in Annulus unknown (June 2022)' (no matching field).
- Drawing text box at row 107: 'Leak at Seal Assembly @ 3108m MDDF (June 2022) from WLD. Not isolated' (no matching field).
