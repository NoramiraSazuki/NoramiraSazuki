# Bo / reservoir oil volume / stock-tank oil volume check — MFL 20130041 DULANG-D32 Water Analysis (May 2013)

Source (read only): `/Volumes/homes/public/DulangOSS/20260925/DOCUMENTS/PVT Fluid Analysis Document/MFL 20130041_DULANG-D32_Water Analysis_May 2013.xls`
SHA-256 before and after: `31f4d0f3c0dab77b052ea3529d9c83a9494b8fef9be95002621015ef970ec918` (unchanged). Processed from a local copy in `_work/src/`.

Result: **no Bo, oil volume at reservoir conditions, or oil volume at stock-tank conditions for DLD-32L.** No Excel output created.

What was checked:
- All 19 worksheet cells (none hidden): water analysis only — 2 bottomhole water samples (2853-S1-F, 5099-S1-F, 1850 m MD, 1174 psig, 170.6 °F), flashed-gas compositions, GWR 49.2 / 48.4 scf/bbl, API water analysis, sample info.
- Text boxes (`_work/img/textboxes.json`): analysis-method text for water samples; "stocktank water" only.
- Embedded pictures, rendered and OCR'd (`_work/render/`): Core Lab logo, a second logo, Figure 1 flash-separation schematic. No values.
- Cell comments and the embedded CorelDRAW object: no relevant text.
- Raw workbook stream: Bo strings exist only in cached values of **external links** (BIFF CRN records, `_work/extcache.json`) carried over from the report template. They point to other companies' PVT jobs (e.g. Amerada Hess MFL 04248, Murphy Rompin-1 MPSR 0159, Murphy Endau-1 MPSR 1764, Addax 22017, job 818309, Bosi-5 ST1 CCE). They are not Dulang data and were not extracted.
- One footnote on sheet A.1, row 107 ("(1) These volumes refer to residual oil from the differential vaporisation.") is template leftover with no table or values.
