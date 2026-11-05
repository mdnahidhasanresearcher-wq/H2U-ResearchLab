# Stage A Source-Evidence Pack

This pack holds the source-data foundation for Stage A. It is **transcription only**:
- no unit conversion;
- no per-tonne normalisation;
- no reconciliation;
- no scenario calculation;
- no R code.

| Path | Content | Status |
|---|---|---|
| `data/raw/source_reports/IEAGHG_2017-03.pdf` | IEAGHG Technical Report 2017/03 (159 pp.) | Original, unchanged, read-only |
| `data/raw/source_reports/H2FUTURE_D9.2.pdf` | Dowling, Janssen & Weeda (2022), H2FUTURE D9.2 V1.0 (46 pp.) | Original, unchanged, read-only |
| `data/raw/source_reports/IRENA_Bluerisk_2023.pdf` | IRENA & Bluerisk (2023), *Water for hydrogen production* (66 pp.) | Original, unchanged, read-only |
| `data/source_register/source_manifest.csv` | For each source: bibliographic reference, URL, pages, size, SHA-256, verification status. Also includes Jin et al. (2024), marked pending | Complete |
| `data/source_register/StageA_Source_Extraction_Register.xlsx` | Sheets: README, Source_Register, IEAGHG_Reference, H2FUTURE_Crosscheck, IRENA_Water, Boundary_Register, Unresolved_Issues | 540 records |
| `documentation/StageA_Methodological_Issue_Register.md` | Seven page-referenced methodological issues for supervisor discussion | Complete |
| `documentation/tools/build_stageA_source_register.py` | Rebuilds the manifest and workbook from the transcribed values. No calculations | Reproducible |

## Rules applied
- **IEAGHG:** Case 4A-1 only (ammonia–urea, no additional CCS).
- **H2FUTURE:** kept as a cross-check; none of its values are transferred into the IEAGHG case.
- **`reported_value`:** exactly as printed.
- **`value_numeric`:** only records how the printed number format was read. The IEAGHG heat-balance and utility tables use European format, so "4.490" kW means 4,490 kW and "7,00" MPa means 7.00.
- **Page numbers:** `pdf_page` and `printed_page_or_sheet` are recorded separately.
- **Blanks and dashes:** recorded as `unquantified`, never as zero.
- **"Carbon Not Emitted":** IEAGHG's label is recorded but not adopted as a classification.
- **Check performed:** all 472 numeric values were checked automatically. Each appears verbatim on its cited PDF page.

## Git and storage
- `.gitignore` excludes `data/raw/`, so the copyrighted PDFs are never pushed to the public repository.
- Nothing has been committed.
- The cloud workspace is temporary. Keep your own copy of `data/raw/source_reports/`, and of the other files until they are committed.

## Pending
- **Jin et al. (2024)**, *Clean Energy* 8(1):126–133, DOI 10.1093/ce/zkad072. Bibliographic record verified via Crossref; the PDF download was blocked by the site's bot check. Supply the PDF to complete full-text verification and extraction.
