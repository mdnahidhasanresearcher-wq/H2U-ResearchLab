"""
Build the Stage A Source-Extraction Register (transcription only).

Every value below was transcribed by hand from the original PDFs in
data/raw/source_reports/ and checked against the rendered page images.
No unit conversion, per-tonne normalisation, reconciliation or scenario
calculation is performed here. Values are stored exactly as printed
(reported_value); value_numeric only records how the printed number
format was read (e.g. European "4.490" kW = 4490 kW).

Run (from the repository root):  python3 -I documentation/tools/build_stageA_source_register.py
Requires: openpyxl (already present; nothing installed).
"""
import csv, hashlib, os, subprocess
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
RAW = os.path.join(ROOT, "data", "raw", "source_reports")
OUT_DIR = os.path.join(ROOT, "data", "source_register")
os.makedirs(OUT_DIR, exist_ok=True)

COLS = ["record_id", "parameter_id", "parameter_name", "reported_value", "value_numeric",
        "original_unit", "source_id", "pdf_page", "printed_page_or_sheet",
        "table_figure_equation", "reference_case", "data_year", "system_boundary",
        "evidence_status", "extraction_notes"]
ALLOWED_STATUS = {"theoretical", "sourced coefficient", "company observation",
                  "national statistic", "calculated", "assumed", "unquantified"}

# ----------------------------------------------------------------------------
# Sources
# ----------------------------------------------------------------------------
SOURCES = [
    dict(source_id="SRC_IEAGHG_2017_03", file="IEAGHG_2017-03.pdf",
         reference="IEAGHG (2017). Techno-Economic Evaluation of HYCO Plant Integrated to Ammonia / Urea or Methanol Production with CCS. IEAGHG Technical Report 2017/03, February 2017. Prepared by Amec Foster Wheeler and IEAGHG (G. Collodi, G. Azzaro, N. Ferrari; IEAGHG manager S. Santos). IEA Environmental Projects Ltd., Cheltenham.",
         url="https://ieaghg.org/publications/2017-03%20Techno-Economic%20Evaluation%20of%20HYCO%20Plant%20Integrated%20to%20Ammonia%20%20Urea%20or%20Methanol%20Production%20with%20CCS.pdf",
         url_note="Requested via https://publications.ieaghg.org/technicalreports/... which redirected to this URL (HTTP 200, application/pdf), 4 Nov 2026.",
         access="Free download; Copyright IEAGHG 2017, all rights reserved; PDF permissions block text copying (encrypted, copy:no).",
         status="Full text verified: opens; 159 pages; title page, report number, contractor report (Sheets 1-80), Annex I reference document (Sheets 1-24), Annexes II-IV and back cover present."),
    dict(source_id="SRC_H2FUTURE_D9_2_2022", file="H2FUTURE_D9.2.pdf",
         reference="Dowling, R.A., Janssen, J.L.L.C.C. and Weeda, M. (2022). Fertilizer industry exploitation study: Techno-economic analysis of electrolysis based ammonia production for the fertilizer industry. H2FUTURE Deliverable D9.2, V1.0 (Final), WP9/T9.3, TNO, delivered 8 March 2022. FCH 2 JU grant agreement No 735503. Ref. Ares(2022)1783613.",
         url="https://www.h2future-project.eu/media/bf1mpqss/d9-2_fertiliser-industry-exploitation-study.pdf",
         url_note="Downloaded 4 Nov 2026 (HTTP 200, application/pdf). PDF producer metadata shows the file was re-saved by the website (PDFsharp, 30 Jan 2026); content is V1.0 Final.",
         access="Public (dissemination level PU).",
         status="Full text verified: opens; 46 of 46 pages; document information, revision history, ToC, chapters 1-6 and bibliography present."),
    dict(source_id="SRC_IRENA_BLUERISK_2023", file="IRENA_Bluerisk_2023.pdf",
         reference="IRENA and Bluerisk (2023). Water for hydrogen production. International Renewable Energy Agency and Bluerisk, Abu Dhabi. ISBN 978-92-9260-526-1. Authors: E. Bianco (IRENA), T. Luo (Bluerisk), D. Nagpal (ex-IRENA).",
         url="https://www.irena.org/-/media/Files/IRENA/Agency/Publication/2023/Dec/IRENA_Bluerisk_Water_for_hydrogen_production_2023.pdf",
         url_note="Downloaded 4 Nov 2026 (HTTP 200, application/pdf). Landing page: https://www.irena.org/Publications/2023/Dec/Water-for-hydrogen-production",
         access="Open; may be used with acknowledgement of IRENA (third-party material may carry separate terms).",
         status="Full text verified: opens; 66 pages; glossary, chapters 1-5, references, Appendix Table A1 and back cover present. Printed page = PDF page - 2."),
    dict(source_id="SRC_JIN_2024_CLEANENERGY", file="(not available)",
         reference="Jin, X., You, S., Petersen, M., Riofrio, J., Thakur, S., Traeholt, C. and Feng, Z. (2024). Exploring commercial water electrolyser systems: a data-based analysis of product characteristics. Clean Energy, 8(1), 126-133. DOI: 10.1093/ce/zkad072. Published 10 Jan 2024. Licence (Crossref): CC BY 4.0.",
         url="https://academic.oup.com/ce/article-pdf/8/1/126/55378728/zkad072.pdf",
         url_note="Publisher and DTU Orbit (https://orbit.dtu.dk/files/351087174/...) both returned HTTP 403 Cloudflare bot challenge; not retried.",
         access="Open access (CC BY 4.0) per Crossref; PDF to be supplied by the researcher.",
         status="PENDING: full-text verification, page/table extraction, relevance check, and whether it gives system-level electricity coefficients or only product specifications."),
]

def pdf_pages(path):
    out = subprocess.run(["pdfinfo", path], capture_output=True, text=True).stdout
    for line in out.splitlines():
        if line.startswith("Pages:"):
            return int(line.split()[1])
    return None

def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()

for s in SOURCES:
    p = os.path.join(RAW, s["file"])
    if os.path.isfile(p):
        s["pages"], s["bytes"], s["sha256"] = pdf_pages(p), os.path.getsize(p), sha256(p)
        s["opens"] = "yes" if s["pages"] else "no"
        s["local_path"] = os.path.relpath(p, ROOT)
    else:
        s["pages"] = s["bytes"] = s["sha256"] = ""
        s["opens"] = "not available"
        s["local_path"] = ""

# ----------------------------------------------------------------------------
# Record helpers
# ----------------------------------------------------------------------------
records = {"IEAGHG_Reference": [], "H2FUTURE_Crosscheck": [], "IRENA_Water": []}
counters = {"IEAGHG_Reference": "IEA", "H2FUTURE_Crosscheck": "H2F", "IRENA_Water": "IRW"}

def add(sheet, pid, name, val, unit, src, pdfp, printed, tfe, case, year, boundary,
        status, notes, num=None):
    assert status in ALLOWED_STATUS, status
    rid = f"{counters[sheet]}-{len(records[sheet]) + 1:04d}"
    records[sheet].append(dict(record_id=rid, parameter_id=pid, parameter_name=name,
        reported_value=val, value_numeric=num, original_unit=unit, source_id=src,
        pdf_page=pdfp, printed_page_or_sheet=printed, table_figure_equation=tfe,
        reference_case=case, data_year=year, system_boundary=boundary,
        evidence_status=status, extraction_notes=notes))

def eu(s):
    """Read a European-formatted number exactly as printed: '.' thousands, ',' decimal."""
    s = s.strip()
    try:
        return float(s.replace(".", "").replace(",", "."))
    except ValueError:
        return None

def en(s):
    """Read an English-formatted number: ',' thousands, '.' decimal."""
    try:
        return float(s.strip().replace(",", ""))
    except ValueError:
        return None

# ============================================================================
# IEAGHG 2017-03 -- Case 4A-1 (ammonia-urea plant WITHOUT CO2 capture) ONLY
# ============================================================================
S, C, Y = "SRC_IEAGHG_2017_03", "Case 4A-1 (ammonia-urea, no additional CCS)", "2017 (report); H&MB and utility sheets dated Oct 2015"
A = lambda *a, **k: add("IEAGHG_Reference", *a, **k)

# --- Configuration (text evidence) ---
A("IEA_CFG_REFORMING", "Reforming configuration", "Two-step (tandem) reforming: SMR primary reformer + air-blown ATR secondary reformer", "text", S, 37, "Sheet 10 of 80", "Section 2.4.1", C, Y, "Syngas generation unit", "sourced coefficient", "Also stated on PDF p.5 (IEAGHG overview, Basis of Design) and PDF p.40 (Sheet 11). Do NOT simplify to an SMR-only plant.")
A("IEA_CFG_N2_SOURCE", "Nitrogen supply for ammonia synthesis", "Nitrogen is supplied by introducing process air into the secondary reformer", "text", S, 37, "Sheet 10 of 80", "Section 2.4.1", C, Y, "Syngas generation unit", "sourced coefficient", "Quote on p.37; confirmed p.40 (Sheet 11): 'The introduction of air in the secondary reformer provides the nitrogen required for the synthesis of ammonia.'")
A("IEA_CFG_AIR_FIXED", "Process-air quantity rule", "Since the molar ratio (H2/N2) are fixed to maintain a value as close to 3, the amount of process air introduced into the ATR is therefore fixed", "text", S, 40, "Sheet 11 of 80", "Section 2.4.1 Reforming Section", C, Y, "Secondary reformer", "sourced coefficient", "Methane slip from ATR is controlled by adjusting primary-reformer firing (same paragraph). Source does not describe operation at reduced reforming throughput.")
A("IEA_CFG_SYNGAS_RATIO", "Syngas H2:N2 molar ratio (typical)", "approximately 3:1", "molar ratio", S, 37, "Sheet 10 of 80", "Section 2.4.1", C, Y, "Synthesis gas", "sourced coefficient", "Descriptive statement, not a stream value.", 3)
A("IEA_CFG_SMR_FUEL", "Heat supply to primary reformer", "supplied indirectly by burning the tail gas and supplementary fuel in the SMR furnace", "text", S, 40, "Sheet 11 of 80", "Section 2.4.1 Reforming Section", C, Y, "SMR furnace", "sourced coefficient", "Fuel to burners = NG fuel (stream 3) + tail gas (stream 4); flash gas from the CO2-removal flash column is also sent to SMR burners (p.41) but is not quantified.")
A("IEA_CFG_SLIP_PRIMARY", "Methane slip, primary reformer", "up to 10", "%v", S, 40, "Sheet 11 of 80", "Section 2.4.1", C, Y, "Primary reformer outlet", "sourced coefficient", "", 10)
A("IEA_CFG_SLIP_ATR", "Methane slip, secondary reformer (ATR)", "less than 0.8", "%v", S, 40, "Sheet 11 of 80", "Section 2.4.1", C, Y, "ATR outlet", "sourced coefficient", "Upper bound, not a point value.", 0.8)
A("IEA_CFG_ATR_T", "ATR outlet gas temperature", "around 1000", "degC", S, 41, "Sheet 12 of 80", "Section 2.4.1", C, Y, "ATR outlet", "sourced coefficient", "Approximate.", 1000)
A("IEA_CFG_CO_CONV", "CO converted to CO2 in HTS+LTS", "~95", "% of CO fed", S, 41, "Sheet 12 of 80", "Shift Section", C, Y, "Shift reactors", "sourced coefficient", "Approximate.", 95)
A("IEA_CFG_CO2_REMOVAL", "Bulk CO2 removal technology and destination", "aMDEA absorption; CO2 released in regeneration column is sent to the Urea Plant as feedstock; flash-column volatiles sent to SMR burners", "text", S, 41, "Sheet 12 of 80", "Bulk CO2 Removal Section", C, Y, "CO2 removal unit", "sourced coefficient", "Flash-gas flow and composition not reported.")
A("IEA_CFG_SYNGAS_PURITY", "Purified syngas to synthesis loop: inerts and oxygenates", "mostly N2 and H2 with approximately 1% (mol) of Ar and CH4 and less than 5 ppm (mol) of CO and CO2", "text", S, 42, "Sheet 13 of 80", "Methanation Section", C, Y, "Methanator outlet / make-up gas", "sourced coefficient", "Compare stream 7 composition on PDF p.33.")
A("IEA_SPEC_SYNGAS_COX", "Syngas to ammonia synthesis specification, CO2 + CO", "< 5", "ppm (%v dry basis)", S, 28, "Sheet 7 of 80", "Section 2.1.3 Syngas Specifications", C, Y, "Syngas to ammonia synthesis", "sourced coefficient", "Any hydrogen entering the synthesis loop must meet this oxygenate limit; p.42 states oxygen-containing compounds are severe catalyst poisons.", 5)
A("IEA_CFG_LOOP_P", "Ammonia synthesis operating pressure (converter inlet)", "between 130 and 220", "bar(g)", S, 42, "Sheet 13 of 80", "Ammonia Synthesis Section", C, Y, "Synthesis loop", "sourced coefficient", "Generic range for modern plants, not a Case 4A-1 stream value.")
A("IEA_CFG_LOOP_CONV", "Single-pass conversion of N2 and H2 in converter", "about 25 - 30", "%", S, 42, "Sheet 13 of 80", "Ammonia Synthesis Section", C, Y, "Synthesis converter", "sourced coefficient", "Range.")
A("IEA_CFG_PURGE", "Purge gas system destinations", "Purge gas to ammonia and hydrogen recovery unit; recovered NH3 to loop; recovered H2 used as feedstock to the hydrogenation unit; tail gas to SMR as fuel", "text", S, 43, "Sheet 14 of 80", "Purge Gas System", C, Y, "Synthesis loop / HRU", "sourced coefficient", "BFD (p.31) also labels 'Recycle H2 to NG feed' and 'Hydrogen make-up to NH3 synthesis' leaving the HRU; these internal flows are NOT quantified in the H&MB (p.33).")
A("IEA_HRU_FLOWS", "HRU recovered-hydrogen recycle flows", "", "", S, 31, "BFD Sheet 01 of 02", "Block flow diagram", C, Y, "Hydrogen recovery unit", "unquantified", "Shown on BFD but no flow given anywhere inspected. Do not assume zero.")
A("IEA_FLASHGAS", "CO2-removal flash gas to SMR burners", "", "", S, 41, "Sheet 12 of 80", "Bulk CO2 Removal Section", C, Y, "CO2 removal unit", "unquantified", "Described, not quantified.")

# --- Capacities / products ---
A("IEA_NH3_CAPACITY", "Ammonia plant design capacity (liquid ammonia)", "1,350", "t/d", S, 27, "Sheet 6 of 80", "Section 2.1.1 Capacity", C, Y, "Ammonia plant (gross ammonia)", "sourced coefficient", "Design capacity. Same value given as 'nominal 1350 t/d' on PDF p.4.", 1350)
A("IEA_UREA_CAPACITY", "Urea plant design capacity, Case 4A-1", "2,260", "t/d", S, 27, "Sheet 6 of 80", "Section 2.1.1 Capacity", C, Y, "Urea plant (granulated urea)", "sourced coefficient", "Case 4A-2 (2,380 t/d) deliberately NOT transcribed.", 2260)
A("IEA_NH3_TO_UREA_SHARE", "Ammonia conversion to urea, Case 4A-1", "about 95", "%", S, 27, "Sheet 6 of 80", "Section 2.1.1 Capacity", C, Y, "Ammonia-to-urea", "sourced coefficient", "'corresponding to about 95% ammonia conversion' (approximate).", 95)
A("IEA_NH3_SPEC", "Ammonia product purity (min)", "99.5", "wt. %", S, 27, "Sheet 6 of 80", "Section 2.1.2 Product Specifications", C, Y, "Ammonia product", "sourced coefficient", "Water max 0.5 wt%, oil max 5 ppm wt.", 99.5)
A("IEA_UREA_SPEC_N", "Urea total nitrogen (min)", "46", "wt. %", S, 27, "Sheet 6 of 80", "Section 2.1.2 Product Specifications", C, Y, "Urea product", "sourced coefficient", "", 46)
A("IEA_NH3_TO_UREA_COND", "Ammonia sent to urea plant: conditions", "10-15 degC and 10-20 Bar(g)", "degC; bar(g)", S, 44, "Sheet 15 of 80", "Section 2.4.1 end", C, Y, "Ammonia to urea", "sourced coefficient", "Storage alternative: cryogenic atmospheric tanks at -33 degC.")
A("IEA_CO2_UREA_PURITY", "CO2 feed to urea: purity and pressure", "purity of at least 98.5% vol. and at 0.50 to 0.75 Bar(g); compressed to about 160 Bar(g)", "% vol; bar(g)", S, 44, "Sheet 15 of 80", "Section 2.4.2 Urea Plant", C, Y, "CO2 to urea", "sourced coefficient", "Purity basis (dry/wet) not stated; stream 8 on p.33 is wet (H2O 0,0510 mol/mol). 'coming mainly from the Bulk CO2 Removal Section'.")
A("IEA_UREA_REACTOR", "Urea synthesis reactor conditions", "around 190 degC and 150-160 Bar(g)", "degC; bar(g)", S, 44, "Sheet 15 of 80", "Section 2.4.2 Urea Plant", C, Y, "Urea synthesis", "sourced coefficient", "'Depending on the licensor'.")

# --- Plant performance data (Sheet 22) -- English number format (dot decimal, comma thousands)
perf = [
    ("IEA_PD_NG_FEED", "Natural Gas (as Feedstock)", "27.021", "t/h", "NG feedstock", "Equals stream 2 mass flow 27021 kg/h (p.33)."),
    ("IEA_PD_NG_FUEL", "Natural Gas (as Fuel)", "10.630", "t/h", "NG fuel (all fuel uses combined)", "Single fuel figure; no split into reformer fuel vs other utility fuel. Equals stream 3 'Natural Gas fuel to burners' 10630 kg/h (p.33)."),
    ("IEA_PD_NG_TOTAL", "Natural Gas (Total Consumption)", "37.651", "t/h", "NG at plant battery limit", "Equals stream 1 37651 kg/h (p.33)."),
    ("IEA_PD_NG_LHV", "Natural Gas LHV", "46.50", "MJ/kg", "NG", "Annex I Table 1 gives 46.502 (p.128)."),
    ("IEA_PD_ENERGY_IN", "Total Energy Input", "486.33", "MW", "NG feed + fuel (LHV)", "Basis LHV implied by preceding row; not restated."),
    ("IEA_PD_UREA_TPD", "Urea Product to BL", "2,260", "t/d", "Urea product", ""),
    ("IEA_PD_UREA_TPH", "Urea Product to BL", "94.17", "t/h", "Urea product", "Equals stream 11 94170 kg/h."),
    ("IEA_PD_UREA_LHV", "Urea LHV", "32.70", "MJ/kg", "Urea product", ""),
    ("IEA_PD_UREA_ENERGY", "Total Energy in the Urea Product", "855.4", "MW", "Urea product", ""),
    ("IEA_PD_NH3_TPD", "Ammonia Product to BL", "68.4", "t/d", "Ammonia product (not converted to urea)", "Case 4A-1 only ('applicable to the base case only', p.28)."),
    ("IEA_PD_NH3_TPH", "Ammonia Product to BL", "2.85", "t/h", "Ammonia product", "Equals stream 9 2850 kg/h."),
    ("IEA_PD_NH3_LHV", "Ammonia LHV", "18.60", "MJ/kg", "Ammonia product", ""),
    ("IEA_PD_NH3_ENERGY", "Total Energy in the Ammonia Product", "14.73", "MW", "Ammonia product", ""),
    ("IEA_PD_PWR_NH3UREA", "Power balance: Ammonia / Urea Production", "-4.490", "MWe", "Ammonia + urea process units", "Negative = consumption."),
    ("IEA_PD_PWR_STEAM", "Power balance: Steam + BFW System", "-1.100", "MWe", "Steam and BFW system", ""),
    ("IEA_PD_PWR_UTIL", "Power balance: Utilities + BoP", "-3.300", "MWe", "Utilities and balance of plant", ""),
    ("IEA_PD_PWR_IMPORT", "Imported Power from the Grid", "8.890", "MWe", "Plant battery limit (electricity in)", "No electricity export reported. Steam turbines drive major compressors (p.47-48), so this is NOT total shaft-power demand."),
    ("IEA_PD_SPEC_FEED", "Specific: Natural Gas (as Feedstock)", "13.343", "GJ/t urea", "NG feedstock per t urea", "Reported per tonne UREA, not per tonne ammonia."),
    ("IEA_PD_SPEC_FUEL", "Specific: Natural Gas (as Fuel)", "5.249", "GJ/t urea", "NG fuel per t urea", ""),
    ("IEA_PD_SPEC_FEEDFUEL", "Specific: Feed + Fuel", "18.592", "GJ/t urea", "NG total per t urea", ""),
    ("IEA_PD_SPEC_CO2", "Specific CO2 Emission (Direct)", "0.3250", "t/t urea", "Direct (flue gas) CO2 per t urea", ""),
    ("IEA_PD_EQ_CO2_UREA", "Equivalent CO2 in the Urea Product", "69.3", "%", "Share of input carbon leaving in urea", "Per source definition; see CO2 balance p.52."),
]
for pid, name, v, u, b, n in perf:
    A(pid, name, v, u, S, 53, "Sheet 22 of 80", "Section 2.6 Plant Performance Data, 'Base Case' column", C, Y, b, "sourced coefficient",
      ("English number format on this page. " + n).strip(), en(v))

# --- CO2 balance (Sheet 21) ---
co2 = [
    ("IEA_CB_IN_FEED", "CO2-equivalent carbon in: Natural gas feedstock", "1626", ""),
    ("IEA_CB_IN_FUEL", "CO2-equivalent carbon in: Natural gas fuel", "637", ""),
    ("IEA_CB_IN_TOTAL", "CO2-equivalent carbon in: TOTAL (IN)", "2263", ""),
    ("IEA_CB_OUT_UREA", "Total carbon in urea product", "1568", "Carbon carried in product. Same value as stream 11 urea molar flow (1568 kmol/h)."),
    ("IEA_CB_SUBTOTAL_NOT_EMITTED", "Source row 'SUB-TOTAL (Carbon Not Emitted)'", "1568", "SOURCE LABEL NOT ADOPTED: 'Carbon Not Emitted' is not a permissible classification in this research. Record only as carbon carried in urea product; not storage, not avoided emission."),
    ("IEA_CB_OUT_FLUE", "Flue gas to stack (A)", "695", "Direct emission point."),
    ("IEA_CB_OUT_EMISSION", "Emission", "695", ""),
    ("IEA_CB_OUT_TOTAL", "TOTAL (OUT)", "2263", ""),
]
for pid, name, v, n in co2:
    A(pid, name, v, "kmol/h (equivalent flow of CO2)", S, 52, "Sheet 21 of 80", "Section 2.5 CO2 balance table, 'Ammonia/Urea Plant w/o CCS' column", C, Y, "Whole complex carbon balance", "sourced coefficient", n, en(v))
A("IEA_CB_OUT_VENTS", "Vents", "-", "kmol/h (equivalent flow of CO2)", S, 52, "Sheet 21 of 80", "Section 2.5 CO2 balance table", C, Y, "Whole complex carbon balance", "unquantified", "Printed as a dash. NOT assumed zero. H2FUTURE p.21 states that in the IEAGHG study all CO2 is used for urea and none is sent to stack (secondary interpretation).")
A("IEA_CB_EQ_UREA_PCT", "Equivalent CO2 in Urea product (%)", "69.3%", "%", S, 52, "Sheet 21 of 80", "Section 2.5 CO2 balance table", C, Y, "Whole complex", "sourced coefficient", "", 69.3)
A("IEA_CB_INDIRECT", "Indirect CO2 Emission* (B) (imported power)", "70 - 151", "kmol/h (equivalent flow of CO2)", S, 52, "Sheet 21 of 80", "Section 2.5 CO2 balance table + footnote", C, Y, "Grid electricity (outside plant)", "sourced coefficient", "Range from assumed grid factors 348 kg/MWh (NG) and 746 kg/MWh (coal). Outside Stage A scope (base case uses dedicated renewables for electrolysis); recorded for completeness only.")
A("IEA_CB_TOTAL_EMISSION", "TOTAL CO2 EMISSION - (A) + (B)", "765 - 846", "kmol/h (equivalent flow of CO2)", S, 52, "Sheet 21 of 80", "Section 2.5 CO2 balance table", C, Y, "Plant + grid", "sourced coefficient", "Mixes direct and indirect; do not use as a Stage A emission total.")

# --- Heat and Material Balance, Case 4A-1 (PDF p.33) -- European format
hmb_cols = ["Temperature", "Pressure", "Molar Flow", "Mass Flow"]
hmb_units = ["degC", "MPa", "kmol/h", "kg/h"]
comps = ["CO2", "CO", "Hydrogen", "Nitrogen", "Ar", "Oxygen", "Methane", "Ethane", "Propane", "n-Butane", "n-Pentane", "H2O", "NH3", "Urea"]
NG = ["0,0200", "0,0000", "0,0000", "0,0089", "0,0000", "0,0000", "0,8900", "0,0700", "0,0100", "0,0010", "0,0001", "0,0000", "0,0000", "0,0000"]
Z = ["0,0000"] * 14
def comp(**kw):
    c = list(Z)
    for k, v in kw.items():
        c[comps.index(k.replace("_", "-"))] = v
    return c
streams = [
    (1, "Natural Gas From B.L.", ["9", "7,00", "2090", "37651"], NG, "NG at battery limit (feed + fuel)"),
    (2, "Natural Gas feedstock to Ammonia Plant", ["15", "4,20", "1500", "27021"], NG, "NG feedstock account"),
    (3, "Natural Gas fuel to burners", ["15", "0,15", "590", "10630"], NG, "NG fuel to (SMR) burners"),
    (4, "Tail gas to burners", ["50", "0,50", "282", "5583"], comp(Hydrogen="0,2690", Nitrogen="0,4770", Ar="0,0760", Methane="0,1780"), "HRU tail gas to SMR burners (internal fuel)"),
    (5, "Flue gas to stack", ["155", "atm", "7969", "221304"], comp(CO2="0,0870", CO="(2)", Nitrogen="0,7070", Ar="0,0120", Oxygen="0,0140", H2O="0,1800"), "Direct emission point"),
    (6, "Process Air", ["Ambient", "atm", "2286", "66079"], comp(Nitrogen="0,7800", Ar="0,0100", Oxygen="0,2100"), "Process air to secondary reformer = nitrogen source"),
    (7, "Syngas to compression", ["43", "3,00", "7007", "62853"], comp(Hydrogen="0,7300", Nitrogen="0,2580", Ar="0,0030", Methane="0,0070", H2O="0,0020"), "Candidate common hydrogen-delivery point (make-up syngas after methanation, before syngas compressor)"),
    (8, "CO2 to Urea Plant", ["43", "0,15", "1668", "70593"], comp(CO2="0,9400", Hydrogen="0,0080", Nitrogen="0,0010", H2O="0,0510"), "Process CO2 delivered to urea"),
    (9, "Ammonia to storage", ["-33", "atm", "167,0", "2850"], comp(NH3="1,0000"), "Ammonia product (export/storage)"),
    (10, "Ammonia to Urea Plant", ["10", "2,10", "3136", "53400"], comp(NH3="1,0000"), "Ammonia to urea"),
    (11, "Urea product", ["Ambient", "atm", "1568", "94170"], comp(Urea="1,0000"), "Urea product"),
]
for sn, desc, cond, cmp_, bnd in streams:
    for col, unit, v in zip(hmb_cols, hmb_units, cond):
        note = "European format verified on rendered page (decimal comma)."
        if v in ("atm", "Ambient"):
            note = "Printed as text; no numeric value given."
        if col == "Pressure" and v not in ("atm", "Ambient"):
            note += " Absolute vs gauge not stated in table."
        A(f"IEA_HMB_S{sn:02d}_{col.replace(' ', '_').upper()}", f"Stream {sn} '{desc}': {col}", v, unit, S, 33, "H&MB drawing (no sheet no.); follows Sheet 9", "Heat and Material Balance, Ammonia-Urea plant without CO2 capture", C, Y, bnd, "sourced coefficient", note, eu(v) if v not in ("atm", "Ambient") else None)
    for cname, v in zip(comps, cmp_):
        note = "European format (decimal comma). Value printed in table." if v != "(2)" else "Printed '(2)': note (2) = 30 mg/Nm3 max (CO in flue gas)."
        A(f"IEA_HMB_S{sn:02d}_X_{cname.replace('-', '').upper()}", f"Stream {sn} '{desc}': mole fraction {cname}", v, "mol/mol", S, 33, "H&MB drawing (no sheet no.)", "Heat and Material Balance composition block", C, Y, bnd, "sourced coefficient", note, eu(v) if v != "(2)" else None)
A("IEA_HMB_S12", "Stream 12 column", "(blank)", "", S, 33, "H&MB drawing", "Heat and Material Balance", C, Y, "Unknown", "unquantified", "Column 12 is printed empty. Not interpreted.")
A("IEA_HMB_H2S", "H2S in natural gas (design assumption)", "5", "ppmv", S, 33, "H&MB note (1)", "Heat and Material Balance note", C, Y, "NG feed", "sourced coefficient", "Design-purpose assumption; also Annex I Table 1 (p.128).", 5)

# --- Utility consumption table, Case 4A-1 (PDF p.55) -- European format
util = [
    ("IEA_UT_PWR_NH3UREA", "Electric power: Ammonia and Urea Plant", "4.490", "kW", "Ammonia & urea plant", 4490),
    ("IEA_UT_PWR_STEAM", "Electric power: Steam and BFW Plant", "1.100", "kW", "Steam & BFW plant", 1100),
    ("IEA_UT_PWR_UTIL", "Electric power: Utilities / BoP", "3.300", "kW", "Utilities / BoP", 3300),
    ("IEA_UT_PWR_TOTAL", "Electric power: TOTAL", "8.890", "kW", "Whole complex", 8890),
    ("IEA_UT_MPSTEAM_NH3", "MP steam: Ammonia and Urea Plant (consumption)", "175,0", "t/h", "Ammonia & urea plant", 175.0),
    ("IEA_UT_MPSTEAM_STEAM", "MP steam: Steam and BFW Plant (production)", "-175,0", "t/h", "Steam & BFW plant", -175.0),
    ("IEA_UT_LOSS_NH3_A", "Losses: Ammonia and Urea Plant (row 1)", "28,0", "t/h", "Ammonia & urea plant", 28.0),
    ("IEA_UT_LOSS_NH3_B", "Losses: Ammonia and Urea Plant (row 2)", "-38,7", "t/h", "Ammonia & urea plant", -38.7),
    ("IEA_UT_LOSS_STEAM", "Losses: Steam and BFW Plant", "-6,3", "t/h", "Steam & BFW plant", -6.3),
    ("IEA_UT_LOSS_TOTAL", "Losses: TOTAL", "-17,0", "t/h", "Whole complex", -17.0),
    ("IEA_UT_DMW_NH3", "DMW: Ammonia and Urea Plant", "-164,3", "t/h", "Ammonia & urea plant", -164.3),
    ("IEA_UT_DMW_STEAM_A", "DMW: Steam and BFW Plant (row 1)", "284,3", "t/h", "Steam & BFW plant", 284.3),
    ("IEA_UT_DMW_STEAM_B", "DMW: Steam and BFW Plant (row 2)", "-103,0", "t/h", "Steam & BFW plant", -103.0),
    ("IEA_UT_DMW_UTIL", "DMW: Utilities / BoP", "-17,0 (1)", "t/h", "Utilities / BoP", -17.0),
    ("IEA_UT_EFFLUENT_TOTAL", "Effluent (to WWT): Utilities / BoP = TOTAL", "-5,1", "t/h", "Waste-water discharge", -5.1),
    ("IEA_UT_RAWWATER_TOTAL", "Raw water: Utilities / BoP = TOTAL", "22,1", "t/h", "Raw-water intake at battery limit", 22.1),
    ("IEA_UT_CW_NH3", "Cooling water (closed circuit): Ammonia and Urea Plant", "4.871", "m3/h (dT 11 degC)", "Closed secondary cooling circuit", 4871),
    ("IEA_UT_CW_UTIL", "Cooling water (closed circuit): Utilities / BoP", "-4.871", "m3/h (dT 11 degC)", "Closed secondary cooling circuit", -4871),
    ("IEA_UT_SW_NH3", "Sea water: Ammonia and Urea Plant", "1.510", "m3/h (dT 7 degC)", "Sea-water cooling", 1510),
    ("IEA_UT_SW_STEAM", "Sea water: Steam and BFW Plant", "7.600", "m3/h (dT 7 degC)", "Sea-water cooling", 7600),
    ("IEA_UT_SW_UTIL", "Sea water: Utilities / BoP", "7.661", "m3/h (dT 7 degC)", "Sea-water cooling", 7661),
    ("IEA_UT_SW_TOTAL", "Sea water: TOTAL", "16.771", "m3/h", "Sea-water cooling (in/out)", 16771),
    ("IEA_UT_FUEL_NH3", "Fuel: Ammonia and Urea Plant", "118,1", "MMKcal/h", "Fuel consumers inside ammonia & urea plant", 118.1),
    ("IEA_UT_FUEL_TOTAL", "Fuel: TOTAL", "118,1", "MMKcal/h", "Whole complex", 118.1),
    ("IEA_UT_IA_NH3", "Instrument air: Ammonia and Urea Plant", "200", "Nm3/h", "Utility air", 200),
    ("IEA_UT_N2UTIL_NH3", "Nitrogen (utility): Ammonia and Urea Plant", "(250)", "Nm3/h", "UTILITY nitrogen (not process nitrogen)", None),
    ("IEA_UT_N2UTIL_UTIL", "Nitrogen (utility): Utilities / BoP", "(250) / (-500)", "Nm3/h", "UTILITY nitrogen", None),
]
for pid, name, v, u, b, num in util:
    note = "European format verified on rendered page: '.' = thousands separator, ',' = decimal mark. Negative = production/return per table convention."
    if pid == "IEA_UT_DMW_UTIL":
        note += " '(1)' is note marker: DMW = DMW plus condensate from the process unit."
    if pid.startswith("IEA_UT_LOSS"):
        note += " Note (2): losses = net water consumption (-) and production (+) in reactions and steam/BFW vent and steam losses."
    if pid == "IEA_UT_EFFLUENT_TOTAL":
        note += " Note (3): includes demi-plant eluate, process condensate and steam-system blowdown."
    if pid.startswith("IEA_UT_N2UTIL"):
        note = "Values printed in parentheses; meaning of parentheses not defined in the table. This is a utility (blanketing/purge) nitrogen column and is NOT the process nitrogen for ammonia synthesis. Total row prints 0."
    if pid.startswith("IEA_UT_FUEL"):
        note += " Fuel appears only in the Ammonia and Urea Plant row; Steam/BFW and Utilities rows are blank (blank is not recorded as zero)."
    if pid.startswith("IEA_UT_PWR"):
        note += " Header unit kW; values equal MWe figures on p.53."
    A(pid, name, v, u, S, 55, "Utility table sheet 1 of 1 (follows Sheet 23 of 80)", "Estimated Utility Consumptions, Ammonia-Urea Plant without CO2 Capture", C, Y, b, "sourced coefficient", note, num)
A("IEA_UT_OTHER_FUEL", "Other utility fuel (separately reported)", "", "", S, 55, "Utility table", "Estimated Utility Consumptions", C, Y, "Steam & BFW / Utilities rows", "unquantified", "No separate fuel entry for Steam & BFW plant or Utilities/BoP (cells blank). Source does not establish whether other utility fuel exists or is zero.")
A("IEA_UT_ELEC_EXPORT", "Electricity export", "", "", S, 53, "Sheet 22 of 80", "Power balance", C, Y, "Battery limit", "unquantified", "Only 'Imported Power from the Grid' is reported; no export row.")

# --- Water system description ---
A("IEA_W_DEMI", "Demi-water production route", "raw water processed by reverse-osmosis followed by electro-deionization", "text", S, 48, "Sheet 17 of 80", "Section 2.4.4", C, Y, "Demi-water plant", "sourced coefficient", "")
A("IEA_W_COOLING", "Cooling arrangement", "closed-circuit demi-water secondary system indirectly cooled by sea water via plate heat exchangers", "text", S, 48, "Sheet 17 of 80", "Section 2.4.4", C, Y, "Cooling system", "sourced coefficient", "Sea water is a once-through in/out battery-limit stream (p.28). Coastal-site design; not transferable to freshwater-cooled sites without justification.")
A("IEA_BL_LIST", "Plant battery limits (Case 4A)", "Natural gas (in); Waste water streams (out); Raw water (in); Sea water (in/out); Electric power (in); Ammonia (out) - base case only; Urea (out)", "text", S, 28, "Sheet 7 of 80", "Section 2.1.4 Plant Battery Limits", C, Y, "Plant battery limit", "sourced coefficient", "Compressed CO2 (out) applies to capture case only - excluded.")
A("IEA_STEAM_HP", "HP steam header conditions", "about 12.0 MPa and 510 degC", "MPa; degC", S, 47, "Sheet 16 of 80", "Section 2.4.3", C, Y, "Steam system", "sourced coefficient", "HP steam drives the syngas compressor turbine.")
A("IEA_STEAM_MP", "MP steam header conditions and users", "about 4.2 MPa and 375 degC; drives process air compressor, NH3 refrigeration compressor, urea CO2 compressor, SMR fans, BFW pumps", "text", S, 48, "Sheet 17 of 80", "Section 2.4.3", C, Y, "Steam system", "sourced coefficient", "Major drivers are steam turbines, not electric motors: relevant to electricity boundary.")

# --- Annex I reference document (Task 2) ---
A("IEA_SITE", "Plant location", "greenfield site at the North East coast of The Netherlands", "text", S, 127, "Annex I (Task 2) Sheet 8 of 24", "Section 3.1", C, "Dec 2016 (Task 2 document)", "Site", "sourced coefficient", "Design site; not one of the four study countries.")
A("IEA_AMB_T", "Average ambient air temperature (reference condition)", "9", "degC", S, 127, "Annex I Sheet 8 of 24", "Section 3.2", C, "Dec 2016", "Site", "sourced coefficient", "RH average 80%, pressure 101.3 kPa (reference conditions).", 9)
ng = [("Methane", "89.0"), ("Ethane", "7.0"), ("Propane", "1.0"), ("Butane", "0.1"), ("Pentane", "0.01"), ("Carbon Dioxide", "2.0"), ("Nitrogen", "0.89")]
for comp_name, v in ng:
    A(f"IEA_NGSPEC_{comp_name.replace(' ', '').upper()}", f"Natural gas analysis: {comp_name}", v, "vol%", S, 128, "Annex I Sheet 9 of 24", "Table 1. Natural Gas Specifications", C, "Dec 2016", "NG at battery limit", "sourced coefficient", "English format.", en(v))
A("IEA_NGSPEC_HHV", "Natural gas HHV", "51.473", "MJ/kg", S, 128, "Annex I Sheet 9 of 24", "Table 1", C, "Dec 2016", "NG", "sourced coefficient", "", 51.473)
A("IEA_NGSPEC_LHV", "Natural gas LHV", "46.502", "MJ/kg", S, 128, "Annex I Sheet 9 of 24", "Table 1", C, "Dec 2016", "NG", "sourced coefficient", "", 46.502)
A("IEA_NGSPEC_P", "Natural gas pressure at plant B.L.", "7.0", "MPa", S, 128, "Annex I Sheet 9 of 24", "Table 1", C, "Dec 2016", "NG at battery limit", "sourced coefficient", "", 7.0)
A("IEA_CAPFACTOR", "Capacity factor, ammonia/urea complex (Task 4)", "90", "%", S, 133, "Annex I Sheet 14 of 24", "Section 3.7", C, "Dec 2016", "Annual operation", "sourced coefficient", "Needed only if annual totals are ever used; Stage A functional unit is per tonne.", 90)

# ============================================================================
# H2FUTURE D9.2 -- independent cross-check (NOT to be merged into IEAGHG case)
# ============================================================================
S2, Y2 = "SRC_H2FUTURE_D9_2_2022", "2022"
H = lambda *a, **k: add("H2FUTURE_Crosscheck", *a, **k)
cases = ["Case 1A (standard, w/o CCS)", "Case 1B (hybrid, w/o CCS)", "Case 2A (standard, with CCS)", "Case 2B (hybrid, with CCU)", "Case 3 (100% electrolysis)"]
tab4 = [
    ("REF_NG_FEED", "Reforming: Natural gas feedstock", "t/d", ["649", "356", "649", "255", "-"]),
    ("REF_NG_FUEL", "Reforming: Natural gas fuel", "t/d", ["255", "140", "255", "100", "-"]),
    ("REF_STEAM", "Reforming: Water (steam)", "t/d", ["898", "493", "898", "353", "-"]),
    ("REF_O2_AIR", "Reforming: Oxygen (in process air)", "t/d", ["369", "202", "369", "145", "-"]),
    ("REF_N2_AIR", "Reforming: Nitrogen (in process air)", "t/d", ["1198", "657", "1198", "471", "-"]),
    ("REF_H2", "Reforming: Hydrogen production", "t/d", ["247", "136", "247", "97", "-"]),
    ("REF_ENERGY", "Reformer total energy input", "MWinput (LHV)", ["486", "267", "486", "191", "-"]),
    ("ELY_DEMIWATER", "Water-electrolysis: Water (demineralized)", "t/d", ["-", "996", "-", "1341", "2212"]),
    ("ELY_POWER", "Electrolyser", "MWinput", ["-", "244", "-", "328", "541"]),
    ("ELY_H2", "Water-electrolysis: Hydrogen production", "t/d", ["-", "111", "-", "150", "247"]),
    ("ELY_O2", "Water-electrolysis: Oxygen byproduct", "t/d", ["-", "884", "-", "1191", "1964"]),
    ("SYN_H2_IN", "Ammonia synthesis: Hydrogen total input", "t/d", ["247"] * 5),
    ("SYN_N2_IN", "Ammonia synthesis: Nitrogen total input", "t/d", ["1215"] * 5),
    ("SYN_N2_ASU", "Ammonia synthesis: Nitrogen (from ASU/external source)", "t/d", ["-", "541", "-", "727", "1198"]),
    ("SYN_NH3", "Ammonia total production", "t/d", ["1350"] * 5),
    ("UREA_NH3_IN", "Urea synthesis: Ammonia input", "t/d", ["702"] * 5),
    ("UREA_CO2_IN", "Urea synthesis: Carbon dioxide input", "t/d", ["908"] * 5),
    ("UREA_OUT", "Urea total production", "t/d", ["1238"] * 5),
    ("CB_CO2_FEED", "CO2 balance: CO2 from natural gas feedstock", "t/d", ["1709", "938", "1709", "672", "-"]),
    ("CB_CO2_FUEL", "CO2 balance: CO2 from fuel", "t/d", ["681", "373", "681", "267", "-"]),
    ("CB_CO2_EMIT", "CO2 balance: Total CO2 emissions", "t/d", ["1481", "403", "78", "31", "-"]),
    ("CB_CO2_STORE", "CO2 balance: CO2 to storage", "t/d", ["-", "-", "1403", "-", "-"]),
]
for pid, name, unit, vals in tab4:
    for case, v in zip(cases, vals):
        cid = case.split()[1]
        if v == "-":
            H(f"H2F_T4_{pid}_{cid}", name, "-", unit, S2, 24, "Page 24 of 46", "Table 4: Mass balances of process configurations", case, Y2, "H2FUTURE modelled plant (scaled from IEAGHG; 52% NH3 to urea)", "unquantified", "Printed as a dash; not assumed zero.")
        else:
            H(f"H2F_T4_{pid}_{cid}", name, v, unit, S2, 24, "Page 24 of 46", "Table 4: Mass balances of process configurations", case, Y2, "H2FUTURE modelled plant (scaled from IEAGHG; 52% NH3 to urea)", "sourced coefficient", "Cross-check only. Do not transfer into the IEAGHG Case 4A-1 reference.", en(v))
txt = [
    ("H2F_BASIS_NH3", "Ammonia plant capacity basis", "1350", "t/d", 20, "Section 3.1", "All", "Taken from IEAGHG study [9].", 1350),
    ("H2F_BASIS_UREA_SHARE", "Share of ammonia converted to urea", "52", "%", 20, "Section 3.1", "All", "DIFFERS from IEAGHG Case 4A-1 (~95%). 'average share of ammonia that is consumed globally for the production of urea'.", 52),
    ("H2F_SHARE_1B", "Hydrogen supply split, Case 1B", "reforming (55%) and water-electrolysis (45%)", "% of H2", 20, "Section 3.1 list", "Case 1B", "", None),
    ("H2F_SHARE_2B", "Hydrogen supply split, Case 2B", "reforming (40%) and water-electrolysis (60%)", "% of H2", 20, "Section 3.1 list", "Case 2B", "", None),
    ("H2F_METHOD_SCALING", "Mass-balance method", "mass balance constructed using the data from the IEAGHG study as a starting point and scaling different material flows according to the requirements of each case", "text", 20, "Section 3.1", "All", "Proportional scaling is the authors' method, not an IEAGHG result.", None),
    ("H2F_1A_ENERGY", "HyCO plant total energy input, Case 1A", "486", "MW (LHV)", 20, "Section 3.1.1", "Case 1A", "", 486),
    ("H2F_1A_UREA", "Urea production unit capacity (all cases)", "almost 1240", "t/d", 20, "Section 3.1.1", "All", "Table 4 gives 1238 t/d.", 1240),
    ("H2F_1A_RESID_CO2", "Residual CO2 sent to stack, Case 1A", "about 800", "t/d", 21, "Section 3.1.1", "Case 1A", "Not all CO2 removed from syngas is used for urea in H2FUTURE's 52% case.", 800),
    ("H2F_1A_TOTAL_CO2", "Total CO2 emissions incl. fuel CO2, Case 1A", "more than 1480", "t/d", 21, "Section 3.1.1", "Case 1A", "Table 4: 1481.", 1480),
    ("H2F_IEAGHG_NOVENT", "Statement about IEAGHG", "In the IEAGHG-study all CO2 is used for urea production, and no residual CO2 is sent to the stack.", "text", 21, "Section 3.1.1", "IEAGHG reference", "Secondary statement about the IEAGHG source; supports, but does not replace, IEAGHG p.52.", None),
    ("H2F_2A_EMIT", "CO2 emission / storage, Case 2A", "reduced to 78 t/d and 1403 t/d of CO2 is sent to a storage", "t/d", 21, "Section 3.1.2", "Case 2A", "", None),
    ("H2F_1B_REFORMER", "Reformer capacity reduction, Case 1B", "from 486 MW to 267 MW (LHV)", "MW (LHV)", 21, "Section 3.1.3", "Case 1B", "", None),
    ("H2F_1B_EMIT", "Remaining CO2 emission, Case 1B", "403", "t/d", 21, "Section 3.1.3", "Case 1B", "", 403),
    ("H2F_ELY_EFF", "Electrolysis plant efficiency", "75% (HHV)", "% (HHV basis)", 22, "Section 3.1.3", "Case 1B (and electrolysis cases)", "HHV basis. Boundary of 'electrolysis plant' (rectifier, BoP, purification, drying) not defined in text.", 75),
    ("H2F_ELY_CAP_1B", "Electrolyser capacity (electricity input), Case 1B", "244", "MW", 22, "Section 3.1.3", "Case 1B", "'to produce a little over 110 t/d of hydrogen'. No kWh/kg coefficient is stated in the report.", 244),
    ("H2F_N2_SHORTAGE", "Nitrogen shortage in hybrid case", "capacities of the primary and secondary reformer have been reduced proportionally. The mass balance analysis then showed a shortage of nitrogen ... because of the reduced amount of process air ... The nitrogen is supplemented from an air separation unit.", "text", 22, "Section 3.1.3", "Case 1B", "Authors: 'no data were found for an ASU of a quality comparable to the IEAGHG study'; effect approximated by nitrogen market price + sensitivity. ASU electricity not quantified.", None),
    ("H2F_O2_OPTION", "Oxygen by-product option", "oxygen could be used to replace the oxygen supplied via the process air ... this study did not further investigate", "text", 22, "Section 3.1.3", "Case 1B", "Not investigated by authors; outside Stage A design.", None),
    ("H2F_2B_N2", "Nitrogen supplement increase, Case 2B", "increases from 541 to 727", "t/d", 22, "Section 3.1.4 (numbered 1.1.1 in PDF)", "Case 2B", "", None),
    ("H2F_2B_EMIT", "Direct emissions left, Case 2B", "about 30", "t/d", 22, "Section 3.1.4", "Case 2B", "Table 4 gives 31.", 30),
    ("H2F_3_N2_CO2", "Case 3 nitrogen and urea CO2", "All nitrogen must now be obtained from an ASU ... no CO2 available for the production of urea", "text", 23, "Section 3.1.5 (numbered 1.1.1 in PDF)", "Case 3", "", None),
    ("H2F_COMP_P", "Hydrogen compression pressure (electrolysis cases)", "30", "bar", 26, "Section 3.2.1", "Cases 1B, 2B, 3", "Gauge/absolute not stated. Electrolyser system CAPEX is 'excluding compression'; compressor is a separate item.", 30),
    ("H2F_ELEC_SCOPE", "Electricity demand scope (country analysis)", "electricity required for electrolysis and H2 compression", "text", 34, "Section 4.2", "Cases 1B, 2B, 3", "Compression electricity is included in aggregate demand but no separate coefficient is given in the inspected text.", None),
    ("H2F_STACK", "Stack replacement", "approximately 50% of the total electrolysis system cost (excl. compression); replaced once in 25-year lifetime", "text", 27, "Section 3.2.2", "Electrolysis cases", "Cost item; not used in Stage A physical accounting.", None),
]
for pid, name, v, u, pg, tfe, case, n, num in txt:
    H(pid, name, v, u, S2, pg, f"Page {pg} of 46", tfe, case, Y2, "H2FUTURE modelled plant", "sourced coefficient", n, num)
H("H2F_ELY_KWHKG", "Electrolyser specific electricity per kg H2", "", "", S2, 22, "Page 22 of 46", "Section 3.1.3 / Table 4", "Electrolysis cases", Y2, "Electrolysis plant", "unquantified", "Not stated. Could be derived from MW input and t/d output, but derivation is a calculation outside this transcription task.")
H("H2F_H2_PURITY", "Delivered electrolytic hydrogen purity / drying / deoxo", "", "", S2, 26, "Page 26 of 46", "Section 3.2.1", "Electrolysis cases", Y2, "Electrolysis plant", "unquantified", "No purity, drying or deoxygenation statement found.")
H("H2F_ELY_WATER_WD", "Electrolysis water withdrawal / consumption / cooling", "", "", S2, 24, "Page 24 of 46", "Table 4", "Electrolysis cases", Y2, "Electrolysis plant", "unquantified", "Only demineralised feed water reported; no raw-water withdrawal, treatment reject or cooling water.")

# ============================================================================
# IRENA & Bluerisk 2023 -- water definitions and route coefficients
# ============================================================================
S3, Y3 = "SRC_IRENA_BLUERISK_2023", "2023"
W = lambda *a, **k: add("IRENA_Water", *a, **k)
W("IRW_DEF_WITHDRAWAL", "Definition: water withdrawal", "Measured by the quantity of water withdrawn from a source (e.g. river, lake, groundwater) for use.", "text", S3, 7, "printed p.5", "Glossary", "Definitions", Y3, "Definition", "sourced coefficient", "")
W("IRW_DEF_CONSUMPTION", "Definition: water consumption", "The portion of withdrawn water that is not returned to the source.", "text", S3, 7, "printed p.5", "Glossary", "Definitions", Y3, "Definition", "sourced coefficient", "Consumption is a subset of withdrawal: never add the two.")
W("IRW_DEF_INTENSITY", "Definition: withdrawal/consumption intensity", "The quantity of water withdrawn for or consumed in the generation of a unit of a product", "text", S3, 7, "printed p.5", "Glossary", "Definitions", Y3, "Definition", "sourced coefficient", "")
W("IRW_DEF_MAKEUP_BLOWDOWN", "Definitions: makeup water; blowdown water; cycle of concentration; permeate rate", "Makeup: water added to replace losses; Blowdown: water drained to prevent mineral build-up; CoC: dissolved-solid concentration ratio out/in; Permeate rate: membrane permeate / raw water", "text", S3, 7, "printed p.5", "Glossary", "Definitions", Y3, "Definition", "sourced coefficient", "Paraphrased; see glossary for exact wording.")
t21 = [
    ("COAL", "Brown: Coal gasification", ["49.78", "51.41", "48.14", "31.00", "32.02", "29.98"]),
    ("SMR", "Grey: Natural gas-SMR", ["20.01", "25.16", "16.40", "17.54", "19.80", "15.80"]),
    ("COAL_CCUS", "Blue: Coal gasification-CCUS", ["80.23", "87.21", "73.85", "49.44", "52.47", "46.53"]),
    ("SMR_CCUS", "Blue: Natural gas-SMR-CCUS", ["36.69", "47.79", "29.81", "32.18", "38.96", "24.15"]),
    ("ATR_CCUS", "Blue: Natural gas-ATR-CCUS", ["30.76", "30.76", "30.76", "24.22", "24.22", "24.22"]),
    ("AEL", "Green: Electrolysis-Alkaline", ["32.24", "34.61", "29.88", "22.28", "23.59", "20.96"]),
    ("PEM", "Green: Electrolysis-PEM", ["25.70", "26.46", "24.94", "17.52", "18.04", "17.00"]),
]
stats = [("WD", "AVG", "withdrawal intensity, average"), ("WD", "MAX", "withdrawal intensity, max"), ("WD", "MIN", "withdrawal intensity, min"),
         ("CON", "AVG", "consumption intensity, average"), ("CON", "MAX", "consumption intensity, max"), ("CON", "MIN", "consumption intensity, min")]
for code, tech, vals in t21:
    for (b, st, label), v in zip(stats, vals):
        note = "Hydrogen-production facility only; excludes upstream water for electricity or fossil-fuel supply (footnote 3, p.31). Tap-water source assumed; evaporative cooling; review of industry models + literature (few metered data)."
        if code == "SMR":
            note = "INDICATIVE FOR LEVEL 2 ONLY: generic merchant SMR hydrogen plant (literature + industry models), not an integrated two-step-reforming ammonia plant. " + note
        if code in ("AEL", "PEM"):
            note = "Total electrolysis-facility water (pre-treatment reject + deionised feed + evaporative cooling make-up). Efficiency-dependent (Fig. 2.4). " + note
        W(f"IRW_T21_{code}_{b}_{st}", f"{tech}: {label}", v, "L/kg (H2)", S3, 31, "printed p.29", "Table 2.1 Summary of water withdrawal and consumption intensities", tech, Y3,
          "Hydrogen production facility incl. water pre-treatment and cooling; excludes upstream", "sourced coefficient", note, en(v))
a1 = [
    ("Coal-water slurry gasification", "48.14", "29.98", "Design data from industry interviews"),
    ("Coal-water slurry gasification", "51.41", "32.02", "(Cui et al., 2021)"),
    ("Electrolysis-AE", "34.61", "23.59", "Modelled data from industry interviews"),
    ("Electrolysis-AE", "29.88", "20.96", "Modelled data from industry interviews"),
    ("Electrolysis-PEM", "24.94", "17.00", "(Newborough and Cooley, 2022)"),
    ("Electrolysis-PEM", "26.46", "18.04", "(Mehmeti et al., 2018)"),
    ("Electrolysis-SOEC", "15.86", "10.81", "(Elgowainy et al., 2016)"),
    ("Natural gas-ATR-CCS", "30.76", "24.22", "(Lewis et al., 2022)"),
    ("Natural gas-SMR", "16.40", "15.80", "(Lewis et al., 2022)"),
    ("Natural gas-SMR", "25.16", "17.27", "Modelled data from industry interviews"),
    ("Natural gas-SMR", "20.55", "19.80", "(Spath and Mann, 2001)"),
    ("Natural gas-SMR", "17.92", "17.27", "(Simon, Daily, and White, 2010)"),
    ("Natural gas-SMR-CCS", "29.81", "24.15", "(Lewis et al., 2022)"),
    ("Coal-water slurry gasification CCS", "73.85", "46.53", "Estimated based on (Cui et al., 2021; and Rosa et al., 2021) and industry interview data"),
    ("Coal-water slurry gasification CCS", "79.64", "49.33", "Estimated based on (Cui et al., 2021; and Rosa et al., 2021) and industry interview data"),
    ("Coal-water slurry gasification CCS", "87.21", "52.47", "Estimated based on (Cui et al., 2021; and Rosa et al., 2021) and industry interview data"),
    ("Natural gas-SMR-CCS", "31.15", "31.09", "Estimated based on (Lewis et al., 2022; Rosa et al., 2021; Simon, Daily, and White, 2010; Spath and Mann, 2001) and industry interview data"),
    ("Natural gas-SMR-CCS", "38.01", "34.50", "Estimated based on (Lewis et al., 2022; Rosa et al., 2021; Simon, Daily, and White, 2010; Spath and Mann, 2001) and industry interview data"),
    ("Natural gas-SMR-CCS", "47.79", "38.96", "Estimated based on (Lewis et al., 2022; Rosa et al., 2022; Simon, Daily, and White, 2010; Spath and Mann, 2001) and industry interview data"),
]
for i, (tech, wd, con, src) in enumerate(a1, 1):
    for b, v in (("withdrawal", wd), ("consumption", con)):
        W(f"IRW_A1_{i:02d}_{b[:3].upper()}", f"{tech}: water {b} intensity (data point {i})", v, "L/kg (H2)", S3, 65, "printed p.63", "Appendix Table A1 Water withdrawal and consumption intensity data sources", tech, Y3,
          "As Table 2.1", "sourced coefficient", f"Underlying source: {src}. Secondary compilation; original study boundaries not re-checked.", en(v))
fig_assump = [
    ("IRW_F21_SOURCE", "Fig. 2.1 assumption: source water", "tap water is used as the source water", "text", None),
    ("IRW_F21_PERMEATE", "Fig. 2.1 assumption: pre-treatment permeate rate (green, grey, blue)", "66", "%", 66),
    ("IRW_F21_EFF_GREEN", "Fig. 2.1 assumption: energy efficiency, green (alkaline electrolysis)", "70", "%", 70),
    ("IRW_F21_EFF_GREY", "Fig. 2.1 assumption: energy efficiency, grey (SMR)", "76", "%", 76),
    ("IRW_F21_EFF_BLUE", "Fig. 2.1 assumption: energy efficiency, blue (SMR+CCUS)", "71", "%", 71),
    ("IRW_F21_COOLING", "Fig. 2.1 assumption: cooling", "evaporative cooling is assumed for all the cooling processes; cycle of concentration of 6 for green and brown production", "text", None),
    ("IRW_F21_DATA_ORIGIN", "Fig. 2.1 data origin", "green and brown: engineering design models obtained from the industry; grey and blue: Cases 1 and 2 in NETL 2022", "text", None),
]
for pid, name, v, u, num in fig_assump:
    W(pid, name, v, u, S3, 27, "printed p.25 (figure spans printed pp.24-25)", "Figure 2.1 note", "Fig. 2.1 example plants", Y3, "Example facility (schematic)", "sourced coefficient",
      "LHV/HHV basis of efficiency NOT stated." if "EFF" in pid else "", num)
W("IRW_F21_VALUES", "Fig. 2.1 per-stream schematic volumes (L per kg H2)", "not transcribed", "L/kg", S3, 26, "printed pp.24-25", "Figure 2.1", "Fig. 2.1 example plants", Y3, "Example facility", "unquantified", "Schematic stream values are laid out graphically; transcribe only after visual check of PDF pp.26-27 if needed.")
W("IRW_F22_COOL_SHARE", "Cooling make-up share of total withdrawal: green / brown / grey", "about 56% and 52% (green, brown); about 14% (grey)", "% of withdrawal", S3, 28, "printed p.26", "Figure 2.2 and text", "Example facilities", Y3, "Facility", "sourced coefficient", "Figure axis is labelled 'm3/kg' although values (49.8, 20.0, 32.2) match L/kg elsewhere: unit label inconsistency in source. Grey data attributed to Lewis et al. (2022) Case 1 here vs 'NETL 2022' in Fig 2.1 note.")
W("IRW_COC_RANGE", "Cycle of concentration, evaporative cooling (typical)", "4 to 6", "-", S3, 29, "printed p.27", "Chapter 2 text", "General", Y3, "Cooling system", "sourced coefficient", "'about 70%-85% of the water withdrawn for cooling evaporates (or is consumed)'.")
W("IRW_PERMEATE_RANGE", "Pre-treatment permeate rate by source water", "66% tap water; 58% river or groundwater; 35% seawater", "%", S3, 29, "printed p.27", "Chapter 2 text", "General", Y3, "Water pre-treatment", "sourced coefficient", "Based on data shared by the industry.")
W("IRW_EFF_SLOPE", "Electrolysis efficiency effect on water intensity", "For every 1 percentage point increase in electrolysis efficiency, the intensity of water withdrawal and consumption ... falls by about 2%", "% per percentage point", S3, 32, "printed p.30", "Figure 2.4 and text", "Typical green H2 project", Y3, "Electrolysis facility", "sourced coefficient", "Fig. 2.4 plots 60-80% efficiency; basis (LHV/HHV) not stated.")
W("IRW_UPSTREAM", "Upstream water inclusion", "this value does not consider upstream water withdrawal and consumption - i.e. the water needs for producing the electricity or fossil fuels consumed in hydrogen facilities", "text", S3, 31, "printed p.29", "Footnote 3", "All technologies", Y3, "Boundary statement", "sourced coefficient", "Applies to the report's intensities; excludes electricity-generation water.")
W("IRW_DATAQUALITY", "Data-quality statement", "A significant portion of the source data points are based on industry models rather than from metering", "text", S3, 25, "printed p.23", "Chapter 2 text", "All technologies", Y3, "Data quality", "sourced coefficient", "Also p.24: data on water requirements 'insufficient as well as of inadequate quality'. SOEC/AEM excluded from review (p.25).")

# ----------------------------------------------------------------------------
# Boundary register and unresolved issues
# ----------------------------------------------------------------------------
BOUND_COLS = ["boundary_id", "boundary_name", "definition_in_source", "source_id", "pdf_page", "printed_page_or_sheet", "applies_to_records", "compatible_with", "not_compatible_with", "notes"]
BOUNDS = [
    ("B01", "IEAGHG Case 4A-1 plant battery limit", "NG in; raw water in; sea water in/out; electric power in; waste water out; ammonia out (base case); urea out", "SRC_IEAGHG_2017_03", 28, "Sheet 7 of 80", "IEA_BL_LIST, IEA_PD_*, IEA_UT_*", "Whole-plant (Level 3) S0 descriptions", "Generic per-kg-H2 route coefficients (IRENA)", "Coastal Netherlands design site with sea-water cooling."),
    ("B02", "IEAGHG candidate H2 delivery point: stream 7 'Syngas to compression'", "Mixed make-up syngas after methanation: H2 0,7300; N2 0,2580; Ar 0,0030; CH4 0,0070; H2O 0,0020 mol/mol; 3,00 MPa; 43 degC", "SRC_IEAGHG_2017_03", 33, "H&MB", "IEA_HMB_S07_*", "Any H2 meeting CO+CO2 < 5 ppm (p.28) delivered at equal pressure into syngas compressor suction", "Pure electrolytic H2 at undefined pressure/purity; HRU internal recycle", "Absolute/gauge basis of 3,00 MPa not stated; HRU H2 recycle not quantified."),
    ("B03", "NG feedstock account", "Stream 2 'Natural Gas feedstock to Ammonia Plant'", "SRC_IEAGHG_2017_03", 33, "H&MB; Sheet 22", "IEA_HMB_S02_*, IEA_PD_NG_FEED", "Feedstock account only", "Fuel accounts", ""),
    ("B04", "NG fuel to burners (reformer fuel) account", "Stream 3 'Natural Gas fuel to burners'; BFD: 'Natural gas to SMR burners'", "SRC_IEAGHG_2017_03", 33, "H&MB; BFD p.31; Sheet 22", "IEA_HMB_S03_*, IEA_PD_NG_FUEL, IEA_UT_FUEL_*", "Reformer-fuel account", "Other-utility-fuel account (not separately reported)", "Single fuel figure; tail gas (stream 4) is additional internal fuel."),
    ("B05", "Other utility fuel account", "Not separately reported", "SRC_IEAGHG_2017_03", 55, "Utility table", "IEA_UT_OTHER_FUEL", "-", "-", "Unquantified; blank cells not read as zero."),
    ("B06", "IEAGHG plant electricity (grid import)", "Imported power by area: NH3/urea, steam+BFW, utilities+BoP", "SRC_IEAGHG_2017_03", 53, "Sheet 22 of 80", "IEA_PD_PWR_*, IEA_UT_PWR_*", "Rest-of-plant electricity (S0)", "Electrolyser electricity; shaft power from steam turbines", "Large compressors are steam-turbine driven."),
    ("B07", "H2FUTURE electrolysis plant electricity", "'Electrolyser MWinput'; efficiency 75% (HHV); compression to 30 bar costed separately", "SRC_H2FUTURE_D9_2_2022", 24, "Page 24/22/26 of 46", "H2F_T4_ELY_POWER_*, H2F_ELY_EFF, H2F_COMP_P", "Cross-check only", "IEAGHG reference case values", "BoP/purification inclusion not defined; no kWh/kg stated."),
    ("B08", "Water Level 1 (electrolyser account)", "IRENA: electrolysis facility total (pre-treatment + feed + evaporative cooling), tap water, per kg H2; H2FUTURE: demineralised feed water only", "SRC_IRENA_BLUERISK_2023; SRC_H2FUTURE_D9_2_2022", 31, "printed p.29; H2F p.24", "IRW_T21_AEL_*, IRW_T21_PEM_*, H2F_T4_ELY_DEMIWATER_*", "Same basis (withdrawal OR consumption) only", "Adding demineralised feed to IRENA totals (double counts reaction water)", ""),
    ("B09", "Water Level 2 (hydrogen-route comparison)", "IRENA generic Natural gas-SMR vs electrolysis, same report, same definitions", "SRC_IRENA_BLUERISK_2023", 31, "printed p.29", "IRW_T21_SMR_*", "IRENA electrolysis coefficients (same table)", "IEAGHG plant water; H2FUTURE feed water", "INDICATIVE only: generic SMR, not integrated ammonia reforming."),
    ("B10", "Water Level 3 (whole plant)", "IEAGHG utility table: raw water, losses, effluent, sea water, closed cooling circuit", "SRC_IEAGHG_2017_03", 55, "Utility table", "IEA_UT_RAWWATER_TOTAL, IEA_UT_LOSS_*, IEA_UT_EFFLUENT_TOTAL, IEA_UT_SW_*", "S0 only", "IRENA per-kg coefficients; any scenario balance (none available)", "Sea water is once-through cooling, not freshwater."),
    ("B11", "CO2 account (IEAGHG)", "Equivalent CO2 flows in kmol/h: carbon in feed, fuel, urea product, flue gas; vents '-'", "SRC_IEAGHG_2017_03", 52, "Sheet 21 of 80", "IEA_CB_*, IEA_HMB_S05_*, IEA_HMB_S08_*", "Stage A CO2 module", "'Carbon Not Emitted' classification; H2FUTURE t/d CO2 balance", "kmol/h basis, not tonnes."),
    ("B12", "IRENA upstream exclusion", "Excludes water for producing electricity or fossil fuels", "SRC_IRENA_BLUERISK_2023", 31, "printed p.29 footnote 3", "IRW_*", "Direct facility water", "Life-cycle water footprints", ""),
]
ISSUE_COLS = ["issue_id", "topic", "source_evidence", "pdf_page_and_reference", "can_be_extracted", "cannot_yet_be_quantified", "consequence_for_workbench", "recommendation_for_supervisor", "status"]
ISSUES = [
    ("U01", "Nitrogen supply", "IEAGHG: N2 supplied by process air into the air-blown secondary reformer; air amount fixed by H2/N2 ~3 (p.37, p.40). Process air stream 6 = 2286 kmol/h, 66079 kg/h (p.33). H2FUTURE: proportional reformer reduction causes N2 shortage, filled by ASU (541 t/d in 1B, 727 t/d in 2B); no ASU data comparable to IEAGHG found (p.22, Table 4 p.24).",
     "IEAGHG p.37 Sheet 10; p.40 Sheet 11; p.33 H&MB stream 6. H2FUTURE p.22; p.24 Table 4", "Reference process-air flow and composition; textual rule that air is fixed by H2/N2 ratio; H2FUTURE cross-check values", "N2 deficit at S10-S50; ASU (or other N2 source) electricity; whether secondary-reformer operation is feasible at reduced throughput", "Reference N2 supply is NOT automatically consistent with substitution scenarios; must be recorded as a consistency check/limitation", "Decide how Stage A reports the nitrogen consequence (qualitative limitation vs. sourced ASU coefficient). Interpretation, not source: proportional scaling implies a deficit equal to the N2 associated with the electrolytic share.", "Requires supervisor decision"),
    ("U02", "Natural-gas accounts", "IEAGHG reports NG feedstock (stream 2, 27021 kg/h) and one NG fuel figure 'Natural Gas fuel to burners' (stream 3, 10630 kg/h); utility table shows fuel only in the Ammonia and Urea Plant row (118,1 MMkcal/h). Tail gas (stream 4) and CO2-removal flash gas also fire the SMR.",
     "IEAGHG p.33 H&MB; p.53 Sheet 22; p.55 utility table; p.40-41 Sheets 11-12", "Feedstock account; fuel-to-burners account", "A separate 'other utility fuel' account (not reported); flash-gas fuel; fuel response to reduced reforming", "Third NG account must be carried as 'not separately reported/unquantified'; fuel must not be scaled", "Confirm whether the third account is reported as 'not separately reported in reference' (no zero) or whether another source is sought.", "Requires supervisor decision"),
    ("U03", "CO2 destinations and vent", "IEAGHG CO2 balance: in 2263 kmol/h (feed 1626, fuel 637); out: carbon in urea 1568, flue gas 695; Vents '-'. Source labels urea carbon 'Carbon Not Emitted'. Stream 8 CO2 to urea is wet (CO2 0,9400; H2O 0,0510; H2 0,0080). Flash gas and HRU tail gas carbon routes not separately quantified.",
     "IEAGHG p.52 Sheet 21; p.33 H&MB streams 5, 8, 11; p.41 Sheet 12", "Input carbon by account; carbon in urea; flue-gas CO2; CO2-to-urea stream flow and composition; purity/pressure statement (p.44)", "Vent flow (printed dash); internal carbon routes; usable-CO2 fraction in scenarios", "Shortfall cannot be computed until the S0 carbon balance is reconciled (Step 2). Vent must not be set to zero without approval", "Approve how a printed dash is treated; confirm 'Carbon Not Emitted' wording is replaced by 'carbon carried in urea product'.", "Requires supervisor decision"),
    ("U04", "Ammonia-to-urea relationship", "IEAGHG Case 4A-1: 1,350 t/d NH3, 2,260 t/d urea, ~95% conversion, 68.4 t/d NH3 product, no CO2 vent. H2FUTURE: 52% NH3 to urea, 1238 t/d urea, ~800 t/d residual CO2 vented in Case 1A.",
     "IEAGHG p.27 Sheet 6; p.53 Sheet 22; p.33 streams 9-11. H2FUTURE p.20-21; p.24 Table 4", "Both sources' product splits, separately", "Nothing missing for transcription; reconciliation of IEAGHG product and carbon balances is Step 2", "H2FUTURE balances cannot be mixed with IEAGHG; CO2 constraint behaves differently (IEAGHG has no surplus CO2)", "Confirm IEAGHG 95% relationship is retained for Tier A; H2FUTURE used only as cross-check/discussion.", "Requires supervisor confirmation"),
    ("U05", "Hydrogen delivery point", "IEAGHG make-up syngas (stream 7) is an H2/N2 mixture at 3,00 MPa, 43 degC with CO+CO2 < 5 ppm spec. H2FUTURE electrolytic H2 compressed to 30 bar; no purity, drying or deoxo statement. HRU recycle H2 internal and unquantified.",
     "IEAGHG p.33 stream 7; p.28 Sheet 7; p.42 Sheet 13; p.31 BFD. H2FUTURE p.26", "Stream 7 flow, composition, T, P; purity spec", "Pressure basis (abs/gauge) equivalence; electrolytic H2 purity/conditioning; HRU recycle", "Delivered-H2 boundary should be defined at stream 7; electrolytic H2 conditioning load is unquantified", "Approve stream 7 as the common delivery point and how conditioning (drying/deoxo) is treated.", "Requires supervisor decision"),
    ("U06", "Water definitions and levels", "IRENA defines withdrawal and consumption (p.7) and gives facility-level intensities incl. pre-treatment and evaporative cooling, tap-water source, excluding upstream (p.31). H2FUTURE gives demineralised feed only (p.24). IEAGHG gives plant raw water, losses, effluent, sea water (p.55).",
     "IRENA p.7, p.27, p.31, p.32, p.65. H2FUTURE p.24. IEAGHG p.55, p.48", "Level 1 candidates (IRENA AEL/PEM); Level 2 indicative (IRENA SMR); Level 3 S0 descriptors (IEAGHG)", "Level 3 scenario balances; site-specific cooling; country conditions", "Keep three levels separate; Level 2 labelled indicative; Level 3 S0 only", "Approve which IRENA statistic (average vs range) is used and the Level 2 'indicative' label.", "Requires supervisor decision"),
    ("U07", "Electricity boundaries", "IEAGHG: 8.890 MWe grid import by area; major compressors are steam-turbine driven (p.47-48). H2FUTURE: electrolyser MW input, 75% HHV, compression separate (CAPEX excl. compression), no kWh/kg; ASU electricity not quantified.",
     "IEAGHG p.53 Sheet 22; p.55; p.47-48 Sheets 16-17. H2FUTURE p.22, p.24, p.26, p.34", "Reference electricity by area; H2FUTURE MW inputs and efficiency", "System-level electrolyser kWh/kg incl. BoP/purification/compression; steam-to-electric driver changes", "A system-level electrolyser coefficient is still missing (Jin et al. 2024 pending)", "Decide the electrolyser coefficient source once Jin et al. is verified; confirm compression is a separate line.", "Pending source"),
    ("U08", "H&MB stream 12 blank", "Column 12 printed empty", "IEAGHG p.33", "-", "Meaning", "None", "None needed", "Recorded"),
    ("U09", "Utility nitrogen values in parentheses", "(250), (-500) Nm3/h; meaning of parentheses undefined", "IEAGHG p.55", "As printed", "Meaning", "Not process nitrogen; do not use for U01", "None needed", "Recorded"),
    ("U10", "IRENA Fig 2.2 unit label and data-origin inconsistency", "Axis 'm3/kg' with L/kg magnitudes; grey data credited to NETL 2022 (Fig 2.1) and Lewis et al. 2022 Case 1 (Fig 2.2)", "IRENA p.26-28", "Table 2.1 values (unaffected)", "-", "Use Table 2.1, not Fig 2.2 values", "None needed", "Recorded"),
    ("U11", "Jin et al. (2024) full text", "Bibliographic record verified via Crossref (DOI 10.1093/ce/zkad072, CC BY 4.0); PDF download blocked (HTTP 403 bot check)", "-", "Citation only", "All values; relevance; system-level vs product-spec", "Electrolyser electricity coefficient remains open", "Researcher to supply PDF", "Pending"),
]

# ----------------------------------------------------------------------------
# Write manifest CSV
# ----------------------------------------------------------------------------
man_cols = ["source_id", "file_name", "local_path", "opens", "pages", "file_size_bytes", "sha256", "bibliographic_reference", "download_url", "download_note", "access_licence", "verification_status"]
with open(os.path.join(OUT_DIR, "source_manifest.csv"), "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(man_cols)
    for s in SOURCES:
        w.writerow([s["source_id"], s["file"], s["local_path"], s["opens"], s["pages"], s["bytes"], s["sha256"], s["reference"], s["url"], s["url_note"], s["access"], s["status"]])

# ----------------------------------------------------------------------------
# Write workbook
# ----------------------------------------------------------------------------
wb = Workbook()
HDR = Font(bold=True, color="FFFFFF"); FILL = PatternFill("solid", fgColor="1F4E78"); WRAP = Alignment(wrap_text=True, vertical="top")

def sheet(ws, cols, rows, widths=None):
    ws.append(cols)
    for c in ws[1]:
        c.font, c.fill, c.alignment = HDR, FILL, WRAP
    for r in rows:
        ws.append([r.get(c) for c in cols] if isinstance(r, dict) else list(r))
    for i, col in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = (widths or {}).get(col, 22)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = WRAP
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

ws = wb.active; ws.title = "Source_Register"
sheet(ws, man_cols, [[s["source_id"], s["file"], s["local_path"], s["opens"], s["pages"], s["bytes"], s["sha256"], s["reference"], s["url"], s["url_note"], s["access"], s["status"]] for s in SOURCES],
      {"bibliographic_reference": 70, "verification_status": 60, "download_url": 50, "sha256": 40, "download_note": 50})
wid = {"parameter_name": 40, "reported_value": 30, "extraction_notes": 70, "table_figure_equation": 35, "system_boundary": 30, "reference_case": 28}
for name in ("IEAGHG_Reference", "H2FUTURE_Crosscheck", "IRENA_Water"):
    sheet(wb.create_sheet(name), COLS, records[name], wid)
sheet(wb.create_sheet("Boundary_Register"), BOUND_COLS, BOUNDS, {"definition_in_source": 60, "notes": 50, "applies_to_records": 40, "compatible_with": 35, "not_compatible_with": 35})
sheet(wb.create_sheet("Unresolved_Issues"), ISSUE_COLS, ISSUES, {c: 45 for c in ISSUE_COLS})
rd = wb.create_sheet("README", 0)
for line in [
    "Stage A Source-Extraction Register - TRANSCRIPTION ONLY",
    "No unit conversion, per-tonne normalisation, reconciliation or scenario calculation has been performed.",
    "reported_value = exactly as printed. value_numeric = how the printed number format was read (European '4.490' = 4490; '7,00' = 7.00); it is NOT a converted value.",
    "pdf_page = page of the PDF file; printed_page_or_sheet = page/sheet number printed in the document (IEAGHG contractor report uses 'Sheet N of 80'; IRENA printed page = PDF page - 2).",
    "IEAGHG records are Case 4A-1 only. H2FUTURE values are a cross-check and must not be transferred into the IEAGHG reference case.",
    "evidence_status 'unquantified' = not reported (blank or dash); never interpreted as zero.",
    "The source label 'Carbon Not Emitted' is recorded but not adopted as a classification.",
    "Built by documentation/tools/build_stageA_source_register.py from data/raw/source_reports/ (see source_manifest.csv for checksums).",
]:
    rd.append([line])
rd.column_dimensions["A"].width = 160
wb.save(os.path.join(OUT_DIR, "StageA_Source_Extraction_Register.xlsx"))

for k, v in records.items():
    print(k, len(v), "records;", sum(r["evidence_status"] == "unquantified" for r in v), "unquantified")
print("Boundary rows", len(BOUNDS), "Issues", len(ISSUES))
