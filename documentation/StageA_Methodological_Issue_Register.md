# Stage A Methodological Issue Register

**Status:** source evidence only. Nothing here approves an assumption. Items marked *Interpretation* are my reading, not the source's words.
**Page references:**
- `p.` means the PDF page.
- IEAGHG contractor sheets are shown as "Sheet N of 80"; its Annex I uses "Sheet N of 24".
- IRENA printed page = PDF page − 2.

**Record IDs** refer to `data/source_register/StageA_Source_Extraction_Register.xlsx`.

---

## 1. Nitrogen supply

**What the sources say**
- **IEAGHG:**
  - "the nitrogen is supplied by introducing the process air into the secondary reformer" (p.37, Sheet 10).
  - "the amount of process air introduced into the ATR is therefore fixed" because H₂/N₂ is held close to 3 (p.40, Sheet 11).
  - Process air (stream 6) is 2286 kmol/h, 66079 kg/h, with N₂ 0,7800, O₂ 0,2100 and Ar 0,0100 mol/mol (p.33, heat and material balance).
- **H2FUTURE D9.2:**
  - Reducing the primary and secondary reformers "proportionally" produced "a shortage of nitrogen … because of the reduced amount of process air". The gap was filled from an air separation unit (ASU).
  - The authors state that "no data were found for an ASU of a quality comparable to the IEAGHG study" (p.22).
  - Nitrogen from ASU/external source: 541 t/d in Case 1B, 727 t/d in Case 2B, 1198 t/d in Case 3 (Table 4, p.24).

**Can be extracted safely:** the reference process-air stream; the fixed-air rule; H2FUTURE's shortage values, for cross-check only.

**Cannot yet be quantified:**
- the nitrogen deficit at S10–S50 in the IEAGHG reference;
- the electricity needed for any substitute nitrogen supply;
- whether the air-blown secondary reformer can operate at reduced throughput.

**Consequence for the workbench:** the reference nitrogen-supply arrangement cannot be assumed adequate in S10–S50. It must enter as a consistency check or limitation. No nitrogen-production model will be built.

*Interpretation:* if reforming is scaled down with the SMR-route hydrogen share, process-air nitrogen falls with it. The shortfall then corresponds roughly to the nitrogen paired with the electrolytic hydrogen.

**For supervisor:** should Stage A report this qualitatively as a limitation, or quantify it with a separately sourced ASU coefficient (which would need a new source)?

## 2. Natural-gas accounts

**What the sources say**
- **IEAGHG streams** (p.33):
  - stream 2 "Natural Gas feedstock to Ammonia Plant": 27021 kg/h;
  - stream 3 "Natural Gas fuel to burners": 10630 kg/h;
  - stream 1 at the battery limit: 37651 kg/h.
- **IEAGHG performance table:** one "Natural Gas (as Fuel)" figure of 10.630 t/h (p.53, Sheet 22).
- **IEAGHG utility table:** fuel of 118,1 MMKcal/h appears only in the Ammonia and Urea Plant row; the Steam/BFW and Utilities rows are blank (p.55).
- **Other fuel inputs to the reformer furnace:**
  - tail gas, stream 4 (p.33);
  - CO₂-removal flash gas, described but not quantified (p.41, Sheet 12).
- **H2FUTURE:** reports feedstock and fuel only (Table 4, p.24).

**Can be extracted safely:** the feedstock account and the "fuel to burners" account (records IEA_HMB_S02_*, IEA_HMB_S03_*, IEA_PD_NG_*).

**Cannot yet be quantified:**
- a separate **other utility fuel** account (IEA_UT_OTHER_FUEL, unquantified);
- the fuel response to reduced reforming.

**Consequence for the workbench:** the reported data do not allow reformer fuel to be separated from other utility fuel. The third account must be carried as "not separately reported", never as zero.

**For supervisor:**
- Accept "not separately reported in reference" for the third account, or seek another source.
- Confirm that the blank utility cells are not to be read as zero.

## 3. CO₂ destinations and vent

**What the sources say**
- **IEAGHG CO₂ balance** (kmol/h equivalent CO₂; p.52, Sheet 21):
  - in: feedstock 1626 + fuel 637 = 2263;
  - out: carbon in urea product 1568 and flue gas 695;
  - "Vents" is printed as "-".
- **"Carbon Not Emitted" label:** the source files urea carbon under this subtotal. It is recorded (IEA_CB_SUBTOTAL_NOT_EMITTED) but **not adopted**.
- **Stream 8, CO₂ to the urea plant** (p.33): 1668 kmol/h; CO₂ 0,9400, H₂O 0,0510, H₂ 0,0080 mol/mol.
- **Purity and pressure statement:** at least 98.5 vol% at 0.50–0.75 bar(g), compressed to about 160 bar(g); basis not stated (p.44, Sheet 15).
- **H2FUTURE** (p.21): "In the IEAGHG-study all CO₂ is used for urea production, and no residual CO₂ is sent to the stack."

**Can be extracted safely:** the input carbon by account; carbon in urea; flue-gas CO₂; the stream 8 flow and composition; the CO₂ quality statement.

**Cannot yet be quantified:**
- the vent flow (a printed dash, not zero);
- the internal carbon routes (flash gas, HRU tail gas);
- the usable CO₂ in scenarios.

**Consequence for the workbench:** no CO₂ shortfall may be computed until Step 2 reconciles the S0 carbon balance. *Interpretation:* if there is no surplus CO₂ at S0, any reduction in reforming-route CO₂ would directly constrain urea.

**For supervisor:**
- Decide how a printed "-" is treated.
- Confirm the wording "carbon carried in urea product".
- Confirm that usable CO₂ must be evidenced, not assumed.

## 4. Ammonia-to-urea relationship

**What the sources say**
- **IEAGHG Case 4A-1:**
  - 1,350 t/d ammonia and 2,260 t/d urea, "about 95% ammonia conversion" (p.27, Sheet 6);
  - ammonia product 68.4 t/d (p.53);
  - streams 9 and 10: 2850 and 53400 kg/h (p.33).
- **H2FUTURE:**
  - "conversion of only 52% of the ammonia" (p.20);
  - urea 1238 t/d (Table 4, p.24);
  - about 800 t/d residual CO₂ vented in Case 1A (p.21).

**Can be extracted safely:** both product splits, kept separately.

**Cannot yet be quantified:** nothing is missing for transcription. Checking that the IEAGHG product and carbon balances agree is Step 2.

**Consequence for the workbench:** the two studies' balances must not be mixed. Their CO₂ constraint behaves differently (IEAGHG has no vent; H2FUTURE has a large surplus).

**For supervisor:** confirm that the IEAGHG roughly 95% relationship is retained for Tier A, and that H2FUTURE is used only as a cross-check.

## 5. Hydrogen delivery point

**What the sources say**
- **IEAGHG stream 7, "Syngas to compression"** (p.33):
  - 7007 kmol/h, 3,00 MPa, 43 °C;
  - H₂ 0,7300, N₂ 0,2580, Ar 0,0030, CH₄ 0,0070, H₂O 0,0020 mol/mol;
  - it is a hydrogen–nitrogen mixture, not pure hydrogen.
- **IEAGHG purity requirement:** CO + CO₂ < 5 ppm to synthesis (p.28, Sheet 7). Oxygen-containing compounds are "severe poison" (p.42, Sheet 13).
- **IEAGHG hydrogen recovery unit:** its hydrogen recycle is shown on the block flow diagram (p.31) but not quantified.
- **H2FUTURE:**
  - electrolytic hydrogen is compressed to 30 bar, with compression separate from the electrolyser system (p.26);
  - no purity, drying or deoxygenation statement.

**Can be extracted safely:** the stream 7 conditions; the purity specification.

**Cannot yet be quantified:**
- whether 3,00 MPa and 30 bar are on the same basis (absolute or gauge is not stated);
- the purity and conditioning of electrolytic hydrogen;
- the recovery-unit recycle flows.

**Consequence for the workbench:** stream 7 is the candidate common delivery boundary. The conditioning load for electrolytic hydrogen remains unquantified.

**For supervisor:** approve stream 7 as the delivery point, and decide how drying and deoxygenation are treated (limitation or sourced load).

## 6. Water

**What the sources say**
- **IRENA definitions** (p.7, Glossary):
  - withdrawal = the water taken from a source;
  - consumption = "the portion of withdrawn water that is not returned to the source".
- **IRENA Table 2.1** (p.31, printed p.29):
  - withdrawal and consumption intensities (L/kg H₂), as average, maximum and minimum;
  - covers Natural gas-SMR (20.01 / 17.54 average), alkaline (32.24 / 22.28) and PEM (25.70 / 17.52) electrolysis, among others.
- **IRENA basis and exclusions:**
  - tap-water source, 66% permeate rate, evaporative cooling (Fig. 2.1 note, p.27);
  - excludes upstream water for electricity and fuel (footnote 3, p.31);
  - data are mostly industry models, not metering (p.25).
- **H2FUTURE:** demineralised water feed only (Table 4, p.24).
- **IEAGHG** (p.55): raw water 22,1 t/h; losses −17,0 t/h; effluent −5,1 t/h; sea water 16.771 m³/h (once-through cooling); closed demineralised-water cooling loop (p.48).

**Can be extracted safely:**
- Level 1 candidates (IRENA alkaline and PEM);
- Level 2 indicative comparison (IRENA SMR);
- Level 3 descriptors for S0 only (IEAGHG).

**Cannot yet be quantified:**
- Level 3 scenario balances;
- adding H2FUTURE feed water to the IRENA totals, which would count reaction water twice.

**Consequence for the workbench:** keep the three levels apart. The IRENA SMR figure is generic: label it indicative for Level 2. IRENA figures must not be used to build a whole-plant balance.

**For supervisor:** choose the IRENA statistic (average with the minimum–maximum range as a sensitivity?), and confirm the Level 2 "indicative" label.

## 7. Electricity

**What the sources say**
- **IEAGHG:**
  - imported power 8.890 MWe, split as ammonia/urea 4.490, steam/BFW 1.100 and utilities/BoP 3.300 (p.53, Sheet 22; p.55);
  - no export is reported;
  - the main compressors are driven by steam turbines (p.47–48, Sheets 16–17).
- **H2FUTURE:**
  - "Electrolyser MWinput" 244 / 328 / 541 MW (Table 4, p.24);
  - "electrolysis plant efficiency of 75% (HHV)" (p.22);
  - electrolyser system cost "excluding compression"; the compressor is separate (p.26);
  - no kWh/kg is stated; ASU electricity is not quantified.

**Can be extracted safely:** the reference electricity by plant area; H2FUTURE's MW inputs and efficiency basis.

**Cannot yet be quantified:**
- a system-level electrolyser kWh/kg covering balance-of-plant, purification and compression;
- the effect of changing compressor drives from steam to electric.

**Consequence for the workbench:** rest-of-plant electricity at S0 is available. A source for the system-level electrolyser coefficient is still missing (Jin et al. 2024 pending). Compression must be a separate line to avoid double counting.

**For supervisor:**
- Approve the electrolyser coefficient source after Jin et al. (2024) is verified.
- Confirm that compression and conditioning are separate lines.

---

### Other recorded issues
- **IEAGHG heat and material balance, column 12:** printed blank (U08).
- **IEAGHG utility table:** nitrogen values printed in parentheses with no definition; this is utility nitrogen, not process nitrogen (U09).
- **IRENA Fig. 2.2:** axis labelled m³/kg although the values match L/kg; use Table 2.1 instead (U10).
- **IRENA SMR data origin:** credited to NETL 2022 in Fig. 2.1 but Lewis et al. (2022) in Fig. 2.2 (U10).
- **Jin et al. (2024):** citation verified; full text pending (U11).
