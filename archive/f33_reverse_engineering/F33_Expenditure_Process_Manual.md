# F-33 Expenditure Reporting — Process Manual (Rebuilt)

**Report:** U.S. Census Bureau Annual Survey of School System Finances (Form F-33), expenditure side
**Predecessor process:** Reiner Martens, PED — "FY24 NEW Exp xwalk v2.xlsx" + 2-page notes
**This version:** Decoded, validated, and rebuilt in Python, July 2026 (Lorenzo Dominguez / Claude)
**Validation:** Python replica reproduced Reiner's FY24 workbook exactly — 90/90 statewide form cells and 11,136/11,144 per-entity cells to the penny. The 8 non-ties were stale cached zeros in Reiner's per-entity sheet (CE3 for eight RECs); the replica's values match his newer CE3 Submittal tab, confirming the replica is more current than the cache.

---

## 1. What the process is

F-33 has its own item codes (Z33, E13, V10, CE1, …) organized into form "Parts." OBMS
data lives in NM UCoA coordinates (fund / function / object / program / job class).
The entire expenditure process is three operations:

1. **Extract** actuals expenditures from OBMS as a flat table:
   one row per entity + UCoA combination, with the dollar amount.
2. **Crosswalk**: each unique UCoA combination maps to one or more F-33 item codes
   (a salary dollar simultaneously feeds Z33 instructional salaries, E13 instruction
   spending, Z32 total salaries, Z35 teacher salaries, CE1 current expenditure, …).
   The mapping table is the "xwalk" — the single most valuable asset in the process.
3. **Aggregate**: group-by-sum dollars per item code, statewide and per entity.

Everything in Reiner's 105 MB workbook — the ~13,000 SUMIFS formulas — was these
three operations implemented in Excel.

## 2. The new pipeline

`f33_pipeline.py` replaces the workbook mechanics. Runtime: seconds.

```
python f33_pipeline.py --fact fact.csv --xwalk xwalk.csv \
                       [--overrides overrides.csv] --outdir out/
```

**Inputs**
- `fact.csv` — columns: `entity, fund, function, object, program, jobclass, amount`
  (numeric UCoA codes, actuals YTD expenditure). Source: OBMS export / Power BI /
  Nova pipeline. Final year-end actuals, all funds.
- `xwalk.csv` — Reiner's xwalk sheet as CSV: five `FY23 OBMS.*` key columns + one
  column per F-33 form part, cells holding item codes (blank = not applicable).
- `overrides.csv` (optional) — entity-level exceptions: same shape as the xwalk plus
  an `entity` column. A row replaces the xwalk mapping for that entity + combo only.

**Outputs**
- `statewide_f33.csv` — totals per (part column, item code): the form values.
- `per_entity_f33.csv` — long format per entity: the Census submission detail.
- `ce3_submittal.csv` — the per-REC CE3 exhibit (RESA expenditures).
- `new_combos.csv` — **the annual maintenance list**: UCoA combos present in the data
  with no xwalk match. Assign F-33 codes to these, append to the xwalk, rerun.
  Do not submit while this file is non-empty or unmatched dollars ≠ 0.
- `recon.txt` — reconciliation report (grand total, unmatched dollars, per-part sums).

**Built-in safeguards** (replacing Reiner's manual steps 2–4)
- Keyless rows (e.g., the podhoc Grand Total row) are quarantined and reported —
  they carried $7.2bn in FY24 and would double-count a naive sum.
- Conflicting xwalk duplicates (same combo, different codes) abort the run.
  Exact duplicate rows (340 in the FY24 xwalk) are deduped silently.
- Entity matching is case-insensitive (Excel SUMIFS behavior; "The Ask Academy"
  vs "The ASK Academy" was a live example).

## 3. FY24 facts worth knowing

- FY24 statewide expenditure detail total: **$7,216,563,403.63** (ties the OBMS
  podhoc grand total to the penny).
- **Two hand-edits existed in FY24** (see `hand_edits_review.csv`, encoded in
  `fy24_overrides.csv`):
  1. NM International School — seven benefits rows (objects 52220, 52311; $54,823
     total) were deliberately blanked, excluding them from all F-33 items.
     Rationale undocumented. Decide whether this carries into FY25.
  2. Raton Public Schools — one row (fund 25153, function 2100, object 56118;
     $1,257.11) stamped V12/E17/Z34/CE1/SE1/SE3 in place of the xwalk mapping.
- The xwalk column headers drifted from the workbook's Data-Exp headers for the
  Part XIII columns (one is literally named "AE3"). The xwalk column names are
  canonical; alignment was verified empirically at 99.99–100% per column.

## 4. What the expenditure pipeline does NOT cover

The F-33 form needs four more feeds, each with its own (partially documented) process:

1. **Revenues (Part I, Part XII revenue section)** — lived in separate workbooks
   ("FY24 OBMS Data Revenues.xlsx", "NM FY24 F33 revenue v1/v2"). All revenue links
   in the expenditure workbook are broken (#REF!). Must be reconstructed from the
   revenue workbooks before FY25.
2. **Debt (Part VI)** — beginning-of-year outstanding debt and issuances (revenue
   code 5110) are HAND-ENTERED from debt schedules (955F world). Retirements are
   computed from object 831 (FY24: $430,862,433.01, reproduced by the pipeline).
   Ending = beginning + issued − retired. FY24 values: beginning $2,296,428,359.56;
   issued $390,886,654.68; ending $2,256,452,581.23.
3. **Cash and investments (Part VII)** — from the OBMS cash report, classified by
   fund prefix: funds 1xxxx/2xxxx → W61 (other funds); 3xxxx → W31 (bond funds);
   else → W01 (sinking/debt service). FY24: W01 $620,567,138.58;
   W31 $1,693,214,268.07; W61 $1,265,838,982.13.
4. **Fall membership (Part VIII, item V33)** — straight sum of the 80-day fall
   MEM per entity (FY24 statewide: 302,757). Source: Nova/SSRS membership pulls.

**Open question inherited from Reiner:** cell AE179 on the form held a hand-keyed
**$5,083,488.05** subtracted from the ARP ESSER revenue line (Part XII, code AR1B).
Rationale unknown — reconstruct from the revenue workbook or Census edit
correspondence before reusing.

## 5. FY25 runbook

1. Export final FY25 actuals expenditures (all funds, all entities) as the flat
   fact table. Verify the export total against a known control total.
2. Copy last year's xwalk. Run the pipeline. Open `new_combos.csv` — assign F-33
   codes to any new UCoA combinations (new funds, job classes, UCoA changes),
   append to the xwalk, rerun until unmatched dollars = 0.
3. Review whether the FY24 overrides still apply; edit `overrides.csv`.
4. Produce the other four feeds (revenues, debt, cash, membership).
5. Reconcile: statewide totals vs prior year (large swings need explanations —
   Census will ask); per-entity totals vs OBMS actuals; CE3 vs REC records.
6. Assemble into the Census submission format ("NM FY25 F33" workbook lineage)
   and keep the run inputs/outputs as the audit trail.

## 6. Reference: F-33 item codes seen in the NM xwalk

Selected, from Census F-33 documentation: E13 current spending–instruction;
Z32 total salaries; Z33 salaries–instruction (incl. aides/substitutes);
Z34 total employee benefits; Z35/Z36/Z37/Z38 teacher salaries for regular /
special / vocational / other programs; V10 benefits–instruction; V33 fall
membership; CE1/CE2/CE3 current expenditure exhibits (CE3 = RESA expenditures,
all funds, functions 1000/2000/3100/3200, objects 100–600, 810, 820, 890);
SE1–SE5 special exhibit items; AR1/AR1A/AR1B/AR2 ESSER I / ESSER II / ARP ESSER /
GEER revenues; W01/W31/W61 cash by fund type; 19H/21F/31F/41F long-term debt.
