========================================================================
F33_App  —  PRIMARY TOOL  (annual F-33 populate engine)
========================================================================
THIS IS THE ONE YOU RUN EACH YEAR.

Quick test to know you're in the right place:
  Does the script read Power BI exports?  ->  YES, this one.
  (If a script reads a legacy .xlsx instead, that's Extract_f33.py in
   the F33_Xwalk folder — the archive. See that folder's README.)

------------------------------------------------------------------------
WHAT IT DOES
------------------------------------------------------------------------
Takes two Power BI actuals exports (revenue + expenditure), merges them
against the canonical crosswalks, and produces the F-33 submission
material:
  - statewide totals (Census form order)
  - per-entity matrix
  - new-combos worklist (UCOA combos not yet in the xwalk)
  - reconciliation that will NOT say "submission-ready" until unmatched
    dollars == $0.00 on both sides.

------------------------------------------------------------------------
HOW TO RUN
------------------------------------------------------------------------
  cd F33_App
  source venv/bin/activate        (if using the venv here)
  streamlit run f33_app.py

Then in the browser: upload the Revenue export and the Expenditure
export, click "Generate F-33 Output."

------------------------------------------------------------------------
FILES THAT MUST SIT NEXT TO f33_app.py
------------------------------------------------------------------------
  xwalk_exp_canonical.csv     (expenditure crosswalk — keys + census cols)
  xwalk_rev_canonical.csv     (revenue crosswalk — keys + census cols)
  overrides_exp.csv           (optional; entity-level exceptions)

Keep exactly ONE xwalk_exp*.csv and ONE xwalk_rev*.csv here. The app
grabs the first match of each by filename.

IMPORTANT — keep the canonical xwalks LEAN.
The app treats every column that isn't a key and isn't a meta column as
a Census destination and will sum dollars into it. So do NOT drop the
NCES/name-enriched crosswalk in here as the canonical file — those extra
columns (NCES_Code, NCES_Name, Object_Name, Census_Desc, etc.) would be
read as fake Census line items. Keep enrichment in a separate companion
file. If you ever must carry descriptive columns in the canonical file,
add their names to META_COLS in f33_app.py first.

------------------------------------------------------------------------
ANNUAL CYCLE
------------------------------------------------------------------------
  1. Export final-year actuals from Power BI (revenue + expenditure).
  2. streamlit run f33_app.py  ->  upload both.
  3. Review the "New combos" tab. Assign F-33 codes to each, append the
     rows to the canonical xwalk CSVs.
  4. Rerun until the banner is green (unmatched $ == 0).
  5. Refresh the population workbook, fill the 5 manual cells, transcribe
     to the Census form, reasonableness pass vs. prior year, submit.

========================================================================
