#!/usr/bin/env python3
"""
f33_pipeline.py — NM PED F-33 expenditure pipeline
Replaces the mechanics of the 'FY NEW Exp xwalk' workbook (Reiner Martens).

Inputs
------
1. FACT table (CSV): one row per expenditure line with columns
   entity, fund, function, object, program, jobclass, amount
   (numeric UCoA codes; amount = actuals YTD expenditure)
2. XWALK table (CSV): the '{FY}Xwalk Exp F33' sheet exported as CSV.
   Key columns start with 'FY23 OBMS' (Reiner's legacy header names);
   every other column is an F33 'part' column holding item codes.

Outputs (to --outdir)
---------------------
statewide_f33.csv     item-code totals per part column (the F33 form values)
per_entity_f33.csv    long format: entity, part column, item code, amount
ce3_submittal.csv     per-entity CE3 exhibit (Part X 9 to 13 column)
new_combos.csv        UCoA combos in the fact data with NO xwalk match
                      -> these need F33 codes assigned (the annual task)
recon.txt             reconciliation report

Usage
-----
python f33_pipeline.py --fact fact.csv --xwalk xwalk.csv --outdir out/
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

KEYS = ["fund", "function", "object", "program", "jobclass"]


def load_xwalk(path: Path) -> tuple[pd.DataFrame, list[str]]:
    xw = pd.read_csv(path, dtype=str)
    xw.columns = [c.strip() for c in xw.columns]
    key_cols = [c for c in xw.columns if c.upper().startswith("FY") and "OBMS" in c.upper()]
    if len(key_cols) != 5:
        sys.exit(f"Expected 5 'FY.. OBMS' key columns in xwalk, found {len(key_cols)}: {key_cols}")
    part_cols = [c for c in xw.columns
                 if c not in key_cols and not c.startswith("Unnamed")]
    xw = xw.rename(columns=dict(zip(key_cols, KEYS)))
    for k in KEYS:
        xw[k] = xw[k].fillna("").astype(str).str.strip()
    # Reiner's xwalk contains exact duplicate rows (harmless). Conflicting
    # duplicates (same key, different codes) are a data error -> fail loudly.
    xw = xw.drop_duplicates()
    conflicts = xw[xw.duplicated(subset=KEYS, keep=False)]
    if len(conflicts):
        conflicts.to_csv("xwalk_conflicts.csv", index=False)
        sys.exit(f"XWALK ERROR: {len(conflicts)} conflicting duplicate keys "
                 f"(same UCoA combo, different F33 codes). See xwalk_conflicts.csv")
    for c in part_cols:
        xw[c] = xw[c].fillna("").astype(str).str.strip().replace("0", "")
    return xw, part_cols


def load_fact(path: Path) -> pd.DataFrame:
    fact = pd.read_csv(path, dtype=str)
    fact.columns = [c.strip().lower() for c in fact.columns]
    needed = ["entity"] + KEYS + ["amount"]
    missing = [c for c in needed if c not in fact.columns]
    if missing:
        sys.exit(f"Fact table missing columns: {missing}. Needed: {needed}")
    for k in ["entity"] + KEYS:
        fact[k] = fact[k].fillna("").astype(str).str.strip()
    fact["amount"] = pd.to_numeric(fact["amount"], errors="coerce").fillna(0.0)
    # Normalize numeric-looking keys ('11000.0' -> '11000', '0' stays '0')
    for k in KEYS:
        fact[k] = fact[k].str.replace(r"\.0$", "", regex=True)
    # Drop rows with no keys at all (grand-total artifacts from exports)
    keyless = fact[KEYS].eq("").all(axis=1)
    dropped_total = fact.loc[keyless, "amount"].sum()
    if keyless.any():
        print(f"NOTE: dropping {keyless.sum()} keyless row(s) carrying "
              f"${dropped_total:,.2f} (grand-total artifacts).")
    return fact[~keyless].reset_index(drop=True)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--fact", required=True, type=Path)
    ap.add_argument("--xwalk", required=True, type=Path)
    ap.add_argument("--outdir", default=Path("f33_out"), type=Path)
    ap.add_argument("--overrides", type=Path, default=None,
                    help="Optional CSV of entity-level code overrides: columns "
                         "entity,fund,function,object,program,jobclass + part columns. "
                         "Rows replace the xwalk mapping for that entity+combo only.")
    args = ap.parse_args()
    args.outdir.mkdir(parents=True, exist_ok=True)

    xw, part_cols = load_xwalk(args.xwalk)
    fact = load_fact(args.fact)
    grand_total = fact["amount"].sum()

    merged = fact.merge(xw, on=KEYS, how="left", indicator=True)
    n_overridden = 0
    if args.overrides:
        ov = pd.read_csv(args.overrides, dtype=str)
        ov.columns = [c.strip() for c in ov.columns]
        for c in ["entity"] + KEYS:
            ov[c] = ov[c].fillna("").astype(str).str.strip()
        ov_parts = [c for c in ov.columns if c in part_cols]
        okeys = ["entity"] + KEYS
        ov_ix = ov.set_index(okeys)
        m_ix = pd.MultiIndex.from_frame(merged[okeys])
        hit = m_ix.isin(ov_ix.index)
        n_overridden = int(hit.sum())
        for c in ov_parts:
            vals = ov_ix[c].reindex(m_ix[hit]).fillna("").values
            merged.loc[hit, c] = vals
        for c in part_cols:
            merged[c] = merged[c].fillna("").astype(str).str.strip().replace("0", "")
    un = merged["_merge"] == "left_only"
    unmatched_dollars = merged.loc[un, "amount"].sum()

    # THE annual maintenance list: combos needing F33 code assignment
    new_combos = (merged.loc[un, KEYS + ["amount"]]
                  .groupby(KEYS, as_index=False)["amount"].sum()
                  .sort_values("amount", ascending=False))
    new_combos.to_csv(args.outdir / "new_combos.csv", index=False)

    # Long-format outputs: one row per (entity, part column, item code)
    frames = []
    for pc in part_cols:
        sub = merged[merged[pc].fillna("") != ""]
        if sub.empty:
            continue
        g = (sub.groupby(["entity", pc], as_index=False)["amount"].sum()
                .rename(columns={pc: "item_code"}))
        g.insert(1, "part_column", pc)
        frames.append(g)
    per_entity = pd.concat(frames, ignore_index=True)
    per_entity = per_entity.sort_values(["entity", "part_column", "item_code"])
    per_entity.to_csv(args.outdir / "per_entity_f33.csv", index=False)

    statewide = (per_entity.groupby(["part_column", "item_code"], as_index=False)
                 ["amount"].sum())
    statewide.to_csv(args.outdir / "statewide_f33.csv", index=False)

    ce3 = per_entity.query("item_code == 'CE3'")[["entity", "amount"]]
    ce3.to_csv(args.outdir / "ce3_submittal.csv", index=False)

    # Reconciliation report
    lines = [
        "F-33 EXPENDITURE PIPELINE — RECONCILIATION",
        f"Fact rows (detail):            {len(fact):,}",
        f"Grand total expenditures:      ${grand_total:,.2f}",
        f"Rows without xwalk match:      {int(un.sum()):,}",
        f"Rows with entity overrides:    {n_overridden:,}",
        f"Dollars without xwalk match:   ${unmatched_dollars:,.2f}",
        f"Unique new combos to map:      {len(new_combos):,} (see new_combos.csv)",
        "",
        "CHECK: every part column's coded dollars must be <= grand total;",
        "codes are overlapping categories (a dollar feeds multiple items).",
    ]
    for pc in part_cols:
        coded = per_entity.loc[per_entity["part_column"] == pc, "amount"].sum()
        lines.append(f"  {pc:<28s} ${coded:,.2f}")
    report = "\n".join(lines)
    (args.outdir / "recon.txt").write_text(report)
    print(report)
    if unmatched_dollars != 0:
        print("\nWARNING: unmatched dollars are EXCLUDED from all outputs "
              "until mapped in the xwalk. Do not submit until new_combos.csv "
              "is resolved and this run is repeated.")


if __name__ == "__main__":
    main()
