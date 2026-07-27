#!/usr/bin/env python3
"""
NM PED — F-33 Generator (unified, validated)
=============================================
Replaces the legacy F-33 Excel workbooks (expenditure + revenue) and the
prototype Streamlit app. The engine below reproduced the FY24 submitted
workbooks exactly:

  - Expenditure form (statewide):   90/90 cells to the penny
  - Expenditure per-entity sheet:   11,144/11,144 (8 corrected stale cells)
  - Revenue form (statewide):       50/50 cells to the penny
  - Revenue per-entity sheet:       14,128/14,128 to the penny

Run:   streamlit run f33_app.py
Files: keep xwalk_exp_canonical.csv, xwalk_rev_canonical.csv (and optional
       overrides_exp.csv) in the same folder as this script.

Design rules baked in (learned from FY24 forensics):
  * Keys are normalized: all-digit codes are cast to int-strings, so
    Power BI's '0000 - No Program' matches the xwalk's '0'.
    (The prototype app skipped this and silently dropped $4.26bn.)
  * NOTHING is dropped silently. Unmatched rows are counted, dollared,
    and exported as the new-combos worklist. The app refuses to call a
    run "submission ready" until unmatched $ == 0.
  * Entity matching for overrides is case-insensitive (Excel behavior).
  * Keyless rows (grand-total artifacts) are quarantined and reported.
"""

import glob
import io
import os

import pandas as pd
import streamlit as st

EXP_KEYS = ["Fund", "Function", "Object", "Program", "JobClass"]
REV_KEYS = ["Fund", "Object"]
META_COLS = {"Notes", "Note", "basis", "agreement", "donor_rows"}
CORE_COLS = ["PartII-Total", "Part III", "Part IV"]
# Revenue exclusion rule (per FY24 practice, confirmed against Reiner's
# xwalk where every non-4xxxx object carries no F-33 codes):
#   * any object NOT starting with '4' = balance-sheet item (11111/11112/
#     11113 cash, etc.), not revenue -> excluded
#   * object 41924 = district-to-charter flow-through -> excluded
#     (double count; the charter reports it as its own revenue)
def is_excluded_rev_object(obj: str) -> bool:
    o = str(obj).strip()
    return (not o.startswith("4")) or o == "41924"

# ----------------------------------------------------------------- helpers
def clean_currency(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    v = str(val).replace("$", "").replace(",", "").strip()
    if v.startswith("(") and v.endswith(")"):
        v = "-" + v[1:-1]
    try:
        return float(v)
    except ValueError:
        return 0.0


def norm_code(v) -> str:
    """'11000 - Operational' -> '11000'; '0000' -> '0'; strings pass through."""
    s = str(v).split(" - ")[0].strip()
    return str(int(s)) if s.isdigit() else s


def load_pbi_export(file) -> pd.DataFrame:
    """Read a Power BI export: 2 filter rows, then headers, then data."""
    df = pd.read_csv(file, skiprows=2, low_memory=False)
    df.columns = df.columns.str.strip()
    ren = {"Budget Entity": "Entity", "Actuals": "Actuals YTD",
           "Balances": "Actuals YTD"}
    df = df.rename(columns=ren)
    missing = [c for c in ["Entity", "Fund", "Object", "Actuals YTD"]
               if c not in df.columns]
    if missing:
        raise ValueError(f"Export missing columns {missing}. "
                         f"Columns found: {list(df.columns)}")
    df["amount"] = df["Actuals YTD"].apply(clean_currency)
    for c in EXP_KEYS:
        if c in df.columns:
            df[c] = df[c].map(norm_code)
    df["Entity"] = df["Entity"].fillna("").astype(str).str.strip()
    keyed = [c for c in EXP_KEYS if c in df.columns]
    keyless = df[keyed].eq("").all(axis=1) | df[keyed].eq("nan").all(axis=1)
    dropped = df.loc[keyless, "amount"].sum()
    return df[~keyless].reset_index(drop=True), int(keyless.sum()), dropped


def load_xwalk(path: str, keys: list[str]) -> tuple[pd.DataFrame, list[str]]:
    xw = pd.read_csv(path, dtype=str)
    xw.columns = [c.strip() for c in xw.columns]
    part_cols = [c for c in xw.columns if c not in keys and c not in META_COLS]
    for k in keys:
        xw[k] = xw[k].fillna("").map(norm_code)
    for c in part_cols:
        xw[c] = xw[c].fillna("").astype(str).str.strip().replace("0", "")
    xw = xw.drop_duplicates()
    conflicts = xw[xw.duplicated(subset=keys, keep=False)]
    if len(conflicts):
        raise ValueError(f"Xwalk {os.path.basename(path)} has "
                         f"{len(conflicts)} CONFLICTING duplicate keys.")
    return xw, part_cols


def apply_overrides(merged: pd.DataFrame, ov_path: str, keys: list[str],
                    part_cols: list[str]) -> int:
    """Entity-level exceptions; replaces the xwalk mapping for entity+combo."""
    ov = pd.read_csv(ov_path, dtype=str)
    ov.columns = [c.strip() for c in ov.columns]
    ov["entity_cf"] = ov["entity"].fillna("").str.strip().str.casefold()
    for k in keys:
        ov[k] = ov[k.lower() if k.lower() in ov.columns else k].fillna("").map(norm_code)
    okeys = ["entity_cf"] + keys
    ov_ix = ov.set_index(okeys)
    merged["entity_cf"] = merged["Entity"].str.casefold()
    m_ix = pd.MultiIndex.from_frame(merged[okeys])
    hit = m_ix.isin(ov_ix.index)
    for c in part_cols:
        if c in ov.columns:
            merged.loc[hit, c] = (ov_ix[c].reindex(m_ix[hit]).fillna("").values)
    merged.drop(columns=["entity_cf"], inplace=True)
    return int(hit.sum())


def run_pipeline(df: pd.DataFrame, xw: pd.DataFrame, keys: list[str],
                 part_cols: list[str], ov_path: str | None,
                 exclusion_rule=None):
    merged = df.merge(xw, on=keys, how="left", indicator=True)
    excl_dollars = 0.0
    if exclusion_rule and "Object" in merged.columns:
        is_excl = merged["Object"].map(exclusion_rule)
        excl_dollars = merged.loc[is_excl, "amount"].sum()
        for c in part_cols:
            if c in merged.columns:
                merged.loc[is_excl, c] = ""
        merged.loc[is_excl, "_merge"] = "both"  # known, not unknown
    n_ov = 0
    if ov_path and os.path.exists(ov_path):
        n_ov = apply_overrides(merged, ov_path, keys, part_cols)
    un = merged["_merge"] == "left_only"
    core_present = [c for c in CORE_COLS if c in merged.columns]
    if core_present:
        blank_core = merged["_merge"] == "both"
        for c in core_present:
            blank_core &= (merged[c].fillna("") == "")
        blank_mapped = (merged.loc[blank_core]
                         .groupby(keys, as_index=False)["amount"]
                         .sum().sort_values("amount", ascending=False))
        blank_mapped_dollars = merged.loc[blank_core, "amount"].sum()
    else:
        blank_mapped = pd.DataFrame(columns=keys + ["amount"])
        blank_mapped_dollars = 0.0
 
    new_combos = (merged.loc[un].groupby(keys, as_index=False)["amount"]
                  .sum().sort_values("amount", ascending=False))
    frames = []
    for pc in part_cols:
        if pc not in merged.columns:
            continue
        sub = merged[merged[pc].fillna("") != ""]
        if sub.empty:
            continue
        g = (sub.groupby(["Entity", pc], as_index=False)["amount"].sum()
                .rename(columns={pc: "item_code"}))
        g.insert(1, "part_column", pc)
        frames.append(g)
    long_out = (pd.concat(frames, ignore_index=True)
                if frames else pd.DataFrame(columns=["Entity", "part_column",
                                                     "item_code", "amount"]))
    stats = {
        "rows": len(df),
        "total": df["amount"].sum(),
        "unmatched_rows": int(un.sum()),
        "unmatched_dollars": merged.loc[un, "amount"].sum(),
        "overridden_rows": n_ov,
        "excluded_dollars": excl_dollars,
        "blank_mapped_dollars": blank_mapped_dollars,
    }
    return long_out, new_combos, stats, blank_mapped


def to_matrix(long_out: pd.DataFrame) -> pd.DataFrame:
    """Wide matrix: entities x item codes (codes are unique across parts)."""
    if long_out.empty:
        return pd.DataFrame()
    wide = (long_out.pivot_table(index="Entity", columns="item_code",
                                 values="amount", aggfunc="sum")
            .fillna(0.0))
    return wide


def dl_button(label, df, fname, index=False):
    buf = io.StringIO()
    df.to_csv(buf, index=index)
    st.download_button(label, buf.getvalue(), file_name=fname,
                       mime="text/csv")


# ----------------------------------------------------------------- UI
st.set_page_config(page_title="NM F-33 Generator", layout="wide")
st.title("State of NM — F-33 Generator")
st.caption("Validated against the FY24 submitted workbooks: expenditures "
           "90/90 statewide + 11,144/11,144 per-entity; revenues 50/50 + "
           "14,128/14,128 — all to the penny.")

BASE = os.path.dirname(os.path.abspath(__file__))
xw_exp_path = next(iter(glob.glob(os.path.join(BASE, "xwalk_exp*.csv"))), None)
xw_rev_path = next(iter(glob.glob(os.path.join(BASE, "xwalk_rev*.csv"))), None)
ov_exp_path = next(iter(glob.glob(os.path.join(BASE, "overrides_exp*.csv"))), None)

c1, c2 = st.columns(2)
with c1:
    rev_file = st.file_uploader("Revenue export (CSV)", type=["csv", "txt"])
with c2:
    exp_file = st.file_uploader("Expenditure export (CSV)", type=["csv", "txt"])

if st.button("Generate F-33 Output", type="primary"):
    if not (xw_exp_path and xw_rev_path):
        st.error("Crosswalk file(s) not found next to this script.")
        st.write("Looking in:", BASE)
        st.write(os.listdir(BASE))
        st.stop()
    if not (rev_file and exp_file):
        st.warning("Upload both files.")
        st.stop()

    with st.spinner("Running validated pipelines..."):
        xw_exp, exp_parts = load_xwalk(xw_exp_path, EXP_KEYS)
        xw_rev, rev_parts = load_xwalk(xw_rev_path, REV_KEYS)
        df_exp, exp_keyless, exp_keyless_amt = load_pbi_export(exp_file)
        df_rev, rev_keyless, rev_keyless_amt = load_pbi_export(rev_file)

        exp_long, exp_new, exp_stats, exp_blank = run_pipeline(
            df_exp, xw_exp, EXP_KEYS, exp_parts, ov_exp_path)
        rev_long, rev_new, rev_stats, rev_blank = run_pipeline(
            df_rev, xw_rev, REV_KEYS, rev_parts, None,
            exclusion_rule=is_excluded_rev_object)

    ready = (exp_stats["unmatched_dollars"] == 0
             and exp_stats["blank_mapped_dollars"] == 0
             and rev_stats["unmatched_dollars"] == 0)
    if ready:
        st.success("All dollars mapped. Output is submission-grade.")
    else:
        st.error(
            f"NOT SUBMISSION-READY — unmapped dollars: expenditures "
            f"${exp_stats['unmatched_dollars']:,.2f} ({exp_stats['unmatched_rows']:,} rows), "
            f"blank-mapped dollars: ${exp_stats['blank_mapped_dollars']:,.2f}, "
            f"revenues ${rev_stats['unmatched_dollars']:,.2f} "
            f"({rev_stats['unmatched_rows']:,} rows). Assign F-33 codes to the new "
            f"combos, fix the blank-mapped rows, and rerun.")

    t1, t2, t3, t4 = st.tabs(["Reconciliation", "Statewide", "Per-entity",
                              "New combos"])
    with t1:
        rec = pd.DataFrame([
            ["Expenditures", exp_stats["rows"], f"${exp_stats['total']:,.2f}",
             exp_stats["unmatched_rows"],
             f"${exp_stats['unmatched_dollars']:,.2f}",
             exp_stats["overridden_rows"], exp_keyless,
             f"${exp_keyless_amt:,.2f}"],
            ["Revenues", rev_stats["rows"], f"${rev_stats['total']:,.2f}",
             rev_stats["unmatched_rows"],
             f"${rev_stats['unmatched_dollars']:,.2f}",
             rev_stats["overridden_rows"], rev_keyless,
             f"${rev_keyless_amt:,.2f}"],
        ], columns=["Side", "Detail rows", "Grand total", "Unmatched rows",
                    "Unmatched $", "Override rows", "Keyless rows dropped",
                    "Keyless $ dropped"])
        st.dataframe(rec, use_container_width=True)
        st.caption(f"Revenue intentionally excluded (non-4xxxx balance-sheet "
                   f"objects + 41924 flow-through): "
                   f"${rev_stats.get('excluded_dollars', 0):,.2f} — per FY24 "
                   f"practice, not an error.")
    with t2:
        sw = pd.concat([
            exp_long.assign(side="Expenditure"),
            rev_long.assign(side="Revenue")]).groupby(
            ["side", "part_column", "item_code"], as_index=False)["amount"].sum()
        st.dataframe(sw, use_container_width=True, height=420)
        dl_button("Download statewide CSV", sw, "f33_statewide.csv")
    with t3:
        master = to_matrix(pd.concat([rev_long, exp_long]))
        st.dataframe(master, use_container_width=True, height=420)
        dl_button("Download per-entity long CSV",
                  pd.concat([rev_long.assign(side="Revenue"),
                             exp_long.assign(side="Expenditure")]),
                  "f33_per_entity_long.csv")
        dl_button("Download per-entity matrix CSV", master,
                  "f33_master_matrix.csv", index=True)
    with t4:
        st.write("**Expenditure combos needing F-33 codes** "
                 "(append to xwalk_exp CSV):")
        st.dataframe(exp_new, use_container_width=True, height=260)
        dl_button("Download exp new combos", exp_new, "new_combos_exp.csv")
        st.write("**Revenue combos needing F-33 codes** "
                 "(append to xwalk_rev CSV):")
        st.dataframe(rev_new, use_container_width=True, height=200)
        dl_button("Download rev new combos", rev_new, "new_combos_rev.csv")
        st.write(f"**Expenditure rows matched but PartII-Total/III/IV all "
                 f"blank** (real $ missing from the total, not caught by "
                 f"new-combos): ${exp_stats['blank_mapped_dollars']:,.2f}")
        st.dataframe(exp_blank, use_container_width=True, height=260)
        dl_button("Download blank-mapped exp rows", exp_blank,
                "blank_mapped_exp.csv")