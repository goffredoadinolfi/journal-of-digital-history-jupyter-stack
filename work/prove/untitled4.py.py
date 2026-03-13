#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
INPUT
-----
- senatori_regno_1848_1943_master_dataset_final_v9.xlsx

OUTPUT
------
# TABLE 1 (Art.33 ALL categories found: cat_01...cat_XX), by period
- table_art33_all_by_period_counts_v9.csv
- table_art33_all_by_period_pct_v9.csv
- table_art33_all_by_period_bw.png

# TABLE 2 (ALL professions in profession_main_en), by period
- table_prof_all_by_period_counts_v9.csv
- table_prof_all_by_period_pct_v9.csv
- table_prof_all_by_period_bw.png

What the script does
--------------------
- Applies the same "in office" rule by period:
  nomination_dt <= period_end AND (effective_death_dt missing OR effective_death_dt >= period_start)
- Applies the same missing-death rule (A/B/C).
- Builds TWO tables for each block (counts + percentages), with:
  - a TOTAL column (n / 100%),
  - a TOTAL row (sum across periods),
  - a "100%" check row at the bottom of the % table (row-sum check).

Notes
-----
- Art.33 categories are detected automatically as columns starting with "cat_".
- Professions are taken from profession_main_en; blanks are labelled "Missing".
- Table PNGs are black-and-white, with wrapped headers so they stay inside cells.
"""

from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# =========================
# CONFIG / I-O
# =========================
INPUT_FILE = "senatori_regno_1848_1943_master_dataset_final_v9.xlsx"
INPUT_SHEET = 0

OUT_CAT_COUNTS_CSV = "table_art33_all_by_period_counts_v9.csv"
OUT_CAT_PCT_CSV    = "table_art33_all_by_period_pct_v9.csv"
OUT_CAT_PNG        = "table_art33_all_by_period_bw.png"

OUT_PROF_COUNTS_CSV = "table_prof_all_by_period_counts_v9.csv"
OUT_PROF_PCT_CSV    = "table_prof_all_by_period_pct_v9.csv"
OUT_PROF_PNG        = "table_prof_all_by_period_bw.png"

EXPORT_PNG = True

PERIODS = [
    ("1848–1859", 1848, 1859),
    ("1860–1882", 1860, 1882),
    ("1883–1913", 1883, 1913),
    ("1914–1924", 1914, 1924),
    ("1925–1946", 1925, 1946),
]

TRUTHY_STRINGS = {"x", "1", "true", "yes", "si", "sì", "ok"}
COL_PROF = "profession_main_en"

# Table appearance
FONT_TABLE = 8
HEADER_FONT_TABLE = 8
HEADER_HEIGHT_MULT = 1.55
ROW_HEIGHT_MULT = 1.20

# For very wide tables (many professions), keep readable width in PNG:
MAX_COLS_IN_PNG = 18   # if more columns exist, PNG renders top-N + "Other" (CSVs still include ALL)

# =========================
# HELPERS
# =========================
def pick_first_existing(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def parse_date_col(series):
    s = series.astype("string").str.strip()
    dt = pd.to_datetime(s, errors="coerce")
    m = dt.isna() & s.notna() & (s != "")
    if m.any():
        dt.loc[m] = pd.to_datetime(s[m], dayfirst=True, errors="coerce")
    return dt

def is_truthy_cell(v) -> bool:
    if pd.isna(v):
        return False
    if isinstance(v, (bool, np.bool_)):
        return bool(v)
    if isinstance(v, (int, np.integer)):
        return int(v) == 1
    if isinstance(v, (float, np.floating)):
        return float(v) == 1.0
    s = str(v).strip().lower()
    return s in TRUTHY_STRINGS

def in_office_mask(df, start_dt, end_dt):
    return (
        df["nomination_dt"].notna()
        & (df["nomination_dt"] <= end_dt)
        & (df["effective_death_dt"].isna() | (df["effective_death_dt"] >= start_dt))
    )

def apply_missing_death_rule(df):
    cutoff_sardinia_end = pd.Timestamp(year=1859, month=12, day=31)
    cutoff_fascism_keep = pd.Timestamp(year=1929, month=1, day=1)

    missing_death = df["effective_death_dt"].isna() & df["nomination_dt"].notna()
    mask_A = missing_death & (df["nomination_dt"] <= cutoff_sardinia_end)
    mask_B = missing_death & (df["nomination_dt"] >= cutoff_fascism_keep)
    mask_C = missing_death & ~(mask_A | mask_B)

    df.loc[mask_A, "effective_death_dt"] = cutoff_sardinia_end
    df.loc[mask_C, "effective_death_dt"] = pd.Timestamp(year=1800, month=1, day=1)  # excluded
    # mask_B stays NaT (kept alive)
    return mask_A, mask_B, mask_C

def wrap_label(label: str, width: int = 16) -> str:
    s = str(label).strip()
    if len(s) <= width:
        return s
    words = s.split()
    lines, cur = [], ""
    for w in words:
        if len(cur) + (1 if cur else 0) + len(w) <= width:
            cur = (cur + " " + w).strip()
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return "\n".join(lines)

def add_totals_counts(df_counts: pd.DataFrame) -> pd.DataFrame:
    out = df_counts.copy()
    out["TOTAL (n)"] = out.sum(axis=1).astype(int)
    total_row = pd.DataFrame(out.sum(axis=0)).T
    total_row.index = ["TOTAL (n)"]
    total_row = total_row.astype(int)
    out = pd.concat([out, total_row], axis=0)
    return out

def add_totals_pct(df_pct: pd.DataFrame) -> pd.DataFrame:
    out = df_pct.copy()
    out["TOTAL (%)"] = 100.0

    # Row-sum check across substantive columns (exclude TOTAL)
    base_cols = [c for c in out.columns if c != "TOTAL (%)"]
    out["Row sum (%)"] = out[base_cols].sum(axis=1)

    # TOTAL row (grand distribution across all periods, weighted by counts)
    # We compute it from the corresponding counts later; here we keep a placeholder and fill in main().
    return out

def render_bw_table_png(df_to_show: pd.DataFrame, title: str, out_png: str):
    show_df = df_to_show.copy()

    col_labels = [wrap_label(c, width=18) for c in show_df.columns.tolist()]
    row_labels = [wrap_label(r, width=14) for r in show_df.index.tolist()]

    # strings: ints stay ints, floats 1 decimal
    cell_text = []
    for _, row in show_df.iterrows():
        row_cells = []
        for v in row.values:
            if pd.isna(v):
                row_cells.append("")
            elif isinstance(v, (int, np.integer)):
                row_cells.append(f"{int(v)}")
            elif isinstance(v, (float, np.floating)):
                row_cells.append(f"{float(v):.1f}")
            else:
                try:
                    fv = float(v)
                    row_cells.append(f"{fv:.1f}")
                except Exception:
                    row_cells.append(str(v))
        cell_text.append(row_cells)

    fig_w = max(9.5, 2.2 + 1.15 * show_df.shape[1])
    fig_h = max(3.2, 1.4 + 0.72 * show_df.shape[0])
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis("off")

    table = ax.table(
        cellText=cell_text,
        rowLabels=row_labels,
        colLabels=col_labels,
        cellLoc="center",
        loc="center",
    )

    table.auto_set_font_size(False)
    table.set_fontsize(FONT_TABLE)

    # Column widths (row labels wider)
    ncols = show_df.shape[1]
    row_label_width = 0.22
    data_col_width = (1.0 - row_label_width) / max(1, ncols)

    for (r, c), cell in table.get_celld().items():
        if c == -1:
            cell.set_width(row_label_width)
        else:
            cell.set_width(data_col_width)

        cell.set_facecolor("white")
        cell.set_edgecolor("black")
        cell.set_linewidth(0.8 if (r == 0 or c == -1) else 0.6)

        if r == 0:
            cell.set_text_props(weight="bold", fontsize=HEADER_FONT_TABLE, va="center")
            cell.set_height(cell.get_height() * HEADER_HEIGHT_MULT)
        else:
            cell.set_height(cell.get_height() * ROW_HEIGHT_MULT)

        if c == -1:
            cell.set_text_props(weight="bold", va="center")

        cell.get_text().set_va("center")
        cell.get_text().set_ha("center")

    ax.set_title(title, fontsize=10, pad=14)

    plt.tight_layout()
    if EXPORT_PNG:
        fig.savefig(out_png, dpi=300, bbox_inches="tight")
    plt.show()

def compress_for_png(df: pd.DataFrame, max_cols: int) -> pd.DataFrame:
    """
    For PNG only: if too many columns, keep the leftmost (max_cols-1) and merge the rest into 'Other'.
    CSVs will still include ALL columns.
    """
    if df.shape[1] <= max_cols:
        return df
    keep = df.columns.tolist()[: max_cols - 1]
    rest = [c for c in df.columns.tolist() if c not in keep]
    out = df[keep].copy()
    out["Other (collapsed)"] = df[rest].sum(axis=1)
    return out

# =========================
# MAIN
# =========================
def main():
    if not Path(INPUT_FILE).is_file():
        raise FileNotFoundError(f"[STOP] Input file not found: {INPUT_FILE}")

    df = pd.read_excel(INPUT_FILE, sheet_name=INPUT_SHEET)

    # nomination/death columns (robust)
    col_nomina = pick_first_existing(
        df,
        ["data_nomina_dt", "data_nomina", "nomination_date_dt", "nomination_date", "nomination_dt"]
    )
    col_morte = pick_first_existing(
        df,
        ["effective_death_dt", "death_date_dt", "data_decesso_dt", "data_decesso", "death_date_raw", "death_dt"]
    )
    if col_nomina is None:
        raise ValueError("[STOP] Nomination date column missing.")
    if col_morte is None:
        raise ValueError("[STOP] Death/effective death column missing.")

    df["nomination_dt"] = parse_date_col(df[col_nomina])

    if col_morte == "effective_death_dt":
        df["effective_death_dt"] = pd.to_datetime(df[col_morte], errors="coerce")
    else:
        df["effective_death_dt"] = parse_date_col(df[col_morte])

    mask_A, mask_B, mask_C = apply_missing_death_rule(df)

    # =========================================================
    # TABLE 1: Art.33 ALL categories (auto-detected cat_*)
    # =========================================================
    cat_cols = [c for c in df.columns if str(c).startswith("cat_")]
    if not cat_cols:
        raise KeyError("[STOP] No Art.33 category columns found (expected columns starting with 'cat_').")

    # Keep a stable order: cat_01, cat_02, ...
    def cat_sort_key(x):
        try:
            return int(str(x).split("_")[1])
        except Exception:
            return 999
    cat_cols = sorted(cat_cols, key=cat_sort_key)

    cat_counts_rows, cat_pct_rows = [], []
    totals_by_period = {}

    for label, y0, y1 in PERIODS:
        start = pd.Timestamp(year=y0, month=1, day=1)
        end = pd.Timestamp(year=y1, month=12, day=31)

        m = in_office_mask(df, start, end)
        block = df.loc[m, cat_cols].copy()
        total = int(len(block))
        totals_by_period[label] = total

        counts = {}
        for c in cat_cols:
            counts[c] = int(block[c].apply(is_truthy_cell).sum()) if total else 0
        cat_counts_rows.append(pd.Series(counts, name=label))

        pcts = {k: (100.0 * v / total) if total else 0.0 for k, v in counts.items()}
        cat_pct_rows.append(pd.Series(pcts, name=label))

    cat_counts = pd.DataFrame(cat_counts_rows).fillna(0).astype(int)
    cat_pct = pd.DataFrame(cat_pct_rows).fillna(0.0)

    cat_counts_out = add_totals_counts(cat_counts)
    cat_pct_out = add_totals_pct(cat_pct)

    # Fill TOTAL row in % table as a grand distribution (weighted by counts)
    grand_total = int(cat_counts.sum(axis=0).sum())
    if grand_total > 0:
        grand_counts = cat_counts.sum(axis=0)
        grand_pct = (100.0 * grand_counts / grand_total).to_dict()
    else:
        grand_pct = {c: 0.0 for c in cat_cols}

    total_pct_row = pd.Series({**grand_pct, "TOTAL (%)": 100.0, "Row sum (%)": sum(grand_pct.values())}, name="TOTAL (%)")
    cat_pct_out = pd.concat([cat_pct_out, total_pct_row.to_frame().T], axis=0)

    # Add explicit "100%" check row under TOTAL
    check_row = pd.Series({c: np.nan for c in cat_cols}, name="100%")
    check_row["TOTAL (%)"] = 100.0
    check_row["Row sum (%)"] = np.nan
    cat_pct_out = pd.concat([cat_pct_out, check_row.to_frame().T], axis=0)

    cat_counts_out.to_csv(OUT_CAT_COUNTS_CSV, encoding="utf-8-sig")
    cat_pct_out.to_csv(OUT_CAT_PCT_CSV, encoding="utf-8-sig")

    # PNG (compressed if too wide)
    cat_pct_png = compress_for_png(cat_pct_out, MAX_COLS_IN_PNG)
    render_bw_table_png(
        cat_pct_png,
        title="Art.33 categories by period (percent of senators in office) — with totals and 100% check",
        out_png=OUT_CAT_PNG
    )

    # =========================================================
    # TABLE 2: ALL professions (profession_main_en)
    # =========================================================
    if COL_PROF not in df.columns:
        raise KeyError(f"[STOP] Missing column: {COL_PROF}")

    prof_counts_rows, prof_pct_rows = [], []
    # Build full list of professions across union of in-office (to keep columns manageable but complete)
    all_in_office = pd.Series(False, index=df.index)
    for _, y0, y1 in PERIODS:
        start = pd.Timestamp(year=y0, month=1, day=1)
        end = pd.Timestamp(year=y1, month=12, day=31)
        all_in_office = all_in_office | in_office_mask(df, start, end)

    prof_all = df.loc[all_in_office, COL_PROF].astype("string").fillna("").str.strip()
    prof_all = prof_all.replace("", "Missing")
    prof_levels = prof_all.value_counts().index.tolist()  # sorted by overall frequency

    prof_totals_by_period = {}

    for label, y0, y1 in PERIODS:
        start = pd.Timestamp(year=y0, month=1, day=1)
        end = pd.Timestamp(year=y1, month=12, day=31)

        m = in_office_mask(df, start, end)
        block = df.loc[m, [COL_PROF]].copy()
        total = int(len(block))
        prof_totals_by_period[label] = total

        s = block[COL_PROF].astype("string").fillna("").str.strip()
        s = s.replace("", "Missing")

        counts = {p: int((s == p).sum()) for p in prof_levels}
        prof_counts_rows.append(pd.Series(counts, name=label))

        pcts = {k: (100.0 * v / total) if total else 0.0 for k, v in counts.items()}
        prof_pct_rows.append(pd.Series(pcts, name=label))

    prof_counts = pd.DataFrame(prof_counts_rows).fillna(0).astype(int)[prof_levels]
    prof_pct = pd.DataFrame(prof_pct_rows).fillna(0.0)[prof_levels]

    prof_counts_out = add_totals_counts(prof_counts)
    prof_pct_out = add_totals_pct(prof_pct)

    # TOTAL row in % table as grand distribution (weighted by counts)
    grand_total_prof = int(prof_counts.sum(axis=0).sum())
    if grand_total_prof > 0:
        grand_counts_prof = prof_counts.sum(axis=0)
        grand_pct_prof = (100.0 * grand_counts_prof / grand_total_prof).to_dict()
    else:
        grand_pct_prof = {p: 0.0 for p in prof_levels}

    total_pct_row_prof = pd.Series(
        {**grand_pct_prof, "TOTAL (%)": 100.0, "Row sum (%)": sum(grand_pct_prof.values())},
        name="TOTAL (%)"
    )
    prof_pct_out = pd.concat([prof_pct_out, total_pct_row_prof.to_frame().T], axis=0)

    check_row_prof = pd.Series({p: np.nan for p in prof_levels}, name="100%")
    check_row_prof["TOTAL (%)"] = 100.0
    check_row_prof["Row sum (%)"] = np.nan
    prof_pct_out = pd.concat([prof_pct_out, check_row_prof.to_frame().T], axis=0)

    prof_counts_out.to_csv(OUT_PROF_COUNTS_CSV, encoding="utf-8-sig")
    prof_pct_out.to_csv(OUT_PROF_PCT_CSV, encoding="utf-8-sig")

    # PNG (compressed if too wide)
    prof_pct_png = compress_for_png(prof_pct_out, MAX_COLS_IN_PNG)
    render_bw_table_png(
        prof_pct_png,
        title="Professions by period (percent of senators in office) — with totals and 100% check",
        out_png=OUT_PROF_PNG
    )

    # =========================================================
    # Compact prints
    # =========================================================
    print(f"[OK] Input: {INPUT_FILE} | Rows read: {len(df)}")
    print(f"[OK] Missing-death rule -> A: {int(mask_A.sum())} | B: {int(mask_B.sum())} | C: {int(mask_C.sum())}")
    print(f"[OK] Art.33 categories detected: {len(cat_cols)} columns")
    print(f"[OK] Saved Art.33 CSVs: {OUT_CAT_COUNTS_CSV} | {OUT_CAT_PCT_CSV}")
    if EXPORT_PNG:
        print(f"[OK] Saved Art.33 PNG:  {OUT_CAT_PNG}")
    print(f"[OK] Professions detected (in-office union): {len(prof_levels)} unique labels")
    print(f"[OK] Saved Prof CSVs: {OUT_PROF_COUNTS_CSV} | {OUT_PROF_PCT_CSV}")
    if EXPORT_PNG:
        print(f"[OK] Saved Prof PNG:  {OUT_PROF_PNG}")
    if len(prof_levels) > MAX_COLS_IN_PNG:
        print(f"[NOTE] PNG columns were compressed (CSV still contains ALL professions).")

if __name__ == "__main__":
    main()