#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
INPUT
-----
- senatori_regno_1848_1943_master_dataset_final_v9.xlsx

OUTPUT
------
# TABLE 1 (Art.33 two dominant categories, by period) - PERCENT + TOTAL/100, periods in columns
- table_art33_2cats_by_period_pct_v9.csv
- table_art33_2cats_by_period_bw_v9.png

# TABLE 2 (Top-5 professione_macro, by period) - PERCENT + TOTAL/100, periods in columns
- table_prof_macro_top5_by_period_pct_v9.csv
- table_prof_macro_top5_by_period_bw_v9.png
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

OUT_CAT_PCT_CSV = "table_art33_2cats_by_period_pct_v9.csv"
OUT_CAT_PNG     = "table_art33_2cats_by_period_bw_v9.png"

OUT_PROF_PCT_CSV = "table_prof_macro_top5_by_period_pct_v9.csv"
OUT_PROF_PNG     = "table_prof_macro_top5_by_period_bw_v9.png"

EXPORT_PNG = True

PERIODS = [
    ("1848–1859", 1848, 1859),
    ("1860–1882", 1860, 1882),
    ("1883–1913", 1883, 1913),
    ("1914–1924", 1914, 1924),
    ("1925–1946", 1925, 1946),
]

TRUTHY_STRINGS = {"x", "1", "true", "yes", "si", "sì", "ok"}

CAT_2 = [
    ("cat_03", "cat_03\nDeputy"),
    ("cat_21", "cat_21\nTop taxpayer"),
]

COL_PROF_MACRO = "professione_macro"
TOP_N_PROF = 5

# Rendering controls
ROW_HEIGHT_SCALE_CAT = 2.10
ROW_HEIGHT_SCALE_PROF = 1.35
FONT_SIZE = 9

# --- NEW: margin / crop controls (tight PNGs)
PNG_DPI = 300
PAD_INCHES = 0.00          # removes outer padding
USE_CONSTRAINED_LAYOUT = False  # keep False; we control margins manually

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
    return mask_A, mask_B, mask_C

def pct(v, total):
    return (100.0 * v / total) if total else 0.0

def _crop_margins(fig, ax, table, pad=0.0):
    """
    Force axes to occupy full canvas, then crop tightly to the table bbox.
    This removes the 'big white margins' around the table PNG.
    """
    # Make the axes fill the whole figure (no outer margins)
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)

    # Need a draw to get correct bounding boxes
    fig.canvas.draw()

    # Tight bbox around table only
    renderer = fig.canvas.get_renderer()
    bbox = table.get_window_extent(renderer=renderer).expanded(1.01, 1.02)  # tiny safety expand
    # Convert display bbox -> figure inches
    bbox_inches = bbox.transformed(fig.dpi_scale_trans.inverted())

    return bbox_inches

def bw_table_png_matrix(df_pct: pd.DataFrame, out_png: str, row_height_scale: float):
    """
    Clean black-and-white table; no title inside image.
    Periods in columns; variables in rows.
    Exports a tightly-cropped PNG (minimal margins).
    """
    show_df = df_pct.copy()

    # Start with a sensible canvas; final PNG will be cropped to the table bbox anyway.
    fig_w = max(10.5, 1.8 + 1.05 * show_df.shape[1])
    fig_h = max(3.8, 1.4 + 0.55 * show_df.shape[0])

    fig, ax = plt.subplots(figsize=(fig_w, fig_h), constrained_layout=USE_CONSTRAINED_LAYOUT)
    ax.axis("off")

    cell_text = [[f"{v:.1f}" for v in row] for row in show_df.values]
    table = ax.table(
        cellText=cell_text,
        rowLabels=show_df.index.tolist(),
        colLabels=show_df.columns.tolist(),
        cellLoc="center",
        loc="center",
    )

    table.auto_set_font_size(False)
    table.set_fontsize(FONT_SIZE)
    table.scale(1.0, row_height_scale)

    # widen columns so headers stay inside cells
    try:
        table.auto_set_column_width(col=list(range(show_df.shape[1])))
    except Exception:
        pass

    # force taller cells (works even when scale is partly ignored)
    base_h = table[(1, 0)].get_height() if (1, 0) in table.get_celld() else None
    if base_h is not None:
        for (r, c), cell in table.get_celld().items():
            cell.set_height(base_h * row_height_scale)

    for (r, c), cell in table.get_celld().items():
        cell.set_facecolor("white")
        cell.set_edgecolor("black")
        cell.set_linewidth(0.85 if (r == 0 or c == -1) else 0.6)
        if r == 0:
            cell.set_text_props(weight="bold")
        if c == -1:
            cell.set_text_props(weight="bold")

    if EXPORT_PNG:
        bbox_inches = _crop_margins(fig, ax, table, pad=PAD_INCHES)
        fig.savefig(
            out_png,
            dpi=PNG_DPI,
            bbox_inches=bbox_inches,   # crop to table itself
            pad_inches=PAD_INCHES      # no extra padding
        )

    plt.show()

# =========================
# MAIN
# =========================
def main():
    if not Path(INPUT_FILE).is_file():
        raise FileNotFoundError(f"[STOP] Input file not found: {INPUT_FILE}")

    df = pd.read_excel(INPUT_FILE, sheet_name=INPUT_SHEET)

    col_nomina = pick_first_existing(
        df, ["data_nomina_dt", "data_nomina", "nomination_date_dt", "nomination_date", "nomination_dt"]
    )
    col_morte = pick_first_existing(
        df, ["effective_death_dt", "death_date_dt", "data_decesso_dt", "data_decesso", "death_date_raw", "death_dt"]
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

    apply_missing_death_rule(df)

    # union in-office across all periods
    union_mask = pd.Series(False, index=df.index)
    for _, y0, y1 in PERIODS:
        start = pd.Timestamp(year=y0, month=1, day=1)
        end = pd.Timestamp(year=y1, month=12, day=31)
        union_mask = union_mask | in_office_mask(df, start, end)

    block_all = df.loc[union_mask].copy()
    total_all = int(len(block_all))

    period_labels = [p[0] for p in PERIODS] + ["TOTAL (union)"]

    # -------------------------
    # TABLE 1: Art.33 (2 categories)
    # -------------------------
    missing_cats = [c for c, _ in CAT_2 if c not in df.columns]
    if missing_cats:
        raise KeyError(f"[STOP] Missing expected category columns: {missing_cats}")

    row_names_cat = [lab for _, lab in CAT_2] + ["Other", "Total"]
    cat_table = pd.DataFrame(index=row_names_cat, columns=period_labels, dtype=float)

    for label, y0, y1 in PERIODS:
        start = pd.Timestamp(year=y0, month=1, day=1)
        end = pd.Timestamp(year=y1, month=12, day=31)
        m = in_office_mask(df, start, end)
        block = df.loc[m].copy()
        total = int(len(block))

        cat_true = {lab: (block[col].apply(is_truthy_cell) if total else pd.Series([], dtype=bool))
                    for col, lab in CAT_2}

        other_n = int((~pd.DataFrame(cat_true).any(axis=1)).sum()) if total else 0

        for _, lab in CAT_2:
            cat_table.loc[lab, label] = pct(int(cat_true[lab].sum()) if total else 0, total)

        cat_table.loc["Other", label] = pct(other_n, total)
        cat_table.loc["Total", label] = 100.0 if total else 0.0

    if total_all:
        cat_true_all = {lab: block_all[col].apply(is_truthy_cell) for col, lab in CAT_2}
        other_all = int((~pd.DataFrame(cat_true_all).any(axis=1)).sum())

        for _, lab in CAT_2:
            cat_table.loc[lab, "TOTAL (union)"] = pct(int(cat_true_all[lab].sum()), total_all)

        cat_table.loc["Other", "TOTAL (union)"] = pct(other_all, total_all)
        cat_table.loc["Total", "TOTAL (union)"] = 100.0
    else:
        cat_table.loc[:, "TOTAL (union)"] = 0.0

    cat_table.to_csv(OUT_CAT_PCT_CSV, encoding="utf-8-sig")
    bw_table_png_matrix(cat_table, OUT_CAT_PNG, row_height_scale=ROW_HEIGHT_SCALE_CAT)

    # -------------------------
    # TABLE 2: professione_macro top-5 + Other
    # -------------------------
    if COL_PROF_MACRO not in df.columns:
        raise KeyError(f"[STOP] Missing column: {COL_PROF_MACRO}")

    prof_all = block_all[COL_PROF_MACRO].astype("string").fillna("").str.strip() if total_all else pd.Series([], dtype="string")
    prof_all = prof_all.replace("", "Missing")
    top_profs = prof_all.value_counts().head(TOP_N_PROF).index.tolist()

    row_names_prof = top_profs + ["Other", "Total"]
    prof_table = pd.DataFrame(index=row_names_prof, columns=period_labels, dtype=float)

    for label, y0, y1 in PERIODS:
        start = pd.Timestamp(year=y0, month=1, day=1)
        end = pd.Timestamp(year=y1, month=12, day=31)
        m = in_office_mask(df, start, end)

        s = df.loc[m, COL_PROF_MACRO].astype("string").fillna("").str.strip()
        s = s.replace("", "Missing")
        total = int(len(s))

        counts = {p: int((s == p).sum()) for p in top_profs}
        other_n = int(total - sum(counts.values()))

        for p in top_profs:
            prof_table.loc[p, label] = pct(counts[p], total)

        prof_table.loc["Other", label] = pct(other_n, total)
        prof_table.loc["Total", label] = 100.0 if total else 0.0

    if total_all:
        counts_all = {p: int((prof_all == p).sum()) for p in top_profs}
        other_all = int(total_all - sum(counts_all.values()))

        for p in top_profs:
            prof_table.loc[p, "TOTAL (union)"] = pct(counts_all[p], total_all)

        prof_table.loc["Other", "TOTAL (union)"] = pct(other_all, total_all)
        prof_table.loc["Total", "TOTAL (union)"] = 100.0
    else:
        prof_table.loc[:, "TOTAL (union)"] = 0.0

    prof_table.to_csv(OUT_PROF_PCT_CSV, encoding="utf-8-sig")
    bw_table_png_matrix(prof_table, OUT_PROF_PNG, row_height_scale=ROW_HEIGHT_SCALE_PROF)

    print("[OK] Saved:")
    print(f" - {OUT_CAT_PCT_CSV}")
    print(f" - {OUT_CAT_PNG}")
    print(f" - {OUT_PROF_PCT_CSV}")
    print(f" - {OUT_PROF_PNG}")

if __name__ == "__main__":
    main()