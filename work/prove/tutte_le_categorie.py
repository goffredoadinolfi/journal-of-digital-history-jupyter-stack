#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
INPUT
-----
- senatori_regno_1848_1943_master_dataset_final_v9.xlsx

OUTPUT
------
# Overall (1848–1946) table for ALL 21 Art.33 categories (in-office union)
- art33_all21_overall_counts_pct_v9.csv
- art33_all21_overall_bw_v9.png

# Category selection at threshold (>= 5% of total)
- art33_selected_ge5pct_v9.csv   (cat_code, count, pct)

What the script does
--------------------
- Applies the same "in office" rule and missing-death rule (A/B/C).
- Builds ONE overall table (1848–1946) with counts and percentages for cat_01..cat_21
  among senators who are "in office" in at least one of the five periods (union mask).
- Exports CSV + a black-and-white PNG table.
- Computes which categories are >= 5% of the total and exports them.
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

OUT_OVERALL_CSV = "art33_all21_overall_counts_pct_v9.csv"
OUT_OVERALL_PNG = "art33_all21_overall_bw_v9.png"
OUT_SELECTED_CSV = "art33_selected_ge5pct_v9.csv"

EXPORT_PNG = True
THRESHOLD_PCT = 5.0

PERIODS = [
    ("1848–1859", 1848, 1859),
    ("1860–1882", 1860, 1882),
    ("1883–1913", 1883, 1913),
    ("1914–1924", 1914, 1924),
    ("1925–1946", 1925, 1946),
]

TRUTHY_STRINGS = {"x", "1", "true", "yes", "si", "sì", "ok"}

CAT_COLS = [f"cat_{i:02d}" for i in range(1, 22)]  # cat_01 ... cat_21

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

def bw_table_png_overall(df_show: pd.DataFrame, title: str, out_png: str):
    """
    df_show expected columns: ["Category","Count","Percent"]
    """
    fig_w = 10.5
    fig_h = max(3.2, 0.55 * (len(df_show) + 2))

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis("off")

    cell_text = []
    for _, r in df_show.iterrows():
        cell_text.append([str(r["Category"]), f'{int(r["Count"])}', f'{float(r["Percent"]):.1f}'])

    col_labels = ["Category", "Count", "Percent"]

    table = ax.table(
        cellText=cell_text,
        colLabels=col_labels,
        cellLoc="center",
        loc="center"
    )

    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1.0, 1.25)

    # Make columns wide enough so headers fit inside their cells
    # (prevents header text spilling out)
    try:
        table.auto_set_column_width(col=list(range(len(col_labels))))
    except Exception:
        pass

    for (r, c), cell in table.get_celld().items():
        cell.set_facecolor("white")
        cell.set_edgecolor("black")
        cell.set_linewidth(0.8 if r == 0 else 0.6)
        if r == 0:
            cell.set_text_props(weight="bold")

    ax.set_title(title, fontsize=10, pad=12)
    plt.tight_layout()

    if EXPORT_PNG:
        fig.savefig(out_png, dpi=300, bbox_inches="tight")
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

    mask_A, mask_B, mask_C = apply_missing_death_rule(df)

    missing_cols = [c for c in CAT_COLS if c not in df.columns]
    if missing_cols:
        raise KeyError(f"[STOP] Missing Art.33 category columns: {missing_cols}")

    # Union "in office" mask across the five periods
    union_mask = pd.Series(False, index=df.index)
    for _, y0, y1 in PERIODS:
        start = pd.Timestamp(year=y0, month=1, day=1)
        end = pd.Timestamp(year=y1, month=12, day=31)
        union_mask = union_mask | in_office_mask(df, start, end)

    block = df.loc[union_mask, CAT_COLS].copy()
    total_in_union = int(len(block))
    if total_in_union == 0:
        raise SystemExit("[STOP] No in-office rows found across the requested periods.")

    counts = {}
    for c in CAT_COLS:
        counts[c] = int(block[c].apply(is_truthy_cell).sum())

    out = pd.DataFrame({
        "Category": list(counts.keys()),
        "Count": list(counts.values())
    })
    out["Percent"] = (out["Count"] / total_in_union * 100.0).round(2)
    out = out.sort_values(["Percent", "Count"], ascending=[False, False]).reset_index(drop=True)

    # Add TOTAL row (Count = total senators in union, Percent = 100)
    total_row = pd.DataFrame([{
        "Category": "TOTAL (senators in office, union)",
        "Count": total_in_union,
        "Percent": 100.0
    }])
    out_with_total = pd.concat([out, total_row], ignore_index=True)

    out_with_total.to_csv(OUT_OVERALL_CSV, index=False, encoding="utf-8-sig")

    bw_table_png_overall(
        df_show=out_with_total,
        title="Art.33 categories among senators in office (1848–1946 union): counts and percentages",
        out_png=OUT_OVERALL_PNG
    )

    # Select categories >= threshold
    selected = out[out["Percent"] >= THRESHOLD_PCT].copy()
    selected = selected.rename(columns={"Category": "cat_code", "Count": "n", "Percent": "pct"})
    selected.to_csv(OUT_SELECTED_CSV, index=False, encoding="utf-8-sig")

    print("[OK] Input:", INPUT_FILE)
    print("[OK] Rows read:", int(len(df)))
    print("[OK] Missing-death rule -> A:", int(mask_A.sum()), "| B:", int(mask_B.sum()), "| C:", int(mask_C.sum()))
    print("[OK] Senators in office (union across periods):", total_in_union)
    print("[OK] Saved overall CSV:", OUT_OVERALL_CSV)
    if EXPORT_PNG:
        print("[OK] Saved overall PNG:", OUT_OVERALL_PNG)
    print("[OK] Threshold:", THRESHOLD_PCT, "% | Categories kept:", int(len(selected)))
    print("[OK] Saved selected CSV:", OUT_SELECTED_CSV)
    print("[OK] Selected categories:", selected["cat_code"].tolist())

if __name__ == "__main__":
    main()