#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
SCOPO
-----
Questo script legge il file Excel del dataset dei senatori del Regno e produce
una nuova tabella Art. 33 aggregata in gruppi più leggibili per il paper.

La tabella mantiene gli stessi nomi di output già usati in precedenza, così può
sostituire direttamente la versione vecchia, ma viene salvata in una cartella
separata di lavoro.

LIBRERIE UTILIZZATE
-------------------
- pathlib
- numpy
- pandas
- matplotlib

INPUT
-----
- senatori_regno_1848_1943_master_dataset_final_v9.xlsx

OUTPUT
------
Cartella output:
- rifacimento_categorie_e_tabella/

File prodotti:
- rifacimento_categorie_e_tabella/table_art33_2cats_by_period_pct_v9.csv
- rifacimento_categorie_e_tabella/table_art33_2cats_by_period_bw_v9.png

LOGICA DELLA TABELLA
--------------------
La tabella riorganizza le categorie dell'Art. 33 in sei gruppi:

1. cat. 3   Deputies
2. cat. 21  Top taxpayers
3. cat. 14  General officers
4. cats. 4–13, 15–17  High state and judicial offices
5. cats. 18–19        Learned and educational elites
6. cats. 1–2, 20      Symbolic or exceptional routes

I valori sono percentuali "in office" per periodo.
"""

from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# =========================
# CONFIG / INPUT / OUTPUT
# =========================
INPUT_FILE = "senatori_regno_1848_1943_master_dataset_final_v9.xlsx"
INPUT_SHEET = 0

OUTPUT_DIR = Path("rifacimento_categorie_e_tabella")

# Manteniamo gli stessi identici nomi dei file vecchi
OUT_CAT_PCT_CSV = OUTPUT_DIR / "table_art33_2cats_by_period_pct_v9.csv"
OUT_CAT_PNG     = OUTPUT_DIR / "table_art33_2cats_by_period_bw_v9.png"

EXPORT_PNG = True

PERIODS = [
    ("1848–1859", 1848, 1859),
    ("1860–1882", 1860, 1882),
    ("1883–1913", 1883, 1913),
    ("1914–1924", 1914, 1924),
    ("1925–1946", 1925, 1946),
]

TRUTHY_STRINGS = {"x", "1", "true", "yes", "si", "sì", "ok"}

# =========================
# TABLE GROUPS
# =========================
ART33_GROUPS = [
    (
        "cat. 3\nDeputies",
        ["cat_03"]
    ),
    (
        "cat. 21\nTop taxpayers",
        ["cat_21"]
    ),
    (
        "cat. 14\nGeneral officers",
        ["cat_14"]
    ),
    (
        "cats. 4–13, 15–17\nHigh state and judicial offices",
        ["cat_04", "cat_05", "cat_06", "cat_07", "cat_08", "cat_09",
         "cat_10", "cat_11", "cat_12", "cat_13", "cat_15", "cat_16", "cat_17"]
    ),
    (
        "cats. 18–19\nLearned and educational elites",
        ["cat_18", "cat_19"]
    ),
    (
        "cats. 1–2, 20\nSymbolic or exceptional routes",
        ["cat_01", "cat_02", "cat_20"]
    ),
]

# =========================
# TABLE STYLE CONTROLS
# Regola questi valori all'inizio: sono quelli che ti servono davvero
# =========================

# Figura complessiva
FIG_WIDTH = 13.8
FIG_HEIGHT = 4.6

# Font
FONT_SIZE = 8.5
HEADER_FONT_SIZE = 8.5
ROW_LABEL_FONT_SIZE = 8.5

# Larghezze colonne
ROW_LABEL_WIDTH = 0.38     # colonna delle etichette a sinistra
DATA_COL_WIDTH = 0.105     # ogni colonna-periodo

# Altezze righe
HEADER_HEIGHT = 0.11       # altezza riga intestazione
DATA_ROW_HEIGHT = 0.095    # altezza righe dati

# Bordi
HEADER_LINEWIDTH = 0.85
BODY_LINEWIDTH = 0.60

# Esportazione PNG
PNG_DPI = 300
PAD_INCHES = 0.00
USE_CONSTRAINED_LAYOUT = False

# Piccola espansione del crop finale
BBOX_EXPAND_X = 1.01
BBOX_EXPAND_Y = 1.02

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
    df.loc[mask_C, "effective_death_dt"] = pd.Timestamp(year=1800, month=1, day=1)
    return mask_A, mask_B, mask_C

def pct(v, total):
    return (100.0 * v / total) if total else 0.0

def _crop_margins(fig, ax, table):
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    fig.canvas.draw()
    renderer = fig.canvas.get_renderer()
    bbox = table.get_window_extent(renderer=renderer).expanded(BBOX_EXPAND_X, BBOX_EXPAND_Y)
    bbox_inches = bbox.transformed(fig.dpi_scale_trans.inverted())
    return bbox_inches

def apply_table_dimensions(table, n_data_cols, n_rows_total):
    """
    Applica dimensioni controllate manualmente:
    - colonna etichette righe
    - colonne dati
    - altezza header
    - altezza righe dati
    """
    cells = table.get_celld()

    # Colonne dati: header e corpo
    for c in range(n_data_cols):
        if (0, c) in cells:
            cells[(0, c)].set_width(DATA_COL_WIDTH)
            cells[(0, c)].set_height(HEADER_HEIGHT)
        for r in range(1, n_rows_total + 1):
            if (r, c) in cells:
                cells[(r, c)].set_width(DATA_COL_WIDTH)
                cells[(r, c)].set_height(DATA_ROW_HEIGHT)

    # Colonna etichette riga
    for r in range(1, n_rows_total + 1):
        if (r, -1) in cells:
            cells[(r, -1)].set_width(ROW_LABEL_WIDTH)
            cells[(r, -1)].set_height(DATA_ROW_HEIGHT)

def style_table(table, n_data_cols, n_rows_total):
    cells = table.get_celld()

    for (r, c), cell in cells.items():
        cell.set_facecolor("white")
        cell.set_edgecolor("black")
        cell.set_linewidth(HEADER_LINEWIDTH if (r == 0 or c == -1) else BODY_LINEWIDTH)

        # Header
        if r == 0:
            cell.set_text_props(weight="bold", fontsize=HEADER_FONT_SIZE, ha="center", va="center")

        # Etichette riga
        elif c == -1:
            cell.set_text_props(weight="bold", fontsize=ROW_LABEL_FONT_SIZE, ha="left", va="center")

        # Corpo tabella
        else:
            cell.set_text_props(fontsize=FONT_SIZE, ha="center", va="center")

def bw_table_png_matrix(df_pct: pd.DataFrame, out_png: Path):
    show_df = df_pct.copy()

    fig, ax = plt.subplots(
        figsize=(FIG_WIDTH, FIG_HEIGHT),
        constrained_layout=USE_CONSTRAINED_LAYOUT
    )
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

    n_data_cols = show_df.shape[1]
    n_rows_total = show_df.shape[0]

    apply_table_dimensions(table, n_data_cols=n_data_cols, n_rows_total=n_rows_total)
    style_table(table, n_data_cols=n_data_cols, n_rows_total=n_rows_total)

    if EXPORT_PNG:
        bbox_inches = _crop_margins(fig, ax, table)
        fig.savefig(
            out_png,
            dpi=PNG_DPI,
            bbox_inches=bbox_inches,
            pad_inches=PAD_INCHES
        )

    plt.show()

def count_group_members(block: pd.DataFrame, cols: list[str]) -> int:
    """
    Conta quanti senatori nel blocco hanno almeno una categoria truthy
    tra quelle appartenenti al gruppo.
    """
    if block.empty:
        return 0

    available_cols = [c for c in cols if c in block.columns]
    if not available_cols:
        return 0

    truth_df = pd.DataFrame(
        {c: block[c].apply(is_truthy_cell) for c in available_cols},
        index=block.index
    )
    return int(truth_df.any(axis=1).sum())

# =========================
# MAIN
# =========================
def main():
    if not Path(INPUT_FILE).is_file():
        raise FileNotFoundError(f"[STOP] Input file not found: {INPUT_FILE}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(INPUT_FILE, sheet_name=INPUT_SHEET)

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

    apply_missing_death_rule(df)

    required_cat_cols = sorted({col for _, cols in ART33_GROUPS for col in cols})
    missing_cols = [c for c in required_cat_cols if c not in df.columns]
    if missing_cols:
        raise KeyError(f"[STOP] Missing expected category columns: {missing_cols}")

    union_mask = pd.Series(False, index=df.index)
    for _, y0, y1 in PERIODS:
        start = pd.Timestamp(year=y0, month=1, day=1)
        end = pd.Timestamp(year=y1, month=12, day=31)
        union_mask = union_mask | in_office_mask(df, start, end)

    block_all = df.loc[union_mask].copy()
    total_all = int(len(block_all))

    period_labels = [p[0] for p in PERIODS] + ["TOTAL (union)"]
    row_names = [label for label, _ in ART33_GROUPS] + ["Total"]

    cat_table = pd.DataFrame(index=row_names, columns=period_labels, dtype=float)

    for label, y0, y1 in PERIODS:
        start = pd.Timestamp(year=y0, month=1, day=1)
        end = pd.Timestamp(year=y1, month=12, day=31)
        m = in_office_mask(df, start, end)
        block = df.loc[m].copy()
        total = int(len(block))

        for row_label, cols in ART33_GROUPS:
            n = count_group_members(block, cols)
            cat_table.loc[row_label, label] = pct(n, total)

        cat_table.loc["Total", label] = 100.0 if total else 0.0

    for row_label, cols in ART33_GROUPS:
        n_all = count_group_members(block_all, cols)
        cat_table.loc[row_label, "TOTAL (union)"] = pct(n_all, total_all)

    cat_table.loc["Total", "TOTAL (union)"] = 100.0 if total_all else 0.0

    cat_table.to_csv(OUT_CAT_PCT_CSV, encoding="utf-8-sig")
    bw_table_png_matrix(cat_table, OUT_CAT_PNG)

    print("[OK] Output directory:")
    print(f" - {OUTPUT_DIR}")
    print("\n[OK] Saved:")
    print(f" - {OUT_CAT_PCT_CSV}")
    print(f" - {OUT_CAT_PNG}")

if __name__ == "__main__":
    main()