#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
SCOPO
-----
Estrarre il testo da un file Jupyter Notebook (.ipynb) escludendo le celle di codice.

LIBRERIE UTILIZZATE
-------------------
- nbformat
- pathlib

INPUT
-----
- article.ipynb

OUTPUT
------
- article_solo_testo.txt
"""

from pathlib import Path
import nbformat
from docx import Document

# =========================
# INPUT / OUTPUT
# =========================
INPUT_FILE = "article.ipynb"
OUTPUT_FILE = "article_solo_testo.docx"

# =========================
# MAIN
# =========================
def main():
    input_path = Path(INPUT_FILE)
    output_path = Path(OUTPUT_FILE)

    if not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {INPUT_FILE}")

    with input_path.open("r", encoding="utf-8") as f:
        nb = nbformat.read(f, as_version=4)

    parts = []

    for cell in nb.cells:
        if cell.cell_type in {"markdown", "raw"}:
            text = cell.source.strip()
            if text:
                parts.append(text)

    final_text = "\n\n" + ("\n\n" + ("-" * 80) + "\n\n").join(parts) if parts else ""

    with output_path.open("w", encoding="utf-8") as f:
        f.write(final_text.strip())

    print(f"[OK] Extracted text saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()