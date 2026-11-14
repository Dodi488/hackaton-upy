#!/usr/bin/env python3
"""
merge_xlsx.py (v2) - Stack many .xlsx files that share the same table layout
into one .xlsx (and optionally a .csv).

Every row keeps its original values. Two columns are added at the front:
    Source_File : the workbook the row came from
    Hoja        : the sheet name inside that workbook. In the RENAMECA files
                  this is the period ("2012-2022"), which is NOT stored
                  anywhere else, so it must not be lost.

Nothing is cleaned, renamed or converted - this is a faithful stack. (In the
original CONAGUA files the numbers are stored as text, e.g. '355.68', and they
stay that way here. profile_agua.py converts them.)

After merging, the script re-counts the data rows of every input file in an
independent way and prints OK / MISMATCH, so you do not have to trust it.

Usage (from a terminal):

    python merge_xlsx.py INPUT_FOLDER OUTPUT.xlsx

    python merge_xlsx.py "C:\\Users\\diego\\Downloads\\datos" merged_calidad_agua.xlsx --csv

Options:
    --pattern "*.xlsx"      which files to merge
                            (default: Calidad_de_agua_Estado_*.xlsx)
    --all-sheets            read every sheet of every workbook, not only the
                            first one
    --csv                   also write a UTF-8 .csv next to the output

Files that do not match the pattern (e.g. master_*.xlsx, matriz_*.xlsx, which
are already combined/reshaped copies of the same data) are skipped and listed.

Requires: pandas, openpyxl   (pip install pandas openpyxl)
"""

import argparse
import fnmatch
import sys
from pathlib import Path

import openpyxl
import pandas as pd

HEADER_KEY = "CLAVE"      # a cell with this text marks the header row
MAX_HEADER_SEARCH = 20    # how many top rows to scan for the header row
EXCEL_MAX_ROWS = 1_048_575


def _is_blank(c) -> bool:
    return c is None or str(c).strip() == ""


def find_header_row(raw: pd.DataFrame) -> int:
    """Index of the row that holds the column names (first row containing
    the text CLAVE within the top rows; otherwise row 0)."""
    for i in range(min(MAX_HEADER_SEARCH, len(raw))):
        cells = raw.iloc[i].astype(str).str.strip().str.upper()
        if (cells == HEADER_KEY).any():
            return i
    return 0


def read_sheet(xl: pd.ExcelFile, sheet: str) -> pd.DataFrame:
    # header=None + dtype=object: keep every cell exactly as stored, so text
    # such as "<0.01" is not forced into numbers and numbers are not mangled.
    raw = xl.parse(sheet, header=None, dtype=object)
    raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    if raw.empty:
        return pd.DataFrame()
    h = find_header_row(raw)
    names = [str(v).strip() if pd.notna(v) else f"col_{j}"
             for j, v in enumerate(raw.iloc[h])]
    df = raw.iloc[h + 1:].copy()
    df.columns = names
    return df.reset_index(drop=True)


def count_data_rows(path: Path, sheet: str) -> int:
    """Independent row count straight from openpyxl (does not use the pandas
    parsing above): non-empty rows below the header row."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    nonempty = [r for r in wb[sheet].iter_rows(values_only=True)
                if any(not _is_blank(c) for c in r)]
    wb.close()
    hdr = 0
    for i, r in enumerate(nonempty[:MAX_HEADER_SEARCH]):
        if any(not _is_blank(c) and str(c).strip().upper() == HEADER_KEY for c in r):
            hdr = i
            break
    return max(len(nonempty) - hdr - 1, 0)


def main() -> int:
    ap = argparse.ArgumentParser(description="Stack many xlsx files into one.")
    ap.add_argument("input_folder", type=Path)
    ap.add_argument("output", type=Path, help="output .xlsx path")
    ap.add_argument("--pattern", default="Calidad_de_agua_Estado_*.xlsx")
    ap.add_argument("--all-sheets", action="store_true")
    ap.add_argument("--csv", action="store_true")
    args = ap.parse_args()

    folder = args.input_folder
    if not folder.is_dir():
        print(f"ERROR: folder not found: {folder}")
        return 1

    out_resolved = args.output.resolve()
    all_files = sorted(
        p for p in folder.glob("*.xlsx")
        if not p.name.startswith(("~$", ".~lock"))   # Excel / LibreOffice locks
        and p.resolve() != out_resolved)              # never re-read the output
    chosen = [p for p in all_files if fnmatch.fnmatch(p.name, args.pattern)]
    skipped = [p for p in all_files if p not in chosen]
    if not chosen:
        print(f"ERROR: no files in {folder} match '{args.pattern}'")
        return 1

    print(f"Merging {len(chosen)} file(s) from {folder}\n")
    frames, failed, mismatches = [], [], []
    for p in chosen:
        try:
            xl = pd.ExcelFile(p)
            sheets = xl.sheet_names if args.all_sheets else xl.sheet_names[:1]
            if not args.all_sheets and len(xl.sheet_names) > 1:
                print(f"    note: {p.name} has {len(xl.sheet_names)} sheets; "
                      f"using only '{sheets[0]}' (use --all-sheets for all)")
            for sh in sheets:
                df = read_sheet(xl, sh)
                if df.empty:
                    print(f"  [EMPTY] {p.name} / {sh}")
                    continue
                expected = count_data_rows(p, sh)
                if expected != len(df):
                    mismatches.append((p.name, sh, len(df), expected))
                df.insert(0, "Hoja", sh)
                df.insert(0, "Source_File", p.name)
                frames.append(df)
                print(f"  [ok] {p.name} [{sh}]: {len(df)} rows x "
                      f"{df.shape[1] - 2} cols")
        except Exception as e:  # keep going, report at the end
            failed.append((p.name, str(e)))
            print(f"  [FAIL] {p.name}: {e}")

    if not frames:
        print("ERROR: nothing could be read.")
        return 1

    ref_cols = list(frames[0].columns)
    for df in frames[1:]:
        cols = list(df.columns)
        if cols != ref_cols:
            extra = [c for c in cols if c not in ref_cols]
            missing = [c for c in ref_cols if c not in cols]
            print(f"\n  WARNING column layout differs in {df['Source_File'].iloc[0]}"
                  f"\n    extra: {extra}\n    missing: {missing}")

    merged = pd.concat(frames, ignore_index=True, sort=False)

    print(f"\nTotal: {len(merged)} rows x {merged.shape[1]} columns "
          f"from {merged['Source_File'].nunique()} file(s)")
    print("Sheet names found (Hoja):",
          ", ".join(f"{k} ({v} files)" for k, v in
                    merged.groupby("Hoja")["Source_File"].nunique().items()))

    # --- self-check -------------------------------------------------------
    if mismatches:
        print("\nCHECK: MISMATCH between rows read and rows counted independently:")
        for n, sh, got, exp in mismatches:
            print(f"  - {n} [{sh}]: merged {got}, counted {exp}")
    else:
        print(f"CHECK: OK - every file's row count matches an independent "
              f"recount ({len(merged)} rows).")

    if skipped:
        print(f"\nSkipped {len(skipped)} file(s) that do not match "
              f"'{args.pattern}':")
        for p in skipped:
            print(f"  - {p.name}")
    if failed:
        print(f"\n{len(failed)} file(s) FAILED - fix these and re-run:")
        for n, e in failed:
            print(f"  - {n}: {e}")

    args.output.parent.mkdir(parents=True, exist_ok=True)
    if len(merged) > EXCEL_MAX_ROWS:
        print("\nToo many rows for one Excel sheet - writing CSV only.")
        args.csv = True
    else:
        with pd.ExcelWriter(args.output, engine="openpyxl") as xw:
            merged.to_excel(xw, sheet_name="datos", index=False)
            ws = xw.sheets["datos"]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
        print(f"\nSaved: {args.output}")

    if args.csv:
        csv_path = args.output.with_suffix(".csv")
        merged.to_csv(csv_path, index=False, encoding="utf-8-sig")
        print(f"Saved: {csv_path}")

    return 0 if not (failed or mismatches) else 2


if __name__ == "__main__":
    sys.exit(main())
