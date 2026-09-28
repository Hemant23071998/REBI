"""
Excel pipeline: combine -> filter by chosen dates -> save .xlsx -> sort -> dedupe -> save tab-delimited .txt

Usage (interactive - just run it and answer the prompts):
    py excel_pipeline.py

Usage (command line) - dates are MM-DD-YYYY (or YYYY-MM-DD), single days and/or "A to B" ranges:
    py excel_pipeline.py "C:\\Data\\Input" "10-05-2026, 10-06-2026, 10-07-2026, 10-15-2026, 10-16-2026"
    py excel_pipeline.py "C:\\Data\\Input" "10-05-2026 to 10-07-2026, 10-15-2026 to 10-16-2026"
    py excel_pipeline.py "C:\\Data\\Input" "05/10/2026, 15/10/2026" --date-col "Order Date" --dayfirst

Outputs (written to <input folder>\\Output unless --out is given):
    Filtered_<first>_to_<last>.xlsx  -> steps 1-5 (combined + date-filtered, with header)
    Final_<first>_to_<last>.txt      -> steps 6-9 (sorted, deduped on Brand Name, no header, tab-delimited)
"""

import argparse
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
SORT_COLUMNS = ["Phone Number", "Website", "Brand Name"]
DEDUPE_COLUMN = "Brand Name"


def parse_user_date(text, dayfirst):
    text = text.strip()
    if dayfirst:
        formats = ["%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y"]
    else:
        formats = ["%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%m.%d.%Y"]
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    sys.exit(f"ERROR: could not understand date '{text}'. Use MM-DD-YYYY, e.g. 10-05-2026.")


def parse_date_spec(text, dayfirst):
    """'2026-10-05, 2026-10-07 to 2026-10-09' -> {Oct 5, Oct 7, Oct 8, Oct 9}"""
    text = re.sub(r"\s+to\s+", "..", text.strip(), flags=re.IGNORECASE)
    days = set()
    for item in filter(None, re.split(r"[,;\s]+", text)):
        if ".." in item:
            a, b = (parse_user_date(x, dayfirst) for x in item.split("..", 1))
            if b < a:
                sys.exit(f"ERROR: range '{item}' ends before it starts")
            days.update(a + timedelta(days=i) for i in range((b - a).days + 1))
        else:
            days.add(parse_user_date(item, dayfirst))
    if not days:
        sys.exit("ERROR: no dates given")
    return days


def find_column(df, wanted):
    """Match a column name ignoring case and extra spaces."""
    norm = lambda s: " ".join(str(s).split()).lower()
    for col in df.columns:
        if norm(col) == norm(wanted):
            return col
    return None


def to_text(v):
    """Keep real dates as dates; turn everything else into text (9876543210.0 -> '9876543210')."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return pd.NA
    if isinstance(v, (datetime, pd.Timestamp)):
        return v
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def to_date(v, dayfirst):
    """Real Excel dates pass through; text dates are parsed (DD/MM if dayfirst)."""
    if isinstance(v, (datetime, pd.Timestamp)):
        return pd.Timestamp(v)
    if v is pd.NA or v is None:
        return pd.NaT
    return pd.to_datetime(str(v).strip(), errors="coerce", dayfirst=dayfirst)


# ---------- Steps 1-2: read every Excel file and combine ----------
def combine_files(folder, all_sheets):
    files = sorted(
        p for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS and not p.name.startswith("~$")
    )
    if not files:
        sys.exit(f"ERROR: no Excel files found in {folder}")

    frames = []
    for f in files:
        sheets = pd.read_excel(f, sheet_name=None if all_sheets else 0, dtype=object)
        if not isinstance(sheets, dict):
            sheets = {"(first sheet)": sheets}
        for sheet_name, df in sheets.items():
            df.columns = [" ".join(str(c).split()) for c in df.columns]  # tidy header spacing
            df = df.dropna(how="all").map(to_text)
            print(f"  read {f.name} [{sheet_name}]: {len(df)} rows")
            frames.append(df)

    combined = pd.concat(frames, ignore_index=True, sort=False)
    print(f"Combined {len(files)} file(s): {len(combined)} rows, {len(combined.columns)} columns")
    return combined


# ---------- Step 4: keep only rows whose date is one of the chosen days ----------
def filter_by_date(df, date_col, days, dayfirst):
    col = find_column(df, date_col) if date_col else None
    if date_col and col is None:
        sys.exit(f"ERROR: date column '{date_col}' not found. Columns are: {list(df.columns)}")
    if col is None:  # auto-detect: first column with "date" in its name
        candidates = [c for c in df.columns if "date" in str(c).lower()]
        if not candidates:
            sys.exit(f"ERROR: no date column found; pass --date-col. Columns are: {list(df.columns)}")
        col = candidates[0]
        print(f"Using date column: '{col}'")

    parsed = pd.to_datetime(df[col].map(lambda v: to_date(v, dayfirst)))
    bad = parsed.isna() & df[col].notna()
    if bad.any():
        print(f"  WARNING: {bad.sum()} row(s) have an unreadable date in '{col}' and were excluded")

    # compare calendar day only, so 2026-10-05 15:30 matches 2026-10-05
    mask = parsed.dt.date.isin(days)
    out = df.loc[mask].copy()
    out[col] = parsed[mask]
    print(f"Date filter ({len(days)} day(s), {min(days):%m-%d-%Y} .. {max(days):%m-%d-%Y}): {len(out)} rows kept")
    missing = sorted(days - set(out[col].dt.date))
    if missing:
        print(f"  note: no rows found for {len(missing)} chosen day(s): {', '.join(f'{m:%m-%d-%Y}' for m in missing)}")
    return out, col


# ---------- Steps 6-7: sort, then remove duplicates on Brand Name ----------
def sort_and_dedupe(df):
    cols = []
    for wanted in SORT_COLUMNS:
        col = find_column(df, wanted)
        if col is None:
            sys.exit(f"ERROR: column '{wanted}' not found. Columns are: {list(df.columns)}")
        cols.append(col)

    df = df.copy()
    for c in cols:  # blanks sort last, like Excel
        df[c] = df[c].where(df[c].astype(str).str.strip() != "", pd.NA)

    # case-insensitive A-Z sort (Excel behaviour); stable so ties keep file order
    df = df.sort_values(
        by=cols,
        key=lambda s: s.astype("string").str.strip().str.lower(),
        na_position="last",
        kind="mergesort",
    )

    brand = cols[SORT_COLUMNS.index(DEDUPE_COLUMN)]
    brand_key = df[brand].astype("string").str.strip().str.lower()
    before = len(df)
    df = df.loc[~brand_key.duplicated(keep="first")]
    print(f"Sorted by {cols}; removed {before - len(df)} duplicate '{brand}' row(s): {len(df)} rows left")
    return df


def main():
    ap = argparse.ArgumentParser(description="Combine, filter, sort and dedupe Excel files.")
    ap.add_argument("folder", nargs="?", help="folder containing the Excel files")
    ap.add_argument("dates", nargs="*",
                    help='days and/or ranges, e.g. "10-05-2026, 10-06-2026, 10-15-2026 to 10-16-2026"')
    ap.add_argument("--date-col", help="name of the date column (default: auto-detect)")
    ap.add_argument("--dayfirst", action="store_true",
                    help="treat ambiguous dates like 03/04/2026 as 3 April (DD/MM)")
    ap.add_argument("--all-sheets", action="store_true", help="read every sheet, not just the first")
    ap.add_argument("--out", help="output folder (default: <folder>\\Output)")
    ap.add_argument("--txt-date-format", default="%m-%d-%Y",
                    help="date format in the .txt file (default %%m-%%d-%%Y)")
    a = ap.parse_args()

    folder = Path((a.folder or input("Input folder path: ")).strip().strip('"'))
    if not folder.is_dir():
        sys.exit(f"ERROR: folder not found: {folder}")
    spec = ",".join(a.dates) if a.dates else input(
        "Dates (MM-DD-YYYY, comma-separated; ranges as 'A to B'): ")
    days = parse_date_spec(spec, a.dayfirst)

    out_dir = Path(a.out) if a.out else folder / "Output"
    out_dir.mkdir(parents=True, exist_ok=True)
    tag = f"{min(days):%Y%m%d}_to_{max(days):%Y%m%d}"

    combined = combine_files(folder, a.all_sheets)
    filtered, date_col = filter_by_date(combined, a.date_col, days, a.dayfirst)

    # Step 5: save filtered data (with header) as Excel
    xlsx_path = out_dir / f"Filtered_{tag}.xlsx"
    with pd.ExcelWriter(xlsx_path) as xw:
        filtered.to_excel(xw, index=False)
        date_idx = list(filtered.columns).index(date_col) + 1
        for (cell,) in xw.sheets["Sheet1"].iter_rows(min_row=2, min_col=date_idx, max_col=date_idx):
            cell.number_format = "mm-dd-yyyy"
    print(f"Saved: {xlsx_path}")

    # Steps 6-9
    final = sort_and_dedupe(filtered)
    final[date_col] = final[date_col].dt.strftime(a.txt_date_format)
    txt_path = out_dir / f"Final_{tag}.txt"
    final.to_csv(txt_path, sep="\t", index=False, header=False, encoding="utf-8", lineterminator="\r\n")
    print(f"Saved: {txt_path}")
    print("Done.")


if __name__ == "__main__":
    main()
