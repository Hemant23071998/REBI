"""
Excel pipeline - saves two Excel files:
    1. Combined_<first>_to_<last>.xlsx  all files joined, chosen dates only, otherwise as-is  (COMBINED_FOLDER)
    2. Final_<first>_to_<last>.xlsx     chosen dates only, sorted (Phone, Website, Brand), duplicate brands
                                      removed, DROP_COLUMNS deleted, "-" removed from phone numbers,
                                      "^" removed from every column                      (OUTPUT_FOLDER)

Run it (asks for anything not given):
    py excel_pipeline.py
    py excel_pipeline.py "C:\\Data\\Input" "10-05-2026, 10-06-2026, 10-15-2026 to 10-16-2026"

All settings are in the SETTINGS block below. Columns are found by their HEADER TEXT (any position),
ignoring upper/lower case and extra spaces.
"""

# ==============================================================================
#  SETTINGS - EDIT HERE
# ==============================================================================

# ---- Folders (use r"..." so backslashes work) ----
INPUT_FOLDER = r""          # folder with the Excel files; "" = ask each time
COMBINED_FOLDER = r""       # where Combined_*.xlsx is saved;   "" = same as OUTPUT_FOLDER
OUTPUT_FOLDER = r""         # where Final_*.xlsx is saved;      "" = <input folder>\Output

# ---- Column headers (type them exactly as they appear in row 1 of your files) ----
PHONE_COLUMN = "Phone Number"
WEBSITE_COLUMN = "Website"
BRAND_COLUMN = "Brand Name"   # also the column used to remove duplicates
DATE_COLUMN = ""              # "" = first column with "date" in its header

# ---- Columns to DELETE from the Final file (exact header names) ----
DROP_COLUMNS = [
    # "Email",
    # "Address",
]

# ---- Characters to remove from the Final file ----
PHONE_REMOVE_CHARS = ["-"]        # removed from the phone column only
REMOVE_CHARS_EVERYWHERE = ["^"]   # removed from every column

# ---- Dates ----
DAY_FIRST = False                 # False = MM-DD-YYYY, True = DD-MM-YYYY
DATE_DISPLAY_FORMAT = "mm-dd-yyyy"  # how dates show in the Final file

# ==============================================================================

import argparse
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}


def parse_user_date(text, dayfirst):
    text = text.strip()
    if dayfirst:
        formats = ["%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y"]
    else:
        formats = ["%Y-%m-%d", "%m-%d-%Y", "%m/%d/%Y", "%m.%d.%Y"]
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    sys.exit(f"ERROR: could not understand date '{text}'. Use MM-DD-YYYY, e.g. 10-05-2026.")


def parse_date_spec(text, dayfirst):
    """'10-05-2026, 10-07-2026 to 10-09-2026' -> {Oct 5, Oct 7, Oct 8, Oct 9}"""
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
    """Real Excel dates pass through; text dates are parsed (MM-DD unless dayfirst)."""
    if isinstance(v, (datetime, pd.Timestamp)):
        return pd.Timestamp(v)
    if v is pd.NA or v is None:
        return pd.NaT
    return pd.to_datetime(str(v).strip(), errors="coerce", dayfirst=dayfirst)


def save_xlsx(df, path, date_col=None):
    with pd.ExcelWriter(path) as xw:
        df.to_excel(xw, index=False)
        if date_col is not None:
            idx = list(df.columns).index(date_col) + 1
            for (cell,) in xw.sheets["Sheet1"].iter_rows(min_row=2, min_col=idx, max_col=idx):
                cell.number_format = DATE_DISPLAY_FORMAT
    print(f"Saved: {path}")


# ---------- Steps 1-2: read every Excel file and combine ----------
def combine_files(folder, all_sheets):
    files = sorted(
        p for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS and not p.name.startswith("~$")
    )
    if not files:
        sys.exit(f"ERROR: no Excel files found in {folder}")

    raw_frames, frames, row_id = [], [], 0
    for f in files:
        sheets = pd.read_excel(f, sheet_name=None if all_sheets else 0, dtype=object)
        if not isinstance(sheets, dict):
            sheets = {"(first sheet)": sheets}
        for sheet_name, raw in sheets.items():
            # unique row numbers link each cleaned row back to its untouched original
            raw.index = range(row_id, row_id + len(raw))
            row_id += len(raw)
            raw_frames.append(raw)                                       # untouched copy
            df = raw.copy()
            df.columns = [" ".join(str(c).split()) for c in df.columns]  # tidy header spacing
            df = df.dropna(how="all").map(to_text)
            print(f"  read {f.name} [{sheet_name}]: {len(df)} rows")
            frames.append(df)

    raw_combined = pd.concat(raw_frames, sort=False)
    combined = pd.concat(frames, sort=False)
    print(f"Combined {len(files)} file(s): {len(combined)} rows, {len(combined.columns)} columns")
    return raw_combined, combined


# ---------- Step 4: keep only rows whose date is one of the chosen days ----------
def filter_by_date(df, date_col, days, dayfirst):
    col = find_column(df, date_col) if date_col else None
    if date_col and col is None:
        sys.exit(f"ERROR: date column '{date_col}' not found. Columns are: {list(df.columns)}")
    if col is None:  # auto-detect: first column with "date" in its name
        candidates = [c for c in df.columns if "date" in str(c).lower()]
        if not candidates:
            sys.exit(f"ERROR: no date column found; set DATE_COLUMN. Columns are: {list(df.columns)}")
        col = candidates[0]
        print(f"Using date column: '{col}'")

    parsed = pd.to_datetime(df[col].map(lambda v: to_date(v, dayfirst)))
    bad = parsed.isna() & df[col].notna()
    if bad.any():
        print(f"  WARNING: {bad.sum()} row(s) have an unreadable date in '{col}' and were excluded")

    # compare calendar day only, so 10-05-2026 15:30 matches 10-05-2026
    mask = parsed.dt.date.isin(days)
    out = df.loc[mask].copy()
    out[col] = parsed[mask]
    print(f"Date filter ({len(days)} day(s), {min(days):%m-%d-%Y} .. {max(days):%m-%d-%Y}): {len(out)} rows kept")
    missing = sorted(days - set(out[col].dt.date))
    if missing:
        print(f"  note: no rows found for {len(missing)} chosen day(s): {', '.join(f'{m:%m-%d-%Y}' for m in missing)}")
    return out, col


# ---------- Steps 6-7: sort, then remove duplicates on Brand Name ----------
def sort_and_dedupe(df, phone, website, brand):
    cols = [find_column(df, c) for c in (phone, website, brand)]

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

    brand_col = cols[2]
    brand_key = df[brand_col].astype("string").str.strip().str.lower()
    before = len(df)
    df = df.loc[~brand_key.duplicated(keep="first")]
    print(f"Sorted by {cols}; removed {before - len(df)} duplicate '{brand_col}' row(s): {len(df)} rows left")
    return df


# ---------- Final file: drop columns, clean characters ----------
def drop_and_clean(df, date_col, phone):
    df = df.copy()
    drop = []
    for name in DROP_COLUMNS:
        col = find_column(df, name)
        if col is None:
            print(f"  WARNING: column to delete '{name}' not found - skipped")
        else:
            drop.append(col)
    df = df.drop(columns=drop)
    if drop:
        print(f"Deleted column(s): {drop}")

    phone_col = find_column(df, phone)
    if phone_col is not None:
        for ch in PHONE_REMOVE_CHARS:
            df[phone_col] = df[phone_col].str.replace(ch, "", regex=False)

    for col in df.columns:
        if col == date_col:  # real dates, nothing to clean
            continue
        for ch in REMOVE_CHARS_EVERYWHERE:
            df[col] = df[col].astype("string").str.replace(ch, "", regex=False)
    print(f"Removed {PHONE_REMOVE_CHARS} from phone numbers and {REMOVE_CHARS_EVERYWHERE} from all columns")
    return df


def main():
    ap = argparse.ArgumentParser(description="Combine, filter, sort and dedupe Excel files.")
    ap.add_argument("folder", nargs="?", default=INPUT_FOLDER or None, help="folder containing the Excel files")
    ap.add_argument("dates", nargs="*",
                    help='days and/or ranges, e.g. "10-05-2026, 10-06-2026, 10-15-2026 to 10-16-2026"')
    ap.add_argument("--date-col", default=DATE_COLUMN, help="date column header (default: auto-detect)")
    ap.add_argument("--phone-col", default=PHONE_COLUMN, help=f'default "{PHONE_COLUMN}"')
    ap.add_argument("--website-col", default=WEBSITE_COLUMN, help=f'default "{WEBSITE_COLUMN}"')
    ap.add_argument("--brand-col", default=BRAND_COLUMN, help=f'default "{BRAND_COLUMN}"')
    ap.add_argument("--dayfirst", action="store_true", default=DAY_FIRST,
                    help="dates are DD-MM-YYYY instead of MM-DD-YYYY")
    ap.add_argument("--all-sheets", action="store_true", help="read every sheet, not just the first")
    ap.add_argument("--out", default=OUTPUT_FOLDER or None, help="output folder (default: <folder>\\Output)")
    ap.add_argument("--combined-out", default=COMBINED_FOLDER or None,
                    help="folder for the combined file (default: same as --out)")
    a = ap.parse_args()

    folder = Path((a.folder or input("Input folder path: ")).strip().strip('"'))
    if not folder.is_dir():
        sys.exit(f"ERROR: folder not found: {folder}")
    spec = ",".join(a.dates) if a.dates else input(
        "Dates (MM-DD-YYYY, comma-separated; ranges as 'A to B'): ")
    days = parse_date_spec(spec, a.dayfirst)

    out_dir = Path(a.out) if a.out else folder / "Output"
    combined_dir = Path(a.combined_out) if a.combined_out else out_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    combined_dir.mkdir(parents=True, exist_ok=True)
    tag = f"{min(days):%Y%m%d}_to_{max(days):%Y%m%d}"

    raw_combined, combined = combine_files(folder, a.all_sheets)
    for wanted in (a.phone_col, a.website_col, a.brand_col):
        if find_column(combined, wanted) is None:
            sys.exit(f"ERROR: column '{wanted}' not found. Columns are: {list(combined.columns)}")
    filtered, date_col = filter_by_date(combined, a.date_col, days, a.dayfirst)

    # Copy 1: combined, chosen dates only, otherwise exactly as-is
    raw_filtered = raw_combined.loc[filtered.index]
    raw_date_col = next((c for c in raw_filtered.columns if " ".join(str(c).split()) == date_col), None)
    save_xlsx(raw_filtered, combined_dir / f"Combined_{tag}.xlsx", raw_date_col)

    # Copy 2: sort -> dedupe -> drop columns -> clean characters
    final = sort_and_dedupe(filtered, a.phone_col, a.website_col, a.brand_col)
    final = drop_and_clean(final, date_col, a.phone_col)
    save_xlsx(final, out_dir / f"Final_{tag}.xlsx", date_col if date_col in final.columns else None)
    print("Done.")


if __name__ == "__main__":
    main()
