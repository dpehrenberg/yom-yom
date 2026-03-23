#!/usr/bin/env python3
"""AutoAttend - Parse and display attendance data from Excel files."""

import sys
import argparse
from pathlib import Path

try:
    import xlrd
except ImportError:
    print("Error: xlrd is required. Install with: pip install xlrd")
    sys.exit(1)

# Column indices (0-based) for the specific Excel layout
COL_DATE   = 59  # תאריך  - date
COL_DAY    = 56  # יום    - day of week (Hebrew)
COL_ENTRY  = 43  # כניסה  - time in
COL_EXIT   = 39  # יציאה  - time out
COL_HOURS  = 27  # שעות בפועל - actual hours
COL_REPORT = 51  # קוד דיווח - report type

DAY_ABBREV = {
    "ראשון": "א",
    "שני":   "ב",
    "שלישי": "ג",
    "רביעי": "ד",
    "חמישי": "ה",
    "שישי":  "ו",
    "שבת":   "שבת",
}



def xl_time_str(val: float) -> str:
    """Return HH:MM from the fractional part of an Excel datetime serial."""
    if not isinstance(val, float):
        return "--:--"
    frac = val - int(val)
    total_min = round(frac * 24 * 60)
    h, m = divmod(total_min, 60)
    return f"{h:02d}:{m:02d}"


def duration_str(entry: float, exit_: float) -> str:
    """Return H:MM total time between two Excel datetime serials."""
    if not (isinstance(entry, float) and isinstance(exit_, float)):
        return "  --"
    minutes = round((exit_ - entry) * 24 * 60)
    if minutes < 0:
        return "  --"
    h, m = divmod(minutes, 60)
    return f"{h}:{m:02d}"


def xl_date_str(entry_val: float, datemode: int) -> str:
    """Return DD/MM derived from the integer (date) part of the entry timestamp."""
    t = xlrd.xldate_as_tuple(entry_val, datemode)
    return f"{t[2]:02d}/{t[1]:02d}"


def fmt_day_col(day_he: str, date_ddmm: str) -> str:
    """Return combined day column: 'יום א DD/MM' or 'שבת DD/MM'."""
    abbrev = DAY_ABBREV.get(day_he, day_he)
    if abbrev == "שבת":
        return f"שבת {date_ddmm}"
    return f"יום {abbrev} {date_ddmm}"


def parse_attendance(filepath: Path):
    """Return list of dicts with one entry per attendance row."""
    wb = xlrd.open_workbook(str(filepath))
    ws = wb.sheet_by_index(0)
    records = []
    for i in range(ws.nrows):
        row = ws.row_values(i)
        day_val    = row[COL_DAY]    if ws.ncols > COL_DAY    else ""
        entry_val  = row[COL_ENTRY]  if ws.ncols > COL_ENTRY  else ""
        exit_val   = row[COL_EXIT]   if ws.ncols > COL_EXIT   else ""
        report_val = row[COL_REPORT] if ws.ncols > COL_REPORT else ""

        # Only keep rows that have an entry time (float = Excel datetime)
        if not isinstance(entry_val, float):
            continue

        date_ddmm = xl_date_str(entry_val, wb.datemode)
        records.append({
            "day_col":  fmt_day_col(str(day_val).strip(), date_ddmm),
            "report":   str(report_val).strip() if report_val else "",
            "entry":    xl_time_str(entry_val),
            "exit":     xl_time_str(exit_val),
            "duration": duration_str(entry_val, exit_val),
        })
    return records


_L = "\u200e"  # LTR mark — anchors surrounding spaces to left-to-right context

def _col(text, width, align="left"):
    """Pad text to fixed width and append an LTR mark so trailing spaces stay LTR."""
    if align == "right":
        padded = text.rjust(width)
    else:
        padded = text.ljust(width)
    return padded + _L


def print_attendance(records: list):
    sep = _L + "-" * 58
    print()
    print(_L
          + _col("יום",        13) + "  "
          + _col("כניסה",       5, "right") + "  "
          + _col("יציאה",       5, "right") + "  "
          + _col('סה"כ',        5, "right") + "  "
          + _col("סוג דיווח",  17))
    print(sep)
    for r in records:
        print(_L
              + _col(r["day_col"],  13) + "  "
              + _col(r["entry"],     5, "right") + "  "
              + _col(r["exit"],      5, "right") + "  "
              + _col(r["duration"],  5, "right") + "  "
              + _col(r["report"],   17))
    print(sep)
    print(_L + f"  סך הכל {len(records)} ימים\n")


def main():
    parser = argparse.ArgumentParser(description="Display attendance data from an Excel file.")
    parser.add_argument("file", help="Path to the .xls / .xlsx attendance file")
    args = parser.parse_args()

    filepath = Path(args.file)
    if not filepath.exists():
        print(f"שגיאה: הקובץ לא נמצא: {filepath}")
        sys.exit(1)

    records = parse_attendance(filepath)
    if not records:
        print("לא נמצאו רשומות נוכחות.")
        sys.exit(0)

    print_attendance(records)


if __name__ == "__main__":
    main()
