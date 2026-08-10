"""
migrate.py — VCART
====================
ONE-TIME migration for the VCART Totals output workbook.

Moves existing historical time-series data from the OLD column layout
(cols 12-18 = Q1-4 Prioritized Barrier + Prioritized Impl. Barrier +
Current/Start-of-Qtr Stage, cols 19-41 = CEP onward) to the NEW layout
(cols 12-21 = the restructured Barrier Tracking block, cols 22-44 = CEP
onward, shifted +3).

Same pattern as ADFR's migrate.py — see that file for the fuller
rationale. Touches ONLY the top time-series table (row 3 header + month
rows). Does NOT touch the LIVE section (fully rewritten by main.py on its
next run anyway) or any "VCART Systems - ..." snapshot sheets.

What moves, in TS-column terms (ts_col = out_col - 6 + 3 = out_col - 3):
  - Barrier booleans (old/new ts cols 3-8): unchanged position, untouched.
  - Old ts cols 9-15 (Q1-4 Barrier, Prioritized Impl. Barrier, Current
    Stage, Start-of-Qtr Stage): RETIRED — no clean equivalent in the new
    10-field Barrier Tracking structure, so these are cleared, not moved.
  - Old ts cols 16-38 (CEP onward through Last Update): shifted +3 to new
    ts cols 19-41, values + formatting preserved exactly.

Does NOT touch the source file. Writes a new file
(<original_name>.MIGRATED.xlsx) so you can review before replacing the
original.

Usage:
    python migrate.py "C:\\path\\to\\VCART.Region.Nation.Totals.New.xlsx"
"""

import sys
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

import config
from main import MONTH_NAMES
from config import TS_HEADER_ROW, TS_START_ROW, TS_END_ROW, COL_HEADERS, SECTION_HEADERS, OUT_DATA_START

# --- Old layout constants (documented, not derived from current config,
# since config.py now only describes the NEW layout) -----------------------
OLD_MAX_OUT_COL = 41
OLD_POST_BARRIER_OUT_START = 19   # old "CARE Team CEP Stage" output col
OLD_STALE_OUT_START = 12          # old Q1 Prioritized Barrier
OLD_STALE_OUT_END   = 18          # old Start of Qtr Impl. Stage

NEW_POST_BARRIER_OUT_START = 22   # from config.POST_BARRIER_OUT_START, but pinned
                                   # here explicitly so this script is self-contained
                                   # and doesn't silently drift if config changes again
OUT_COL_DELTA = NEW_POST_BARRIER_OUT_START - OLD_POST_BARRIER_OUT_START  # +3


def _out_to_ts_col(out_col):
    return out_col - OUT_DATA_START + 3


# Stale/shift ranges, expressed in TS-column space
STALE_TS_START = _out_to_ts_col(OLD_STALE_OUT_START)              # 9
STALE_TS_END   = _out_to_ts_col(OLD_STALE_OUT_END)                # 15
SHIFT_TS_START = _out_to_ts_col(OLD_POST_BARRIER_OUT_START)       # 16
SHIFT_TS_END   = _out_to_ts_col(OLD_MAX_OUT_COL)                  # 38
TS_COL_DELTA   = OUT_COL_DELTA                                    # +3


def migrate_data_row(ws, row):
    """Shift one time-series row's old CEP-onward ts-cols to their new
    positions (right-to-left, to avoid clobbering source cells before
    they're read), then clear the retired Q1-4/Impl.Barrier/Stage block.
    Barrier boolean ts-cols (3-8) are untouched — they didn't move.
    """
    for old_ts_col in range(SHIFT_TS_END, SHIFT_TS_START - 1, -1):
        new_ts_col = old_ts_col + TS_COL_DELTA
        src = ws.cell(row=row, column=old_ts_col)
        dst = ws.cell(row=row, column=new_ts_col)
        dst.value = src.value
        if src.has_style:
            dst.font          = copy(src.font)
            dst.fill          = copy(src.fill)
            dst.alignment     = copy(src.alignment)
            dst.number_format = src.number_format

    for col in range(STALE_TS_START, STALE_TS_END + 1):
        cell = ws.cell(row=row, column=col)
        cell.value = None
        cell.fill = PatternFill(fill_type=None)


def migrate_ts_header_row(ws):
    """Rewrite row 3 (the time-series header) to the new column labels,
    leaving cols 1-2 ("Months" / "Total Barriers") untouched.
    """
    for out_col, text in COL_HEADERS.items():
        if out_col < OUT_DATA_START:
            continue
        ts_col = _out_to_ts_col(out_col)
        cell = ws.cell(row=TS_HEADER_ROW, column=ts_col, value=text)
        cell.font = Font(bold=True, size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = PatternFill(fill_type=None)
        for start, end, _, color in SECTION_HEADERS:
            if start <= out_col <= end:
                cell.fill = PatternFill("solid", fgColor=color)
                break
    # Clear anything sitting past the new max ts col, leftover from the
    # old (narrower) layout.
    max_ts_col = _out_to_ts_col(max(COL_HEADERS.keys()))
    for col in range(max_ts_col + 1, max_ts_col + 15):
        c = ws.cell(row=TS_HEADER_ROW, column=col)
        if c.value is not None:
            c.value = None


def migrate_workbook(path):
    wb = load_workbook(path)
    ws = wb["VCART Totals"]

    print("  Migrating time-series header (row 3)...")
    migrate_ts_header_row(ws)

    print("  Migrating time-series month rows...")
    migrated_months = 0
    for r in range(TS_START_ROW, TS_END_ROW + 1):
        val = ws.cell(row=r, column=1).value
        if val and any(m.lower() in str(val).lower() for m in MONTH_NAMES):
            migrate_data_row(ws, r)
            migrated_months += 1
    print(f"    {migrated_months} month row(s) migrated.")

    out_path = path.rsplit(".", 1)[0] + ".MIGRATED.xlsx"
    wb.save(out_path)
    print(f"\n  Saved migrated copy -> {out_path}")
    return out_path


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python migrate.py <path-to-existing-output.xlsx>")
        sys.exit(1)
    migrate_workbook(sys.argv[1])