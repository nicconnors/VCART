"""
VCART One-Time Migration Utility
=================================
Run this ONCE after main_vcart.py to backfill historical months
from the old Region_Nation_Totals.xlsx into the new output file.

Usage:
    python migrate_vcart.py

Reads:  OLD_FILE  (the old Region_Nation_Totals.xlsx)
Writes: OUTPUT    (the new VCART.Region.Nation.Totals.New.xlsx, must exist)
"""

import os
import traceback
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

from config import (
    OUTPUT,
    TS_START_ROW, TS_END_ROW,
    MONTH_NAMES,
)

# =============================================================================
# PATH TO OLD FILE — update if yours is named differently
# =============================================================================
OLD_FILE = r"C:\Users\nc139629\GSK\ViiV Field Reimbursement Managers - Documents\General\(VRHR) ViiV Reimbursement Health Report\2026\Region_Nation_Totals.xlsx"
OLD_SHEET = "VCART Totals"

# Old file layout: time series in col B (month name), col C (total barriers)
# Rows 3-7 contain historical months
OLD_TS_MONTH_COL  = 2   # col B
OLD_TS_TOTAL_COL  = 3   # col C


# =============================================================================
# HELPERS
# =============================================================================

def get_ts_months(ws):
    """Return existing month names in the new output file's time series."""
    result = set()
    for r in range(TS_START_ROW, TS_END_ROW + 1):
        val = ws.cell(row=r, column=1).value
        if val and any(m.lower() in str(val).lower() for m in MONTH_NAMES):
            for m in MONTH_NAMES:
                if m.lower() in str(val).lower():
                    result.add(m.lower())
    return result


def shift_ts_rows_down(ws, by, from_row, max_col):
    """Shift existing time series rows down by N to make room at the top."""
    from openpyxl.styles import Border, Alignment as Al
    from copy import copy
    for r in range(TS_END_ROW, from_row - 1, -1):
        for c in range(1, max_col + 1):
            try:
                src = ws.cell(row=r, column=c)
                dst = ws.cell(row=r + by, column=c)
                dst.value = src.value
                if src.has_style:
                    dst.font          = copy(src.font)
                    dst.fill          = copy(src.fill)
                    dst.alignment     = copy(src.alignment)
                    dst.number_format = src.number_format
                src.value = None
            except Exception:
                pass


# =============================================================================
# MAIN MIGRATION
# =============================================================================

def migrate():
    print("\n" + "=" * 55)
    print("  VCART Migration Utility")
    print("=" * 55)

    if not os.path.exists(OUTPUT):
        print("\n  Output file not found. Run main_vcart.py first.")
        input("\nPress Enter to exit...")
        return

    if not os.path.exists(OLD_FILE):
        print(f"\n  Old file not found:\n  {OLD_FILE}")
        input("\nPress Enter to exit...")
        return

    # ------------------------------------------------------------------
    print(f"\n  Reading old file...")
    wb_old = load_workbook(OLD_FILE, read_only=True, data_only=True)
    if OLD_SHEET not in wb_old.sheetnames:
        print(f"  ERROR: '{OLD_SHEET}' sheet not found in old file.")
        wb_old.close()
        input("\nPress Enter to exit...")
        return

    ws_old = wb_old[OLD_SHEET]

    # Collect historical months from old file (rows 3-8)
    # Old layout: col B=month, col C=total barriers, cols D-I=6 barrier counts (indices 2-8)
    historical = []
    for r in range(3, 9):
        row = list(ws_old.iter_rows(min_row=r, max_row=r, values_only=True))[0]
        month_val = row[1]   # col B
        total_val = row[2]   # col C
        barriers  = list(row[3:9])  # cols D-I (6 barrier counts)
        if not month_val:
            continue
        month_str_clean = str(month_val).strip().lower().replace("februaty", "february")
        matched = None
        for m in MONTH_NAMES:
            if m.lower() in month_str_clean:
                matched = m
                break
        if matched and isinstance(total_val, (int, float)):
            historical.append((matched, int(total_val), barriers))
            print(f"    Found: {matched} — {int(total_val)} barriers | {barriers}")

    wb_old.close()

    if not historical:
        print("\n  No historical months found in old file.")
        input("\nPress Enter to exit...")
        return

    # ------------------------------------------------------------------
    print(f"\n  Opening new output file...")
    wb_new = load_workbook(OUTPUT)
    if "VCART Totals" not in wb_new.sheetnames:
        print("  ERROR: 'VCART Totals' sheet not found in output file.")
        input("\nPress Enter to exit...")
        return

    ws_new = wb_new["VCART Totals"]

    # Determine which months already exist
    existing_months = get_ts_months(ws_new)
    print(f"  Existing months in new file: {existing_months or 'none'}")

    months_to_add = [
        (m, b, barriers) for m, b, barriers in historical
        if m.lower() not in existing_months
    ]

    if not months_to_add:
        print("\n  All historical months already present. Nothing to migrate.")
        wb_new.save(OUTPUT)
        input("\nPress Enter to exit...")
        return

    # Sort chronologically
    months_to_add.sort(key=lambda x: MONTH_NAMES.index(x[0]))
    n = len(months_to_add)
    print(f"\n  Backfilling {n} month(s): {[m for m, _, _ in months_to_add]}")

    # Find first filled TS row to shift down from
    first_filled = None
    for r in range(TS_START_ROW, TS_END_ROW + 1):
        val = ws_new.cell(row=r, column=1).value
        if val and str(val).strip():
            first_filled = r
            break

    # Figure out max col used
    from config import MAX_OUT_COL
    if first_filled:
        shift_ts_rows_down(ws_new, n, first_filled, MAX_OUT_COL)

    # Write historical months at the top of the TS block
    for i, (month_name, total_barriers, barriers) in enumerate(months_to_add):
        write_row = TS_START_ROW + i
        if write_row > TS_END_ROW:
            print(f"  WARNING: TS full, skipping {month_name}")
            continue
        cell = ws_new.cell(row=write_row, column=1, value=month_name)
        cell.font = Font(size=11)
        cell.fill = PatternFill(fill_type=None)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell2 = ws_new.cell(row=write_row, column=2, value=total_barriers)
        cell2.font = Font(size=11)
        cell2.fill = PatternFill(fill_type=None)
        # Write per-barrier counts into TS cols 3-8
        for j, val in enumerate(barriers):
            if val is not None and isinstance(val, (int, float)):
                c = ws_new.cell(row=write_row, column=3 + j, value=val)
                c.font = Font(size=11)
                c.fill = PatternFill(fill_type=None)
        ws_new.row_dimensions[write_row].height = 16
        print(f"    Written: {month_name} ({total_barriers} barriers) → row {write_row}")

    wb_new.save(OUTPUT)
    print(f"\n{'=' * 55}")
    print("  Migration complete. You do not need to run this again.")
    print(f"{'=' * 55}")
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    try:
        migrate()
    except Exception as e:
        print(f"\nError: {e}")
        print(traceback.format_exc())
        input("\nPress Enter to exit...")