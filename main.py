"""
VCART VRHR Aggregator
=====================
Reads the VRHR.VCART sheet from each per-territory file, aggregates
campus-level data into a single workbook with:
  - "VCART Totals" sheet  — territory-level summary (raw + %) with
    monthly time series and snapshot history
  - "VCART Systems - LIVE" sheet — campus-level live data (all rows)
  - Snapshot preservation and monthly rollover (same pattern as VRHR)

Usage:
    python main_vcart.py

Put this file and config.py in the same folder.
Set BASE_DIR in config.py to the folder containing all territory files.
"""

import os
import glob
import traceback
import warnings
from copy import copy
from datetime import datetime
from collections import Counter

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, Reference

from config import (
    BASE_DIR, OUTPUT,
    TERRITORIES, REGION_MAP,
    SOURCE_SHEET_KEYWORDS, HEADER_ROW, DATA_START,
    SRC_DATA_START, SRC_DATA_END, NUM_DATA_COLS,
    OUT_LABEL_COL, OUT_TERRNAME_COL, OUT_TERR_COL, OUT_NAME_COL, OUT_CAMPUS_COL,
    OUT_DATA_START, MAX_OUT_COL,
    BARRIER_COLS_OUT, NATION_SUM_COLS, PCT_COLS_OUT,
    SECTION_HEADERS, TS_SECTION_HEADERS, COL_HEADERS,
    YELLOW, NATION_FILL, EAST_FILL, WEST_FILL,
    LIVE_HDR_PEACH, LIVE_HDR_GREEN, LIVE_HDR_NAVY,
    MONTH_NAMES, TS_HEADER_ROW, TS_START_ROW, TS_END_ROW,
)

warnings.filterwarnings("ignore")

NUM_TERRITORIES = len(TERRITORIES)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _src_to_out(src_col):
    """Convert 1-based source column index to 1-based output column index."""
    return OUT_DATA_START + (src_col - SRC_DATA_START)


def _find_source_sheet(wb):
    for name in wb.sheetnames:
        nl = name.lower()
        if all(k in nl for k in SOURCE_SHEET_KEYWORDS):
            return wb[name]
    return None


def _section_fill(out_col, headers=None):
    if headers is None:
        headers = SECTION_HEADERS
    for start, end, _, color in headers:
        if start <= out_col <= end:
            return PatternFill("solid", fgColor=color)
    return PatternFill(fill_type=None)


# ---------------------------------------------------------------------------
# FILE READING
# ---------------------------------------------------------------------------

def read_territory_file(filepath, terr_label, terr_override):
    """
    Read one territory VCART file.
    Returns:
        terr_num    - territory number string
        vcart_name  - VCART rep name
        campus_rows - list of dicts {out_col: value} for each data campus row
        totals_row  - dict {out_col: value} from the pre-computed TOTALS row (row 3)
    """
    try:
        wb = load_workbook(filepath, read_only=True, data_only=True)
    except Exception as e:
        print(f"  ERROR opening {os.path.basename(filepath)}: {e}")
        return None, None, [], {}

    ws = _find_source_sheet(wb)
    if ws is None:
        print(f"  WARNING: no VRHR.VCART sheet found in {os.path.basename(filepath)}")
        wb.close()
        return None, None, [], {}

    # ------------------------------------------------------------------
    # Read TOTALS row (row 3, cols I–AR = indices 8..43)
    # ------------------------------------------------------------------
    totals_raw = {}
    for row in ws.iter_rows(min_row=3, max_row=3, values_only=True):
        for src_idx in range(SRC_DATA_START - 1, SRC_DATA_END):   # 0-based
            v = row[src_idx] if src_idx < len(row) else None
            out_col = OUT_DATA_START + (src_idx - (SRC_DATA_START - 1))
            totals_raw[out_col] = v if v != "#N/A" else None
        break

    # ------------------------------------------------------------------
    # Read data rows (row 7 onward)
    # ------------------------------------------------------------------
    vcart_name = None
    terr_num   = terr_override
    campus_rows = []

    all_rows = list(ws.iter_rows(min_row=DATA_START, values_only=True))
    wb.close()

    for row in all_rows:
        # Stop at fully empty row
        if all(v is None or v == "" for v in row[:SRC_DATA_END]):
            continue

        # Territory # from col A (0-indexed: 0), fall back to override
        if terr_num is None or terr_num == terr_override:
            if row[0] and str(row[0]).strip():
                terr_num = str(row[0]).strip()

        # VCART name from col B (index 1)
        if vcart_name is None and row[1] and str(row[1]).strip():
            vcart_name = str(row[1]).strip()

        # Check if this is a real campus row:
        # col A (territory #) or col B (name) or col C (CID) must be non-empty
        has_identity = (
            (row[0] and str(row[0]).strip()) or
            (row[1] and str(row[1]).strip()) or
            (row[2] and isinstance(row[2], int))
        )
        if not has_identity:
            continue  # blank template rows — skip
        if not isinstance(row[8], (bool, int, float)):
            continue

        rd = {}
        # Store demographics: cols A-H (indices 0-7)
        rd["terr_num"]   = row[0] if row[0] else terr_num
        rd["vcart_name"] = row[1] if row[1] else vcart_name
        rd["cid"]        = row[2] if len(row) > 2 else None
        rd["corp_name"]  = row[3] if len(row) > 3 else None
        rd["address"]    = row[4] if len(row) > 4 else None
        rd["city"]       = row[5] if len(row) > 5 else None
        rd["state"]      = row[6] if len(row) > 6 else None
        rd["zip"]        = row[7] if len(row) > 7 else None
        # Store data cols I-AR (indices 8-43)
        for src_idx in range(SRC_DATA_START - 1, SRC_DATA_END):  # 0-based
            v = row[src_idx] if src_idx < len(row) else None
            out_col = OUT_DATA_START + (src_idx - (SRC_DATA_START - 1))
            rd[out_col] = v if v != "#N/A" else None
        campus_rows.append(rd)

    if terr_num is None:
        terr_num = terr_override or "UNKNOWN"
    if vcart_name is None:
        vcart_name = terr_label

    return terr_num, vcart_name, campus_rows, totals_raw


def load_all_territories(folder):
    """
    Load all territory files. Returns dict:
        label → {terr_num, vcart_name, campus_rows, totals, campus_count, region}
    """
    result = {}
    for label, filename, terr_override in TERRITORIES:
        filepath = os.path.join(folder, filename)
        if not os.path.exists(filepath):
            # try glob match in case filename differs slightly
            matches = glob.glob(os.path.join(folder, f"*{filename.split('_')[0]}*"))
            if matches:
                filepath = matches[0]
                print(f"  NOTE: fuzzy matched {label} → {os.path.basename(filepath)}")
            else:
                print(f"  WARNING: file not found for {label}: {filename}")
                result[label] = None
                continue

        print(f"  Reading {label}...")
        terr_num, vcart_name, campus_rows, totals = read_territory_file(
            filepath, label, terr_override
        )
        if campus_rows is None and not totals:
            result[label] = None
            continue

        # Use actual campus count from data rows
        campus_count = len(campus_rows)
        result[label] = {
            "terr_num":    terr_num,
            "vcart_name":  vcart_name,
            "campus_rows": campus_rows,
            "totals":      totals,
            "campus_count": campus_count,
            "region":      REGION_MAP.get(label, "VCART - Unknown"),
        }
        print(f"    ✓ {label}: {campus_count} campuses (terr: {terr_num})")
    return result


# ---------------------------------------------------------------------------
# STYLE HELPERS
# ---------------------------------------------------------------------------

def unmerge_row(ws, row):
    to_remove = [str(m) for m in ws.merged_cells.ranges
                 if m.min_row <= row <= m.max_row]
    for m in to_remove:
        ws.unmerge_cells(m)


def get_last_ts_row(ws):
    last = TS_HEADER_ROW
    for r in range(TS_START_ROW, TS_END_ROW + 1):
        val = ws.cell(row=r, column=1).value
        if val and any(m.lower() in str(val).lower() for m in MONTH_NAMES):
            last = r
    return last


def get_live_base(ws):
    return get_last_ts_row(ws) + 2


def get_live_offsets(ws):
    base = get_live_base(ws)
    raw_nat = base + 3 + NUM_TERRITORIES
    return {
        "live_row":      base,
        "raw_sec_row":   base + 1,
        "raw_hdr_row":   base + 2,
        "raw_data_row":  base + 3,
        "raw_nat_row":   raw_nat,
        "gap_row":       raw_nat + 1,
        "pct_label_row": raw_nat + 2,
        "pct_sec_row":   raw_nat + 3,
        "pct_hdr_row":   raw_nat + 4,
        "pct_data_row":  raw_nat + 5,
        "pct_nat_row":   raw_nat + 5 + NUM_TERRITORIES,
    }


def get_snap_start(ws):
    o = get_live_offsets(ws)
    return o["pct_nat_row"] + 4


def is_yellow_row(ws, row):
    try:
        fill = ws.cell(row=row, column=1).fill
        if fill and fill.fgColor:
            rgb = fill.fgColor.rgb
            return rgb in ("FFFFFF00", "00FFFF00", "FFFF00", f"FF{YELLOW}", YELLOW)
    except Exception:
        pass
    return False


def find_yellow_rows(ws, start_row, end_row):
    yellow_rows = []
    for r in range(start_row, end_row + 1):
        if is_yellow_row(ws, r):
            val = ws.cell(row=r, column=1).value
            if val and str(val).strip():
                yellow_rows.append(r)
    return yellow_rows


# ---------------------------------------------------------------------------
# TABLE WRITERS — VCART Totals sheet
# ---------------------------------------------------------------------------

def write_section_headers(ws, row, headers=None):
    if headers is None:
        headers = SECTION_HEADERS
    unmerge_row(ws, row)
    for start, end, label, color in headers:
        cell = ws.cell(row=row, column=start, value=label)
        cell.font = Font(bold=True, size=9)
        cell.fill = PatternFill("solid", fgColor=color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for c in range(start + 1, end + 1):
            ws.cell(row=row, column=c).value = None
            ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor=color)
        if end > start:
            try:
                ws.merge_cells(start_row=row, start_column=start,
                               end_row=row, end_column=end)
            except Exception:
                pass
    ws.row_dimensions[row].height = 20


def write_live_col_headers(ws, row):
    """Write LIVE section column headers with custom per-zone colors."""
    for col, text in COL_HEADERS.items():
        cell = ws.cell(row=row, column=col, value=text)
        cell.font = Font(size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if 1 <= col <= 5:
            cell.fill = PatternFill("solid", fgColor=LIVE_HDR_PEACH)
        elif 6 <= col <= 15:
            cell.fill = PatternFill("solid", fgColor=LIVE_HDR_GREEN)
        elif 16 <= col <= 19:
            cell.fill = PatternFill("solid", fgColor=LIVE_HDR_NAVY)
        else:
            cell.fill = _section_fill(col)
    ws.row_dimensions[row].height = 80


def write_col_headers(ws, row, headers=None):
    for col, text in COL_HEADERS.items():
        cell = ws.cell(row=row, column=col, value=text)
        cell.font = Font(size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = _section_fill(col, headers)
    ws.row_dimensions[row].height = 80
    if "East" in region:
        return PatternFill("solid", fgColor=EAST_FILL)
    if "West" in region:
        return PatternFill("solid", fgColor=WEST_FILL)
    return PatternFill(fill_type=None)


def write_territory_row(ws, row_num, label, td, raw=True):
    """Write one territory row to the VCART Totals sheet."""
    for c in range(1, MAX_OUT_COL + 1):
        ws.cell(row=row_num, column=c).value = None
        ws.cell(row=row_num, column=c).fill = PatternFill(fill_type=None)

    if td is None:
        ws.cell(row=row_num, column=OUT_LABEL_COL, value=label)
        return

    region = td["region"]
    rfill  = PatternFill(fill_type=None)  # white — all data rows

    def _set(col, val, bold=False, fmt=None):
        cell = ws.cell(row=row_num, column=col, value=val)
        cell.font = Font(size=11, bold=bold)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        cell.fill = rfill
        if fmt:
            cell.number_format = fmt

    _set(OUT_LABEL_COL,   td["region"])
    _set(OUT_TERRNAME_COL, label)
    _set(OUT_TERR_COL,    td["terr_num"])
    _set(OUT_NAME_COL,    td["vcart_name"])
    _set(OUT_CAMPUS_COL,  td["campus_count"])

    campus_count = td["campus_count"]
    totals = td["totals"]  # pre-computed sums from file's TOTALS row

    for out_col in range(OUT_DATA_START, MAX_OUT_COL + 1):
        val = totals.get(out_col)
        cell = ws.cell(row=row_num, column=out_col)
        cell.font = Font(size=11)
        cell.fill = rfill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)

        if val is None or str(val) in ("#N/A", "None"):
            cell.value = None
            continue

        if not raw:
            if out_col in BARRIER_COLS_OUT and campus_count and isinstance(val, (int, float)):
                cell.value = val / campus_count
                cell.number_format = "0%"
                continue
            if out_col in PCT_COLS_OUT and isinstance(val, (int, float)):
                cell.value = val  # already a pct (0–1) or store as-is
                cell.number_format = "0%"
                continue
        cell.value = val

    ws.row_dimensions[row_num].height = 15


def write_nation_row(ws, row_num, all_td, raw=True):
    """Write NATION summary row."""
    for c in range(1, MAX_OUT_COL + 1):
        ws.cell(row=row_num, column=c).value = None

    nation_fill = PatternFill("solid", fgColor=NATION_FILL)

    def _set(col, val, bold=True, fmt=None):
        cell = ws.cell(row=row_num, column=col, value=val)
        cell.font = Font(size=11, bold=bold)
        cell.fill = nation_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if fmt:
            cell.number_format = fmt

    total_campuses = sum(td["campus_count"] for td in all_td.values() if td)
    _set(OUT_LABEL_COL, "NATION")
    _set(OUT_TERRNAME_COL, "")
    _set(OUT_TERR_COL,  "")
    _set(OUT_NAME_COL,  "")
    _set(OUT_CAMPUS_COL, total_campuses)

    valid = [td for td in all_td.values() if td]
    for out_col in NATION_SUM_COLS:
        vals = [td["totals"].get(out_col) for td in valid
                if td["totals"].get(out_col) is not None]
        numeric = [v for v in vals if isinstance(v, (int, float))]
        cell = ws.cell(row=row_num, column=out_col)
        cell.font = Font(size=11, bold=True)
        cell.fill = nation_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if numeric:
            total = sum(numeric)
            if raw:
                cell.value = total
            elif total_campuses:
                cell.value = total / total_campuses
                cell.number_format = "0%"

    ws.row_dimensions[row_num].height = 15


# ---------------------------------------------------------------------------
# VCART Totals sheet — structure
# ---------------------------------------------------------------------------

def write_totals_sheet_headers(ws):
    ws.cell(row=1, column=1, value="VCART Territory Totals").font = Font(bold=True, size=14)
    ws.row_dimensions[1].height = 22

    # Row 3: section color bands (merged), Row 4 would overlap with data so we combine:
    # Use TS_HEADER_ROW for both section colors AND col labels — write section fills first,
    # then overwrite with col labels (openpyxl allows setting value on merged cell anchor)
    unmerge_row(ws, TS_HEADER_ROW)

    # Write section color fills across the TS header row
    for start, end, label, color in TS_SECTION_HEADERS:
        for c in range(start, end + 1):
            ws.cell(row=TS_HEADER_ROW, column=c).fill = PatternFill("solid", fgColor=color)

    # Col 1 = Months, col 2 = Total Barriers
    for col, text, fill in [
        (1, "Months",         PatternFill(fill_type=None)),
        (2, "Total\nBarriers", PatternFill(fill_type=None)),
    ]:
        cell = ws.cell(row=TS_HEADER_ROW, column=col, value=text)
        cell.font = Font(bold=True, size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = fill

    # Map data cols (6..41) → TS cols (3..38), write col header text with section fill
    for out_col in range(OUT_DATA_START, MAX_OUT_COL + 1):
        ts_col = out_col - OUT_DATA_START + 3
        text = COL_HEADERS.get(out_col, "")
        cell = ws.cell(row=TS_HEADER_ROW, column=ts_col, value=text)
        cell.font = Font(bold=True, size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = _section_fill(ts_col, TS_SECTION_HEADERS)
    ws.row_dimensions[TS_HEADER_ROW].height = 80

    for r in range(TS_START_ROW, TS_END_ROW + 1):
        ws.row_dimensions[r].height = 16

    # Column widths
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 10
    for col_letter in ["C","D","E","F","G","H"]:
        ws.column_dimensions[col_letter].width = 10
    for col_letter in ["I","J","K","L","M","N","O","P"]:
        ws.column_dimensions[col_letter].width = 14
    for col_letter in ["Q","R","S","T","U","V","W","X","Y","Z"]:
        ws.column_dimensions[col_letter].width = 11
    for idx in range(27, 42):
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(idx)].width = 11


def write_live_section(ws, all_td, snap_date):
    o = get_live_offsets(ws)

    date_str = snap_date.strftime("%m.%d.%y").lstrip("0").replace(".0", ".")
    ws.cell(row=o["live_row"], column=1,
            value=f"LIVE - refreshed {date_str}").font = Font(size=11)
    ws.cell(row=o["live_row"], column=1).fill = PatternFill("solid", fgColor=YELLOW)

    total_barriers = sum(
        sum(td["totals"].get(c, 0) or 0 for c in BARRIER_COLS_OUT
            if isinstance(td["totals"].get(c), (int, float)) and not isinstance(td["totals"].get(c), bool))
        for td in all_td.values() if td
    )
    ws.cell(row=o["live_row"], column=2,
            value=f"Total barriers: {total_barriers}").font = Font(size=11)
    ws.cell(row=o["live_row"], column=2).fill = PatternFill("solid", fgColor=YELLOW)

    # RAW section
    write_section_headers(ws, o["raw_sec_row"])
    write_live_col_headers(ws, o["raw_hdr_row"])
    for i, (label, *_) in enumerate(TERRITORIES):
        write_territory_row(ws, o["raw_data_row"] + i, label, all_td.get(label), raw=True)
    write_nation_row(ws, o["raw_nat_row"], all_td, raw=True)
    ws.row_dimensions[o["gap_row"]].height = 8

    # PCT section
    ws.cell(row=o["pct_label_row"], column=1,
            value="% of LIVE").font = Font(size=11, bold=True)
    ws.cell(row=o["pct_label_row"], column=1).fill = PatternFill("solid", fgColor=YELLOW)
    ws.row_dimensions[o["pct_label_row"]].height = 16
    write_section_headers(ws, o["pct_sec_row"])
    write_live_col_headers(ws, o["pct_hdr_row"])
    for i, (label, *_) in enumerate(TERRITORIES):
        write_territory_row(ws, o["pct_data_row"] + i, label, all_td.get(label), raw=False)
    write_nation_row(ws, o["pct_nat_row"], all_td, raw=False)
    ws.row_dimensions[o["pct_nat_row"] + 1].height = 8


# ---------------------------------------------------------------------------
# TIME SERIES
# ---------------------------------------------------------------------------

def update_time_series(ws, all_td, snap_date):
    month_name = snap_date.strftime("%B")
    month_row  = None

    for r in range(TS_START_ROW, TS_END_ROW + 1):
        cell_val = ws.cell(row=r, column=1).value
        if not cell_val or str(cell_val).strip() == "":
            month_row = r
            break
        if month_name.lower() in str(cell_val).lower():
            month_row = r
            break

    if not month_row:
        for r in range(TS_START_ROW, TS_END_ROW):
            for c in range(1, MAX_OUT_COL + 1):
                try:
                    ws.cell(row=r, column=c).value = ws.cell(row=r + 1, column=c).value
                except Exception:
                    pass
        month_row = TS_END_ROW

    valid = [td for td in all_td.values() if td]
    total_barriers = sum(
        sum(td["totals"].get(c, 0) or 0 for c in BARRIER_COLS_OUT
            if isinstance(td["totals"].get(c), (int, float)))
        for td in valid
    )
    total_campuses = sum(td["campus_count"] for td in valid)

    unmerge_row(ws, month_row)
    date_label = f"{month_name} ({snap_date.strftime('%m/%d/%y')})"
    ws.cell(row=month_row, column=1, value=date_label).font = Font(size=11)
    ws.cell(row=month_row, column=1).fill = PatternFill(fill_type=None)
    ws.cell(row=month_row, column=2, value=total_barriers).font = Font(size=11)
    ws.cell(row=month_row, column=2).fill = PatternFill(fill_type=None)
    ws.row_dimensions[month_row].height = 16

    # Write data into shifted TS cols (out col 6 → ts col 3, skipping identity cols 1-5)
    for out_col in range(OUT_DATA_START, MAX_OUT_COL + 1):
        ts_col = out_col - OUT_DATA_START + 3
        cell = ws.cell(row=month_row, column=ts_col)
        cell.fill = PatternFill(fill_type=None)
        if out_col in BARRIER_COLS_OUT:
            nums = [td["totals"].get(out_col) for td in valid
                    if isinstance(td["totals"].get(out_col), (int, float))]
            if nums:
                cell.value = sum(nums)
                cell.font = Font(size=11)
        else:
            vals = [str(td["totals"].get(out_col) or "").strip() for td in valid
                    if td["totals"].get(out_col)]
            vals = [v for v in vals if v and v not in ("#N/A", "None")]
            if vals:
                cell.value = Counter(vals).most_common(1)[0][0]
                cell.font = Font(size=11)
                cell.alignment = Alignment(horizontal="center", vertical="center")


# ---------------------------------------------------------------------------
# SNAPSHOT MANAGEMENT
# ---------------------------------------------------------------------------

def append_snapshot(ws, all_td, snap_date, live_month):
    snap_start = get_snap_start(ws)
    block_size = 1 + (1 + 1 + NUM_TERRITORIES + 1) + 1 + 1 + (1 + 1 + NUM_TERRITORIES + 1) + 1

    max_row = ws.max_row
    for r in range(max_row, snap_start - 1, -1):
        ws.row_dimensions[r + block_size].height = ws.row_dimensions[r].height
        for c in range(1, MAX_OUT_COL + 1):
            try:
                src = ws.cell(row=r, column=c)
                dst = ws.cell(row=r + block_size, column=c)
                dst.value = src.value
                if src.has_style:
                    dst.font          = copy(src.font)
                    dst.fill          = copy(src.fill)
                    dst.border        = copy(src.border)
                    dst.alignment     = copy(src.alignment)
                    dst.number_format = src.number_format
                src.value = None
                src.fill  = PatternFill(fill_type=None)
                src.font  = Font()
                src.border = Border()
                src.alignment = Alignment()
                src.number_format = "General"
            except Exception:
                pass

    _trim_snapshots(ws, keep=11)

    date_str = snap_date.strftime("%m/%d/%y")
    total_barriers = sum(
        sum(td["totals"].get(c, 0) or 0 for c in BARRIER_COLS_OUT
            if isinstance(td["totals"].get(c), (int, float)))
        for td in all_td.values() if td
    )

    r = snap_start
    ws.cell(row=r, column=1,
            value=f"{live_month} snapshot ({date_str})").fill = PatternFill("solid", fgColor=YELLOW)
    ws.cell(row=r, column=1).font = Font(size=11)
    ws.cell(row=r, column=2, value=f"Total barriers - {total_barriers}").font = Font(size=11)
    r += 1

    write_section_headers(ws, r); r += 1
    write_col_headers(ws, r);    r += 1
    for i, (label, *_) in enumerate(TERRITORIES):
        write_territory_row(ws, r + i, label, all_td.get(label), raw=True)
    r += NUM_TERRITORIES
    write_nation_row(ws, r, all_td, raw=True); r += 2

    write_section_headers(ws, r); r += 1
    write_col_headers(ws, r);    r += 1
    for i, (label, *_) in enumerate(TERRITORIES):
        write_territory_row(ws, r + i, label, all_td.get(label), raw=False)
    r += NUM_TERRITORIES
    write_nation_row(ws, r, all_td, raw=False)


def _trim_snapshots(ws, keep=11):
    snap_start = get_snap_start(ws)
    snap_rows  = find_yellow_rows(ws, snap_start, ws.max_row)
    if len(snap_rows) <= keep:
        return
    to_delete    = len(snap_rows) - keep
    delete_up_to = snap_rows[to_delete] - 2
    for r in range(snap_start, delete_up_to + 1):
        for c in range(1, MAX_OUT_COL + 1):
            try:
                ws.cell(row=r, column=c).value = None
            except Exception:
                pass
    remaining, shift_to = delete_up_to + 1, snap_start
    while remaining <= ws.max_row:
        for c in range(1, MAX_OUT_COL + 1):
            try:
                ws.cell(row=shift_to, column=c).value = ws.cell(row=remaining, column=c).value
                ws.cell(row=remaining, column=c).value = None
            except Exception:
                pass
        shift_to += 1
        remaining += 1


# ---------------------------------------------------------------------------
# EXISTING FILE READER (for preserving time series + snapshots)
# ---------------------------------------------------------------------------

def read_existing_totals(filepath):
    existing_ts      = []
    existing_snaps   = []
    live_month       = None
    live_td          = {}

    if not os.path.exists(filepath):
        return existing_ts, live_month, existing_snaps, live_td

    try:
        wb = load_workbook(filepath, data_only=True)
        if "VCART Totals" not in wb.sheetnames:
            wb.close()
            return existing_ts, live_month, existing_snaps, live_td

        ws = wb["VCART Totals"]

        # Time series
        for r in range(TS_START_ROW, TS_END_ROW + 1):
            val = ws.cell(row=r, column=1).value
            if val and any(m.lower() in str(val).lower() for m in MONTH_NAMES):
                existing_ts.append(
                    [ws.cell(row=r, column=c).value for c in range(1, MAX_OUT_COL + 1)]
                )

        # Live month name
        for r in range(1, ws.max_row + 1):
            v = str(ws.cell(row=r, column=1).value or "")
            if "live" in v.lower():
                for m in MONTH_NAMES:
                    if m.lower() in v.lower():
                        live_month = m
                        break
                break

        # Snapshots
        all_yellow = find_yellow_rows(ws, TS_END_ROW + 1, ws.max_row)
        snap_yellow = [
            rr for rr in all_yellow
            if "live" not in str(ws.cell(row=rr, column=1).value or "").lower()
        ]
        for i, snap_start in enumerate(snap_yellow):
            snap_end = snap_yellow[i + 1] - 2 if i + 1 < len(snap_yellow) else ws.max_row
            snap_block = [
                [ws.cell(row=sr, column=c).value for c in range(1, MAX_OUT_COL + 1)]
                for sr in range(snap_start, snap_end + 1)
            ]
            existing_snaps.append(snap_block)

        # Live territory data (for snapshot on month rollover)
        for r in range(1, ws.max_row + 1):
            v = str(ws.cell(row=r, column=1).value or "")
            if "live" in v.lower():
                o = get_live_offsets(ws)
                for i, (label, *_) in enumerate(TERRITORIES):
                    row_num = o["raw_data_row"] + i
                    td_snap = {
                        "terr_num":    ws.cell(row=row_num, column=OUT_TERR_COL).value,
                        "vcart_name":  ws.cell(row=row_num, column=OUT_NAME_COL).value,
                        "campus_count": ws.cell(row=row_num, column=OUT_CAMPUS_COL).value or 0,
                        "region":      ws.cell(row=row_num, column=OUT_LABEL_COL).value or "",
                        "totals": {
                            out_col: ws.cell(row=row_num, column=out_col).value
                            for out_col in range(OUT_DATA_START, MAX_OUT_COL + 1)
                        },
                    }
                    live_td[label] = td_snap
                break

        wb.close()
        print(f"\n  Found {len(existing_ts)} time series rows, {len(existing_snaps)} snapshots")
    except Exception as e:
        print(f"  Could not read existing file: {e}")

    return existing_ts, live_month, existing_snaps, live_td


# ---------------------------------------------------------------------------
# VCART SYSTEMS — LIVE sheet (campus-level)
# ---------------------------------------------------------------------------

def build_systems_live_sheet(wb, all_td, snap_date):
    """Write campus-level data to 'VCART Systems - LIVE' sheet."""
    if "VCART Systems - LIVE" in wb.sheetnames:
        del wb["VCART Systems - LIVE"]
    ws = wb.create_sheet("VCART Systems - LIVE")

    # Section header row (row 1), col headers (row 2), data from row 3
    ws.cell(row=1, column=1, value=f"VCART Systems - LIVE  (refreshed {snap_date.strftime('%m/%d/%y')})").font = Font(bold=True, size=12)

    # Compute nation totals row for row 2
    total_campuses = sum(td["campus_count"] for td in all_td.values() if td)
    nation_barriers = [
        sum(td["totals"].get(c, 0) or 0 for td in all_td.values()
            if td and isinstance(td["totals"].get(c), (int, float)))
        for c in BARRIER_COLS_OUT
    ]
    totals_label_row = 2
    ws.cell(row=totals_label_row, column=8, value="TOTALS").font = Font(bold=True, size=9)
    ws.cell(row=totals_label_row, column=8).fill = PatternFill("solid", fgColor=NATION_FILL)
    for i, bc in enumerate(BARRIER_COLS_OUT):
        out_col_offset = bc - OUT_DATA_START  # 0-based offset into barrier block
        ws.cell(row=totals_label_row, column=9 + out_col_offset,
                value=nation_barriers[i]).font = Font(size=9, bold=True)

    # Section header row 3
    write_section_headers(ws, 3)

    # Column header row 4
    # Campus-level headers: cols A-H = demographics, then data cols I-AR
    campus_hdrs = {
        1: "VCART Territory #",
        2: "VCART Name",
        3: "Corporate\nParent CID",
        4: "Corporate\nParent Name",
        5: "Address",
        6: "City",
        7: "State",
        8: "Zip",
    }
    # Shift section headers and data headers by 4 (demographics take cols 1-8,
    # data cols 9..44 = source col indices mapped directly)
    for col, text in campus_hdrs.items():
        cell = ws.cell(row=4, column=col, value=text)
        cell.font = Font(size=9, bold=False)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for out_col, text in COL_HEADERS.items():
        if out_col >= OUT_DATA_START:
            src_col = SRC_DATA_START + (out_col - OUT_DATA_START)  # 1-based source col
            ws_col  = src_col  # campus sheet: cols 1-8 are demog, cols 9-44 are data
            cell = ws.cell(row=4, column=ws_col, value=text)
            cell.font = Font(size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = _section_fill(out_col)
    ws.row_dimensions[4].height = 80

    # Data rows
    current_row = 5
    for label, filename, _ in TERRITORIES:
        td = all_td.get(label)
        if td is None or not td["campus_rows"]:
            continue
        for campus in td["campus_rows"]:
            ws.cell(row=current_row, column=1, value=campus.get("terr_num", td["terr_num"])).font = Font(size=10)
            ws.cell(row=current_row, column=2, value=campus.get("vcart_name", td["vcart_name"])).font = Font(size=10)
            ws.cell(row=current_row, column=3, value=campus.get("cid")).font = Font(size=10)
            ws.cell(row=current_row, column=4, value=campus.get("corp_name")).font = Font(size=10)
            ws.cell(row=current_row, column=5, value=campus.get("address")).font = Font(size=10)
            ws.cell(row=current_row, column=6, value=campus.get("city")).font = Font(size=10)
            ws.cell(row=current_row, column=7, value=campus.get("state")).font = Font(size=10)
            ws.cell(row=current_row, column=8, value=campus.get("zip")).font = Font(size=10)

            for out_col, val in campus.items():
                if not isinstance(out_col, int):
                    continue  # skip demographic keys
                src_offset = out_col - OUT_DATA_START  # 0-based
                ws_col = SRC_DATA_START + src_offset   # 1-based: matches header layout
                cell = ws.cell(row=current_row, column=ws_col, value=val)
                cell.font = Font(size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[current_row].height = 14
            current_row += 1

    # Column widths
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 6
    ws.column_dimensions["H"].width = 7
    from openpyxl.utils import get_column_letter
    for i in range(9, 45):
        ws.column_dimensions[get_column_letter(i)].width = 12

    print(f"    VCART Systems - LIVE: {current_row - 5} campus rows written")


def build_systems_snapshot_sheet(wb, all_td, snap_date, month_name):
    """Write a frozen campus-level snapshot sheet named 'VCART Systems - {Mon YYYY}'."""
    from openpyxl.utils import get_column_letter
    sheet_name = f"VCART Systems - {month_name[:3]} {snap_date.strftime('%Y')}"
    # Remove if already exists (re-run same month)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    ws.cell(row=1, column=1,
            value=f"VCART Systems Snapshot — {month_name} {snap_date.strftime('%Y')}  (saved {snap_date.strftime('%m/%d/%y')})").font = Font(bold=True, size=12)

    # Reuse same layout as LIVE sheet
    total_campuses = sum(td["campus_count"] for td in all_td.values() if td)
    nation_barriers = [
        sum(td["totals"].get(c, 0) or 0 for td in all_td.values()
            if td and isinstance(td["totals"].get(c), (int, float)) and not isinstance(td["totals"].get(c), bool))
        for c in BARRIER_COLS_OUT
    ]
    ws.cell(row=2, column=8, value="TOTALS").font = Font(bold=True, size=9)
    ws.cell(row=2, column=8).fill = PatternFill("solid", fgColor=NATION_FILL)
    for i, bc in enumerate(BARRIER_COLS_OUT):
        ws.cell(row=2, column=9 + (bc - OUT_DATA_START), value=nation_barriers[i]).font = Font(size=9, bold=True)

    write_section_headers(ws, 3)

    campus_hdrs = {
        1: "VCART Territory #", 2: "VCART Name", 3: "Corporate\nParent CID",
        4: "Corporate\nParent Name", 5: "Address", 6: "City", 7: "State", 8: "Zip",
    }
    for col, text in campus_hdrs.items():
        cell = ws.cell(row=4, column=col, value=text)
        cell.font = Font(size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for out_col, text in COL_HEADERS.items():
        if out_col >= OUT_DATA_START:
            ws_col = SRC_DATA_START + (out_col - OUT_DATA_START)
            cell = ws.cell(row=4, column=ws_col, value=text)
            cell.font = Font(size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = _section_fill(out_col)
    ws.row_dimensions[4].height = 80

    current_row = 5
    for label, filename, _ in TERRITORIES:
        td = all_td.get(label)
        if td is None or not td["campus_rows"]:
            continue
        for campus in td["campus_rows"]:
            ws.cell(row=current_row, column=1, value=campus.get("terr_num", td["terr_num"])).font = Font(size=10)
            ws.cell(row=current_row, column=2, value=campus.get("vcart_name", td["vcart_name"])).font = Font(size=10)
            ws.cell(row=current_row, column=3, value=campus.get("cid")).font = Font(size=10)
            ws.cell(row=current_row, column=4, value=campus.get("corp_name")).font = Font(size=10)
            ws.cell(row=current_row, column=5, value=campus.get("address")).font = Font(size=10)
            ws.cell(row=current_row, column=6, value=campus.get("city")).font = Font(size=10)
            ws.cell(row=current_row, column=7, value=campus.get("state")).font = Font(size=10)
            ws.cell(row=current_row, column=8, value=campus.get("zip")).font = Font(size=10)
            for out_col, val in campus.items():
                if not isinstance(out_col, int):
                    continue
                ws_col = SRC_DATA_START + (out_col - OUT_DATA_START)
                cell = ws.cell(row=current_row, column=ws_col, value=val)
                cell.font = Font(size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[current_row].height = 14
            current_row += 1

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 6
    ws.column_dimensions["H"].width = 7
    for i in range(9, 45):
        ws.column_dimensions[get_column_letter(i)].width = 12

    print(f"        {sheet_name}: {current_row - 5} campus rows written")


def copy_existing_systems_sheets(src_filepath, dst_wb):
    """Copy any existing 'VCART Systems - Mon YYYY' snapshot sheets into the new workbook."""
    if not os.path.exists(src_filepath):
        return
    try:
        wb_src = load_workbook(src_filepath, data_only=True)
        copied = 0
        for sname in wb_src.sheetnames:
            if sname.startswith("VCART Systems -") and "LIVE" not in sname:
                if sname in dst_wb.sheetnames:
                    continue
                ws_src = wb_src[sname]
                ws_dst = dst_wb.create_sheet(sname)
                for row in ws_src.iter_rows():
                    for cell in row:
                        dst_cell = ws_dst.cell(row=cell.row, column=cell.column, value=cell.value)
                        if cell.has_style:
                            from copy import copy as _copy
                            dst_cell.font          = _copy(cell.font)
                            dst_cell.fill          = _copy(cell.fill)
                            dst_cell.alignment     = _copy(cell.alignment)
                            dst_cell.number_format = cell.number_format
                copied += 1
                print(f"        Copied sheet: {sname}")
        wb_src.close()
        if copied:
            print(f"        {copied} systems snapshot sheet(s) restored.")
    except Exception as e:
        print(f"        WARNING: could not copy systems sheets: {e}")


def main():
    t0 = datetime.now()
    print("=" * 60)
    print(f"  VCART VRHR Aggregator  —  {t0.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    print(f"\nReading territory files from:\n  {BASE_DIR}\n")
    all_td = load_all_territories(BASE_DIR)

    loaded = sum(1 for td in all_td.values() if td)
    total_campuses = sum(td["campus_count"] for td in all_td.values() if td)
    print(f"\n  {loaded}/{NUM_TERRITORIES} territories loaded  |  {total_campuses} total campuses")

    if loaded < NUM_TERRITORIES:
        failed = [l for l, td in all_td.items() if not td]
        print(f"\n  WARNING: Failed to load: {', '.join(failed)}")
        answer = input("\n  Continue anyway? (y/n): ").strip().lower()
        if answer != "y":
            print("\n  Aborted.")
            return

    snap_date     = datetime.now()
    current_month = snap_date.strftime("%B")

    print(f"\nChecking for existing output file...")
    if os.path.exists(OUTPUT):
        print(f"  Found: {os.path.basename(OUTPUT)}")
    else:
        print(f"  Not found — this is a first run.")

    existing_ts, live_month, existing_snaps, live_td = read_existing_totals(OUTPUT)

    save_snapshot = bool(live_month and live_month != current_month)
    if save_snapshot:
        print(f"\n  Month rollover detected: {live_month} → {current_month}")
        print(f"  Will save {live_month} as a snapshot before writing {current_month} data.")
    elif live_month:
        print(f"\n  Same month ({current_month}). Refreshing LIVE data only.")
    else:
        print(f"\n  First run. Writing fresh file.")

    print("\n" + "-" * 60)
    print("  Building workbook...")
    print("-" * 60)
    wb = Workbook()

    # Sheet 1: VCART Totals
    print("\n  [1/4] Writing VCART Totals sheet structure...")
    ws_totals = wb.active
    ws_totals.title = "VCART Totals"
    write_totals_sheet_headers(ws_totals)
    print("        Headers and column widths done.")

    # Restore time series
    if existing_ts:
        print(f"\n  [2/4] Restoring {len(existing_ts)} existing time series row(s)...")
        for i, row_data in enumerate(existing_ts):
            ts_row = TS_START_ROW + i
            if ts_row > TS_END_ROW:
                break
            row_label = str(row_data[0] or "").lower()
            if current_month.lower() in row_label or "live" in row_label:
                continue
            unmerge_row(ws_totals, ts_row)
            for c_idx, val in enumerate(row_data, 1):
                try:
                    cell = ws_totals.cell(row=ts_row, column=c_idx, value=val)
                    cell.fill = PatternFill(fill_type=None)
                    if c_idx > 2 and isinstance(val, float):
                        cell.number_format = "0%"
                except Exception:
                    pass
            print(f"        Restored: {row_data[0]}")
    else:
        print(f"\n  [2/4] No existing time series to restore.")

    print(f"\n  [2/4] Writing {current_month} to time series...")
    update_time_series(ws_totals, all_td, snap_date)
    total_barriers = sum(
        sum(td["totals"].get(c, 0) or 0 for c in BARRIER_COLS_OUT
            if isinstance(td["totals"].get(c), (int, float)))
        for td in all_td.values() if td
    )
    print(f"        {current_month} written — {total_barriers} total barriers, {total_campuses} campuses.")

    print(f"\n  [3/4] Writing LIVE section ({NUM_TERRITORIES} territories + NATION, raw + %)...")
    write_live_section(ws_totals, all_td, snap_date)
    print(f"        LIVE section done.")

    # Restore existing snapshots
    if existing_snaps:
        print(f"\n  [3/4] Restoring {len(existing_snaps)} existing snapshot(s)...")
        next_snap_row = get_snap_start(ws_totals)
        for idx, snap in enumerate(existing_snaps):
            label = str(snap[0][0]) if snap and snap[0][0] else f"snapshot {idx+1}"
            r = next_snap_row
            for row_data in snap:
                for c_idx, val in enumerate(row_data, 1):
                    try:
                        ws_totals.cell(row=r, column=c_idx, value=val)
                    except Exception:
                        pass
                r += 1
            next_snap_row = r + 1
            print(f"        Restored: {label}")
    else:
        print(f"\n  [3/4] No existing snapshots to restore.")

    # Save snapshot if month rolled
    if save_snapshot and live_td:
        print(f"\n  [3/4] Saving {live_month} snapshot...")
        append_snapshot(ws_totals, live_td, snap_date, live_month)
        print(f"        {live_month} snapshot written.")

    # Sheet 2: VCART Systems - LIVE
    print(f"\n  [4/4] Building VCART Systems - LIVE sheet (campus-level rows)...")
    build_systems_live_sheet(wb, all_td, snap_date)

    # Restore existing systems snapshot sheets from previous output
    print(f"\n  [4/4] Restoring existing systems snapshot sheets...")
    copy_existing_systems_sheets(OUTPUT, wb)

    # Create new systems snapshot sheet on month rollover
    if save_snapshot and live_month:
        print(f"\n  [4/4] Creating VCART Systems snapshot for {live_month}...")
        build_systems_snapshot_sheet(wb, all_td, snap_date, live_month)

    print(f"\n" + "-" * 60)
    print(f"  Saving → {OUTPUT}")
    print("-" * 60)
    wb.save(OUTPUT)

    elapsed = (datetime.now() - t0).seconds
    print(f"\n{'=' * 60}")
    print(f"  DONE in {elapsed}s")
    print(f"  {loaded}/{NUM_TERRITORIES} territories  |  {total_campuses} campuses  |  {total_barriers} barriers")
    print(f"  Output: {os.path.basename(OUTPUT)}")
    print(f"{'=' * 60}")
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        print(traceback.format_exc())
        input("\nPress Enter to exit...")