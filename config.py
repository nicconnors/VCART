"""
VCART VRHR Configuration
========================
All constants, file mappings, territory definitions, and column layouts
for the VCART Aggregator. Update this file when:
  - Territories change or new files are added
  - Column layouts shift in the source VRHR.VCART sheets
  - Section headers or output column structure changes

CHANGE LOG (this revision):
  - As of Q3 2026, barrier-related data (booleans + tracking) moved off the
    main "VRHR.VCART" sheet onto a separate "<Qn> Barrier Movement" sheet,
    with a restructured layout: a Subcategory column was inserted after
    each barrier category, and the old "Q1-Q4 Prioritized Barrier" +
    "Prioritized Implementation Barrier" fields were replaced by two
    barriers (Category/Subcategory/Start/Close each) + implementation
    stage fields. Same restructuring ADFR went through — see that
    project's config.py CHANGE LOG for the full story.
  - The main "VRHR.VCART" sheet's own barrier data is now STALE — the
    Barrier Movement sheet is the current/authoritative copy. So barrier
    booleans (Access for All -> Tech & Data Limitations), which used to be
    read from the main sheet, are now read from the Barrier Movement sheet
    too, for BOTH the territory-level totals row and every individual
    campus row in the "VCART Systems - LIVE" sheet.
  - Barrier Tracking fields (cols 12-21) are read from the Barrier
    Movement sheet by HEADER TEXT, not fixed position, since that sheet's
    layout has already shifted once. Everything from "CARE Team CEP
    Stage" onward is unaffected and still read positionally from the main
    sheet.
  - Net effect: the barrier block grew from 13 output columns (6-18) to
    16 (6-21), so everything after it shifted from cols 19-41 to 22-44.
    MAX_OUT_COL is now 44 (was 41).
  - Fixed in passing (unavoidable side effect of the column shift, not
    extra scope): PCT_COLS_OUT previously pointed at col 41 ("Last
    Update") instead of the real "% campuses flagged for TFRM" field —
    after the shift, col 41 correctly IS that field. TS_SECTION_HEADERS
    is now DERIVED from SECTION_HEADERS instead of hand-typed as a
    separate parallel list, which is what caused the original off-by-one
    in the first place — deriving it removes that whole bug class going
    forward, not just today's instance.
"""

import os

# =============================================================================
# FILE PATHS
# =============================================================================
# NOTE: paths are built from the current user's home directory so this script
# runs unmodified on any teammate's machine, as long as their OneDrive sync
# creates the same "GSK\ViiV Field Reimbursement Managers - 2026\..."
# folder structure under their own Windows profile.
#
# If your OneDrive sync root is NOT your home dir (e.g. it's redirected, or
# your OneDrive folder has a different name), override BASE_DIR / OUTPUT via
# environment variables instead of editing this file — see VCART_BASE_DIR /
# VCART_OUTPUT below.

_SHAREPOINT_SUBPATH = os.path.join(
    "GSK",
    "ViiV Field Reimbursement Managers - 2026",
)

_DEFAULT_BASE_DIR = os.path.join(_SHAREPOINT_SUBPATH, "VCART VRHR")
_DEFAULT_OUTPUT   = os.path.join(_SHAREPOINT_SUBPATH, "VCART.Region.Nation.Totals.New.xlsx")

HOME = os.path.expanduser("~")

# Root folder containing all per-territory VCART VRHR xlsx files
BASE_DIR = os.environ.get("VCART_BASE_DIR", os.path.join(HOME, _DEFAULT_BASE_DIR))

OUTPUT = os.environ.get("VCART_OUTPUT", os.path.join(HOME, _DEFAULT_OUTPUT))

if not os.path.isdir(BASE_DIR):
    raise FileNotFoundError(
        f"CONFIG ERROR: BASE_DIR not found:\n  {BASE_DIR}\n\n"
        "This usually means your OneDrive sync folder isn't named 'GSK' under "
        "your home directory, or the 'VCART VRHR' subfolder isn't where "
        "expected. Fix by setting an environment variable instead of editing "
        "config.py, e.g. in PowerShell:\n\n"
        '  $env:VCART_BASE_DIR = "C:\\path\\to\\your\\VCART VRHR"\n'
        '  $env:VCART_OUTPUT   = "C:\\path\\to\\your\\output.xlsx"\n'
    )

# =============================================================================
# SOURCE FILES
# =============================================================================

TERRITORIES = [
    # (display_label,             filename)
    ("VCART: National",           "National.VCART.VRHR.2026.xlsx"),
    ("VCART: Northeast",          "NortheastVCART.VRHR.20261.xlsx"),
    ("VCART: NYC North",          "NYCNorth.VCART.VRHR.202611.xlsx"),
    ("VCART: NYC South",          "NYCSouth.VCART.VRHR.20262.xlsx"),
    ("VCART: Mid Atlantic",       "MidAtlantic.VCART.VRHR.20263.xlsx"),
    ("VCART: SoCal",              "SoCal.VCART.VRHR.20264.xlsx"),
    ("VCART: Central North",      "CentralNorth.VCART.VRHR.20265.xlsx"),
    ("VCART: Texas",              "Texas.VCART.VRHR.20266.xlsx"),
    ("VCART: Mid South",          "MidSouth.VCART.VRHR.20267.xlsx"),
    ("VCART: Great Lakes",        "GreatLakes.VCART.VRHR.20268.xlsx"),
    ("VCART: Ohio Valley",        "OhioValley.VCART.VRHR.20269.xlsx"),
    ("VCART: New England",        "NewEngland.VCART.VRHR.202610.xlsx"),
    ("VCART: Carolinas",          "Carolinas.VCART.VRHR.202631.xlsx"),
    ("VCART: Southeast",          "SouthEast.VCART.VRHR.202632.xlsx"),
    ("VCART: Central West",       "CentralWest.VCART.VRHR.202633.xlsx"),
    ("VCART: PacNW",              "PacNW.VCART.VRHR.202661.xlsx"),
]

# =============================================================================
# SOURCE SHEET & LAYOUT
# =============================================================================

SOURCE_SHEET_KEYWORDS = ("vrhr", "vcart")   # main sheet name must contain both (case-insensitive)
HEADER_ROW  = 6   # 1-based row with column headers in source files
DATA_START  = 7   # first data row in source files

# NEW as of Q3 2026: barrier boolean + tracking data now lives on a
# separate sheet (e.g. "Q3 Barrier Movement"). Matched by keyword, not an
# exact name, so this keeps working automatically as the quarter changes.
BARRIER_MOVEMENT_SHEET_KEYWORD = "barrier movement"

# Barrier booleans (6 cols): still read positionally, just from the
# Barrier Movement sheet instead of the main sheet now.
SRC_BARRIER_BOOL_START = 9   # col I
SRC_BARRIER_BOOL_END   = 14  # col N

# Barrier Tracking fields: read from the Barrier Movement sheet by HEADER
# TEXT (row 6), not position — that sheet's layout already shifted once
# (a Subcategory column inserted, "Prioritized Barrier 3" dropped), so a
# fixed position would silently break the next time it shifts again.
# {output_col: [candidate header strings, tried in order]}
BARRIER_TRACKING_FIELDS = {
    12: ["Prioritized Barrier 1 Category", "Prioritized Barrier 1"],
    13: ["Prioritized Barrier 1 Subcategory"],
    14: ["Barrier 1 Start Date"],
    15: ["Barrier 1 Close Date"],
    16: ["Prioritized Barrier 2 Category", "Prioritized Barrier 2"],
    17: ["Prioritized Barrier 2 Subcategory"],
    18: ["Barrier 2 Start Date"],
    19: ["Barrier 2 Close Date"],
    20: ["Start of Quarter Implementation Stage (ENTER AT START of QUARTER AND DO NOT CHANGE)",
         "Start of Quarter Implementation Stage"],
    21: ["Current Implementation Stage"],
}

# Everything from "CARE Team CEP Stage" onward is unaffected by the
# restructure and still read positionally from the MAIN sheet.
# POST_BARRIER_SRC_START = the main sheet's own source column for "CARE
# Team CEP Stage" (unchanged — that sheet's layout didn't move).
# POST_BARRIER_OUT_START = the output column it now lands at.
POST_BARRIER_SRC_START = 22
POST_BARRIER_SRC_END   = 44   # main sheet's "Last Update" column
POST_BARRIER_OUT_START = 22

# =============================================================================
# OUTPUT COLUMN LAYOUT
# Output sheet columns:
#   1 = Region (VCART East / VCART West / NATION)
#   2 = Territory Name (display label)
#   3 = Territory #
#   4 = VCART Name
#   5 = Total Accounts
#   6-11  = Implementation Barriers (Barrier Movement sheet, positional)
#   12-21 = Barrier Tracking (Barrier Movement sheet, header-text lookup)
#   22-44 = CEP onward (main sheet, positional)
# =============================================================================

OUT_LABEL_COL    = 1
OUT_TERRNAME_COL = 2
OUT_TERR_COL     = 3
OUT_NAME_COL     = 4
OUT_CAMPUS_COL   = 5
OUT_DATA_START   = 6   # output col for first data field (Access for All)
MAX_OUT_COL      = POST_BARRIER_OUT_START + (POST_BARRIER_SRC_END - POST_BARRIER_SRC_START)  # 44

# Kept for backward compatibility with anything reading NUM_DATA_COLS /
# SRC_DATA_START / SRC_DATA_END as "the whole data block" — SRC_DATA_START
# is still the first source column read on the MAIN sheet for demographic
# purposes and general layout math (e.g. the Systems sheet's column
# offset), which hasn't changed.
SRC_DATA_START = 9
SRC_DATA_END   = POST_BARRIER_SRC_END  # 44
NUM_DATA_COLS  = MAX_OUT_COL - OUT_DATA_START + 1  # 39

# Columns (output, 1-based) that hold numeric barrier counts → sum for NATION
BARRIER_COLS_OUT = list(range(6, 12))   # Access for All … Tech & Data Limitations (6 cols)

# Columns (output) that hold numeric values to nation-sum (all numeric data fields)
NATION_SUM_COLS = BARRIER_COLS_OUT      # extend if more numeric cols need summing

# Percent cols — numeric fields expressed as % of campuses.
# Output col 41 = "% campuses flagged for TFRM support" (main sheet src 41).
PCT_COLS_OUT = [41]

# =============================================================================
# REGION GROUPINGS
# =============================================================================

REGION_MAP = {
    "VCART: National":      "VCART - East",
    "VCART: Northeast":     "VCART - East",
    "VCART: NYC North":     "VCART - East",
    "VCART: NYC South":     "VCART - East",
    "VCART: Mid Atlantic":  "VCART - East",
    "VCART: Carolinas":     "VCART - East",
    "VCART: Southeast":     "VCART - East",
    "VCART: New England":   "VCART - East",
    "VCART: Mid South":     "VCART - East",
    "VCART: SoCal":         "VCART - West",
    "VCART: Central North": "VCART - West",
    "VCART: Texas":         "VCART - West",
    "VCART: Great Lakes":   "VCART - West",
    "VCART: Ohio Valley":   "VCART - West",
    "VCART: Central West":  "VCART - West",
    "VCART: PacNW":         "VCART - West",
}

# =============================================================================
# SECTION HEADERS  (output cols, label, hex fill color)
# =============================================================================

SECTION_HEADERS = [
    (6,  11, "Implementation Barriers",       "92D050"),   # green
    (12, 19, "Barrier Tracking",              "FFB3C6"),   # pink — Barrier1/2 Category/Subcategory/Start/Close only
    (20, 23, "Customer Engagement Phase",     "2E5F8A"),   # navy — Start/Current Impl Stage, CARE Team CEP Stage, Evolved
    (24, 28, "Discovery",                     "B3D9FF"),   # ends at 340b Eligibility Campuses
    (29, 35, "Mapping the Pathway",           "92D050"),   # starts at Pathway Mapped
    (36, 40, "Tech Embedment",                "B3D9FF"),   # starts at ViiV Connect User Level
    (41, 44, "Sustainment and Accountability","B3D9FF"),   # starts at % Campuses Flagged TFRM
]

# Time series section headers — DERIVED from SECTION_HEADERS, not hand-typed.
# The TS row skips the 5 identity cols (Region/TerrName/Terr#/Name/Campuses)
# that the main Totals sheet has, so its columns run 3 narrower than the
# output layout: ts_col = out_col - OUT_DATA_START + 3 = out_col - 3.
# Hand-typing this as a second parallel list (the old approach) is exactly
# what let it drift out of sync and go off-by-one in the first place.
_TS_COL_OFFSET = OUT_DATA_START - 3  # = 3
TS_SECTION_HEADERS = [
    (start - _TS_COL_OFFSET, end - _TS_COL_OFFSET, label, color)
    for start, end, label, color in SECTION_HEADERS
]

# LIVE section header colors (cols 1-5 peach, 6-11 green [Implementation
# Barriers], 12-19 pink [Barrier Tracking], 20-23 navy [CEP], rest normal
# via _section_fill). These must be kept in sync with SECTION_HEADERS —
# previously this used one hardcoded green zone spanning both
# Implementation Barriers AND Barrier Tracking, which is what caused the
# column-header row (row 13 in the LIVE section) to show green under
# "Barrier Tracking" instead of matching that section's actual pink.
LIVE_HDR_PEACH  = "FFE4C4"   # cols 1-5: Region, Terr Name, Terr#, VCART Name, Total Accounts
LIVE_HDR_GREEN  = "92D050"   # cols 6-11: Implementation Barriers
LIVE_HDR_PINK   = "FFB3C6"   # cols 12-19: Barrier Tracking
LIVE_HDR_NAVY   = "2E5F8A"   # cols 20-23: Customer Engagement Phase

# =============================================================================
# COLUMN HEADERS  (output col → header text)
# =============================================================================

COL_HEADERS = {
    1:  "Region",
    2:  "Territory\nName",
    3:  "Territory #",
    4:  "VCART Name",
    5:  "Total\nAccounts",
    6:  "Access\nfor All",
    7:  "Financial &\nReimb.\nChallenges",
    8:  "Knowledge &\nTraining\nDeficits",
    9:  "Operational &\nInfrastructure\nGaps",
    10: "Stakeholder\nMisalignment",
    11: "Technology &\nData\nLimitations",
    12: "Barrier 1\nCategory",
    13: "Barrier 1\nSubcategory",
    14: "Barrier 1\nStart Date",
    15: "Barrier 1\nClose Date",
    16: "Barrier 2\nCategory",
    17: "Barrier 2\nSubcategory",
    18: "Barrier 2\nStart Date",
    19: "Barrier 2\nClose Date",
    20: "Start of Qtr\nImpl. Stage",
    21: "Current\nImpl. Stage",
    22: "CARE Team\nCEP Stage",
    23: "Evolved\n(DO NOT\nUPDATE)",
    24: "Centralized\nor\nDecentralized",
    25: "C/D Suite\nAccess",
    26: "Current\nEHR",
    27: "ViiV 340b\nContract",
    28: "340b\nEligibility\nCampuses",
    29: "Pathway\nMapped",
    30: "Pathway\nDiffer\nby LAI",
    31: "Campus\nPreferred\nPathway",
    32: "Specialty\nPharmacy",
    33: "AOB",
    34: "Buy and\nBill",
    35: "ASOC",
    36: "ViiV Connect\nUser Level",
    37: "ViiV Claims\nPortal User",
    38: "HIT Resource\nOpportunity",
    39: "HIT Resource\nUsed",
    40: "Recurring\nMeetings\nC/D Suite",
    41: "% Campuses\nFlagged\nTFRM",
    42: "Territory\nFRM\nLaunched",
    43: "Care\nOffered\nAll Patients",
    44: "Last\nUpdate",
}

# =============================================================================
# COLORS & STYLING
# =============================================================================

YELLOW      = "FFFF00"
NATION_FILL = "D9D9D9"
EAST_FILL   = "FFFFFF"   # white — all data rows white
WEST_FILL   = "FFFFFF"   # white — all data rows white

# =============================================================================
# TIME SERIES (same pattern as VRHR)
# =============================================================================

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

TS_HEADER_ROW = 3
TS_START_ROW  = TS_HEADER_ROW + 1   # 4
TS_END_ROW    = TS_START_ROW + 11   # 15

# =============================================================================
# VALIDATION
# =============================================================================

def _validate():
    labels = [t[0] for t in TERRITORIES]
    missing = [l for l in labels if l not in REGION_MAP]
    if missing:
        raise ValueError(f"CONFIG ERROR: territories missing from REGION_MAP: {missing}")

    for hdr_list in (SECTION_HEADERS, TS_SECTION_HEADERS):
        ranges = [(s, e) for s, e, *_ in hdr_list]
        for i, (s1, e1) in enumerate(ranges):
            for s2, e2 in ranges[i + 1:]:
                if s1 <= e2 and s2 <= e1:
                    raise ValueError(
                        f"CONFIG ERROR: section headers overlap: ({s1}-{e1}) and ({s2}-{e2})"
                    )

    header_cols = set(COL_HEADERS.keys())
    expected = set(range(1, MAX_OUT_COL + 1))
    if header_cols != expected:
        missing_cols = expected - header_cols
        extra_cols = header_cols - expected
        raise ValueError(
            f"CONFIG ERROR: COL_HEADERS should cover cols 1-{MAX_OUT_COL} exactly. "
            f"Missing: {sorted(missing_cols)}  Extra: {sorted(extra_cols)}"
        )

_validate()