"""
All constants, file mappings, territory definitions, and column layouts
for the VCART Aggregator. Update this file when:
  - Territories change or new files are added
  - Column layouts shift in the source VRHR.VCART sheets
  - Section headers or output column structure changes

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
# TEMPORARY: pointed at the migrated file so a plain "python main.py" run
# writes there for review, instead of the real production file. Once
# you've checked it and swapped the migrated file into place as your real
# output, change this back to "VCART.Region.Nation.Totals.New.xlsx" (or
# just set VCART_OUTPUT as an env var instead of editing this line again).
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

# Category <-> subcategory taxonomy sheet — confirmed identical structure
# to ADFR's: first 6 columns headed with a category name (row 1), listing
# that category's subcategories below it.
TAXONOMY_SHEET_KEYWORD = "lists"

# Barrier booleans (6 cols): still read positionally, just from the
# Barrier Movement sheet instead of the main sheet now.
SRC_BARRIER_BOOL_START = 9   # col I
SRC_BARRIER_BOOL_END   = 14  # col N

# Barrier Tracking fields: read from the Barrier Movement sheet by HEADER
# TEXT (row 6), not position — that sheet's layout already shifted once
# (a Subcategory column inserted, "Prioritized Barrier 3" dropped), so a
# fixed position would silently break the next time it shifts again.
# {output_col: [candidate header strings, tried in order]}
# Days to Close (16, 21) are NOT header-matched here — the source header
# text "Days to close" is IDENTICAL for both Barrier 1 and Barrier 2, so it
# can't be told apart by text alone. Instead each is read one column to the
# right of its own paired Close Date column (confirmed in the real source:
# col 19 "Barrier 1 Close Date" -> col 20 "Days to close"; col 25 "Barrier 2
# Close Date " -> col 26 "Days to close") — see DAYS_TO_CLOSE_AFTER below.
BARRIER_TRACKING_FIELDS = {
    12: ["Prioritized Barrier 1 Category", "Prioritized Barrier 1"],
    13: ["Prioritized Barrier 1 Subcategory"],
    14: ["Barrier 1 Start Date"],
    15: ["Barrier 1 Close Date"],
    17: ["Prioritized Barrier 2 Category", "Prioritized Barrier 2"],
    18: ["Prioritized Barrier 2 Subcategory"],
    19: ["Barrier 2 Start Date"],
    20: ["Barrier 2 Close Date"],
    22: ["Start of Quarter Implementation Stage (ENTER AT START of QUARTER AND DO NOT CHANGE)",
         "Start of Quarter Implementation Stage"],
    23: ["Current Implementation Stage"],
}

# output_col (Days to Close) -> output_col (its paired Close Date column).
# read_territory_file resolves the *source* column of the close-date field
# first, then reads one column to its right for the days-to-close value —
# same pattern as ADFR.
DAYS_TO_CLOSE_AFTER = {16: 15, 21: 20}

# Confirmed against the real source files: the main sheet's data block
# genuinely only runs from "Centralized or Decentralized..." through
# "Last Update" — there is NO "CARE Team CEP Stage" field anywhere (the
# old positional config assumed one and silently read whatever happened
# to sit at that column offset instead). "Evolved" was also present in
# some territory files but not others — it's a hidden column not meant
# to be used at all, so it's deliberately excluded here rather than
# tried as a header candidate for any output column.
# {output_col: [candidate header strings, tried in order]}
TERRITORY_FIELDS = {
    24: ["Centralized or Decentralized Decision Making"],
    25: ["C/D Suite Access (Y/N)"],
    26: ["Current EHR"],
    27: ["ViiV 340b Contract (Y/N)"],
    28: ["Do they have campuses that are 340b eligibility? (Y/N)", "340b Eligibility"],
    29: ["Pathway mapped (Y/N)"],
    30: ["Pathway Differ by LAI?", "Pathway Differ by LAI"],
    31: ["Campus Preferred Pathway (SP or B&B)"],
    32: ["Specialty Pharmacy"],
    33: ["Assignment of Benefits (AOB)"],
    34: ["Buy and Bill"],
    35: ["ASOC"],
    36: ["ViiV Connect User Level"],
    37: ["ViiV Claims Portal User? (Y/N)", "ViiV Claims Portal User"],
    38: ["HIT Resource Opportunity for Reimbursement? (Y/N)", "HIT Resource Opportunity for Reimbursement"],
    39: ["HIT Resource Opportunity Used? (Y/N)", "HIT Resource Opportunity Used"],
    40: ["Recurring Meetings Scheduled with C/D Suite? (Y/N)", "Recurring Meetings Scheduled with C/D Suite"],
    41: ["What % of campuses have been flagged for TFRM support", "% campuses flagged for TFRM"],
    42: ["Territory FRM support launched? (Y/N)", "Territory FRM support launched"],
    43: ["Is Care offered for all eligible patients? (Y/N)", "Is Care offered for all eligible patients"],
    44: ["Last Update"],
}

# =============================================================================
# OUTPUT COLUMN LAYOUT
# Output sheet columns:
#   1 = Region (VCART East / VCART West / NATION)
#   2 = Territory Name (display label)
#   3 = Territory #
#   4 = VCART Name
#   5 = Total Accounts
#   6-11  = Implementation Barriers (Barrier Movement sheet, positional)
#   12-23 = Barrier Tracking, including Days to Close and Implementation
#           Stage (Barrier Movement sheet, header-text lookup)
#   24-44 = Territory fields — Centralized/Decentralized through Last
#           Update (main sheet, header-text lookup). "Evolved" is a
#           hidden column, deliberately excluded — not read at all.
# =============================================================================

OUT_LABEL_COL    = 1
OUT_TERRNAME_COL = 2
OUT_TERR_COL     = 3
OUT_NAME_COL     = 4
OUT_CAMPUS_COL   = 5
OUT_DATA_START   = 6   # output col for first data field (Access for All)
MAX_OUT_COL      = 44

# SRC_DATA_START is the first source column read on the MAIN sheet for
# demographic purposes and general layout math (e.g. the Systems sheet's
# column offset) — unrelated to the header-matched fields above, still
# positional since demographics (CID, Name, Address...) haven't moved.
# SRC_DATA_END is only used to detect a fully-blank row on the main sheet
# — real files vary in exact last column (29 or 30 confirmed across two
# real territory files), so this is set generously past either.
SRC_DATA_START = 9
SRC_DATA_END   = 30
NUM_DATA_COLS  = MAX_OUT_COL - OUT_DATA_START + 1

# Columns (output, 1-based) that hold numeric barrier counts → sum for NATION
BARRIER_COLS_OUT = list(range(6, 12))   # Access for All … Tech & Data Limitations (6 cols)

# Columns (output) that hold numeric values to nation-sum (all numeric data fields)
NATION_SUM_COLS = BARRIER_COLS_OUT      # extend if more numeric cols need summing

# Percent cols — numeric fields expressed as % of campuses.
# Output col 41 = "% campuses flagged for TFRM support".
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
# COLORS
# ViiV brand palette — muted navy and burgundy tones only, same values as
# ADFR. VIIV_DARK_COLORS marks which need white (vs. black) text overlaid
# for contrast; the rest need black text.
# =============================================================================

VIIV_NAVY_DARKEST = "1E3A5F"
VIIV_NAVY_MED     = "4E739E"
VIIV_NAVY_LIGHT   = "C5D2E3"
VIIV_MAROON_MED   = "A3505C"
VIIV_MAROON_DARK  = "8B3D4A"
VIIV_BLUSH_LIGHT  = "E8CDD1"
VIIV_DARK_COLORS  = {VIIV_NAVY_DARKEST, VIIV_NAVY_MED, VIIV_MAROON_MED, VIIV_MAROON_DARK}

# =============================================================================
# SECTION HEADERS  (output cols, label, hex fill color)
# =============================================================================

SECTION_HEADERS = [
    (6,  11, "Implementation Barriers",       VIIV_NAVY_MED),
    (12, 23, "Barrier Tracking",              VIIV_MAROON_DARK),   # Category/Subcategory/Start/Close/Days to Close x2, Stage x2
    (24, 28, "Discovery",                     VIIV_NAVY_LIGHT),    # Centralized/Decentralized through 340b Eligibility Campuses
    (29, 35, "Mapping the Pathway",           VIIV_MAROON_MED),    # starts at Pathway Mapped
    (36, 40, "Tech Embedment",                VIIV_BLUSH_LIGHT),   # starts at ViiV Connect User Level
    (41, 44, "Sustainment and Accountability",VIIV_NAVY_DARKEST),  # starts at % Campuses Flagged TFRM
]

# Time series section headers. The TS table's data now starts at the
# SAME column as the LIVE/snapshot sections (OUT_DATA_START) — cols 3-5
# are left as blank spacers, taking the place of the 3 identity columns
# (Territory #/VCART Name/Total Accounts) that only the LIVE/snapshot
# sections need. This used to be a -3 offset (ts_col = out_col - 3),
# which meant the SAME column letter held a DIFFERENT field depending on
# which table you were looking at — e.g. column N was "Barrier 1 Start
# Date" in the LIVE section but a completely different field in the TS
# table. Aligning them removes that whole class of confusion — a column
# means the same thing everywhere on this sheet now.
TS_SECTION_HEADERS = SECTION_HEADERS

# LIVE section header colors (cols 1-5, 6-11 [Implementation Barriers],
# 12-23 [Barrier Tracking], rest via _section_fill). These must be kept in
# sync with SECTION_HEADERS — previously this used a separate hardcoded
# navy zone (cols 20-23) for a "Customer Engagement Phase" grouping that
# turned out not to correspond to anything real in the source file; Stage
# fields now live inside Barrier Tracking where they're actually tracked,
# so that zone is gone.
LIVE_HDR_PEACH  = VIIV_NAVY_LIGHT   # cols 1-5: Region, Terr Name, Terr#, VCART Name, Total Accounts
LIVE_HDR_GREEN  = VIIV_NAVY_MED     # cols 6-11: Implementation Barriers
LIVE_HDR_PINK   = VIIV_MAROON_DARK  # cols 12-23: Barrier Tracking

# =============================================================================
# COLUMN HEADERS  (output col → header text)
# =============================================================================

COL_HEADERS = {
    1:  "Region",
    2:  "Territory\nName",
    3:  "Territory #",
    4:  "VCART Name",
    5:  "Total\nAccounts",
    6:  "Access for All /\nCoverage\nChallenges",
    7:  "Financial &\nReimbursement",
    8:  "Knowledge &\nTraining\nDeficits",
    9:  "Operational &\nInfrastructure\nGaps",
    10: "Stakeholder\nMisalignment",
    11: "Technology &\nData\nLimitations",
    12: "Barrier 1\nCategory",
    13: "Barrier 1\nSubcategory",
    14: "Barrier 1\nStart Date",
    15: "Barrier 1\nClose Date",
    16: "Barrier 1\nDays to Close",
    17: "Barrier 2\nCategory",
    18: "Barrier 2\nSubcategory",
    19: "Barrier 2\nStart Date",
    20: "Barrier 2\nClose Date",
    21: "Barrier 2\nDays to Close",
    22: "Start of Qtr\nImpl. Stage",
    23: "Current\nImpl. Stage",
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