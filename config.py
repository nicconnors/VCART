"""
VCART VRHR Configuration
========================
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

# Root folder containing all per-territory VCART VRHR xlsx files
BASE_DIR = r"C:\Users\nc139629\GSK\ViiV Field Reimbursement Managers - Documents\General\(VRHR) ViiV Reimbursement Health Report\2026\VCART VRHR"

OUTPUT = r"C:\Users\nc139629\GSK\ViiV Field Reimbursement Managers - Documents\General\(VRHR) ViiV Reimbursement Health Report\2026\VCART.Region.Nation.Totals.New.xlsx"

# =============================================================================
# SOURCE FILES
# One entry per territory: (display_label, filename, territory_number_override)
# territory_number_override: use None to read col A from file; set a string to
# hard-code (for files missing col A like MidAtlantic, MidSouth).
# =============================================================================

TERRITORIES = [
    # (display_label,             filename,                                          terr_override)
    ("VCART: National",           "National_VCART_VRHR_2026.xlsx",                  "VREJ01"),
    ("VCART: Northeast",          "NortheastVCART_VRHR_20261.xlsx",                  "VREJ02"),
    ("VCART: NYC North",          "NYCNorth_VCART_VRHR_202611_-_newest__1_.xlsx",    "VREJ03"),
    ("VCART: NYC South",          "NYCSouth_VCART_VRHR_20262.xlsx",                  "VCNC05"),
    ("VCART: Mid Atlantic",       "MidAtlantic_VCART_VRHR_20263.xlsx",               None),
    ("VCART: SoCal",              "SoCal_VCART_VRHR_20264.xlsx",                     "VREK02"),
    ("VCART: Central North",      "CentralNorth_VCART_VRHR_20265.xlsx",              "VREK03"),
    ("VCART: Texas",              "Texas_VCART_VRHR_20266.xlsx",                     "VREK04"),
    ("VCART: Mid South",          "MidSouth_VCART_VRHR_20267.xlsx",                  None),
    ("VCART: Great Lakes",        "GreatLakes_VCART_VRHR_20268.xlsx",                "VREK05"),
    ("VCART: Ohio Valley",        "OhioValley_VCART_VRHR_20269.xlsx",                "VREK06"),
    ("VCART: New England",        "NewEngland_VCART_VRHR_202610.xlsx",               "VREJ08"),
    ("VCART: Carolinas",          "Carolinas_VCART_VRHR_202631.xlsx",                "VREJ06"),
    ("VCART: Southeast",          "SouthEast_VCART_VRHR_202632.xlsx",                "VREJ07"),
    ("VCART: Central West",       "CentralWest_VCART_VRHR_202633.xlsx",              "VREK01"),
    ("VCART: PacNW",              "PacNW_VCART_VRHR_202661.xlsx",                    "VREK07"),
]

# =============================================================================
# SOURCE SHEET & LAYOUT
# =============================================================================

SOURCE_SHEET_KEYWORDS = ("vrhr", "vcart")   # sheet name must contain both (case-insensitive)
HEADER_ROW  = 6   # 1-based row with column headers in source files
DATA_START  = 7   # first data row in source files

# Source columns (1-based) that contain the 36 output data values
# Cols A-H (1-8): demographics (terr #, name, CID, corp name, addr, city, state, zip)
# Col I-N  (9-14): 6 barrier booleans
# Col O-R  (15-18): Q1-Q4 prioritized barrier text
# Col S-AP (19-44): remaining data fields
# We read cols 9-44 as data (36 cols) → output cols 7-42 in aggregator
# Col 1=Region/label, col 2=Territory#, col 3=VCART Name, col 4=Total Campuses,
# then cols 5-40 = source cols 9-44

NUM_DATA_COLS = 36   # source cols I through AP (indices 8..43)
SRC_DATA_START = 9  # 1-based source column where data begins (col I)
SRC_DATA_END   = 44 # 1-based source column where data ends (col AR)  inclusive

# =============================================================================
# OUTPUT COLUMN LAYOUT
# Output sheet columns:
#   1 = Region (VCART East / VCART West / NATION)
#   2 = Territory Name (display label)
#   3 = Territory #
#   4 = VCART Name
#   5 = Total Accounts
#   6..41 = data cols (maps from source cols 9..44)
# =============================================================================

OUT_LABEL_COL   = 1
OUT_TERRNAME_COL = 2
OUT_TERR_COL    = 3
OUT_NAME_COL    = 4
OUT_CAMPUS_COL  = 5
OUT_DATA_START  = 6   # output col for first data field (Access for All)
MAX_OUT_COL     = OUT_DATA_START + NUM_DATA_COLS - 1  # = 41

# Columns (output, 1-based) that hold numeric barrier counts → sum for NATION
BARRIER_COLS_OUT = list(range(6, 12))   # Access for All … Tech & Data Limitations (6 cols)

# Columns (output) that hold numeric values to nation-sum (all numeric data fields)
NATION_SUM_COLS = BARRIER_COLS_OUT      # extend if more numeric cols need summing

# Percent cols — numeric fields expressed as % of campuses (e.g. %flagged for TFRM)
# Output col 41 = "What % of campuses flagged for TFRM support" (src col 44 → out col 41)
PCT_COLS_OUT = [41]

# =============================================================================
# REGION GROUPINGS
# Maps territory labels to their region display name
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
    (6,  11, "Implementation Barriers",      "92D050"),   # green
    (12, 15, "Q Prioritized Barriers",        "92D050"),   # green (extends through Q4)
    (16, 19, "Customer Engagement Phase",     "2E5F8A"),   # light navy blue
    (20, 24, "Discovery",                     "B3D9FF"),
    (25, 31, "Mapping the Pathway",           "92D050"),
    (32, 36, "Tech Embedment",                "B3D9FF"),
    (37, 41, "Sustainment and Accountability","B3D9FF"),
]

# Time series section headers — only barrier cols, no identity cols
TS_SECTION_HEADERS = [
    (3,  8,  "Implementation Barriers",      "92D050"),
    (9,  12, "Q Prioritized Barriers",        "92D050"),
    (13, 15, "Customer Engagement Phase",     "2E5F8A"),
    (16, 20, "Discovery",                     "B3D9FF"),
    (21, 27, "Mapping the Pathway",           "92D050"),
    (28, 32, "Tech Embedment",                "B3D9FF"),
    (33, 37, "Sustainment and Accountability","B3D9FF"),
]

# LIVE section header colors (cols 1-5 peach, 6-15 green, 16-19 navy, rest normal)
LIVE_HDR_PEACH  = "FFE4C4"   # cols 1-5: Region, Terr Name, Terr#, VCART Name, Total Accounts
LIVE_HDR_GREEN  = "92D050"   # cols 6-15: Access for All → Q4 Prioritized Barrier
LIVE_HDR_NAVY   = "2E5F8A"   # cols 16-19: Customer Engagement Phase cols

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
    12: "Q1 Prioritized\nBarrier",
    13: "Q2 Prioritized\nBarrier",
    14: "Q3 Prioritized\nBarrier",
    15: "Q4 Prioritized\nBarrier",
    16: "Prioritized\nImpl. Barrier",
    17: "Current\nImpl. Stage",
    18: "Start of Qtr\nImpl. Stage",
    19: "CARE Team\nCEP Stage",
    20: "Evolved\n(DO NOT\nUPDATE)",
    21: "Centralized\nor\nDecentralized",
    22: "C/D Suite\nAccess",
    23: "Current\nEHR",
    24: "ViiV 340b\nContract",
    25: "340b\nEligibility\nCampuses",
    26: "Pathway\nMapped",
    27: "Pathway\nDiffer\nby LAI",
    28: "Campus\nPreferred\nPathway",
    29: "Specialty\nPharmacy",
    30: "AOB",
    31: "Buy and\nBill",
    32: "ASOC",
    33: "ViiV Connect\nUser Level",
    34: "ViiV Claims\nPortal User",
    35: "HIT Resource\nOpportunity",
    36: "HIT Resource\nUsed",
    37: "Recurring\nMeetings\nC/D Suite",
    38: "% Campuses\nFlagged\nTFRM",
    39: "Territory\nFRM\nLaunched",
    40: "Care\nOffered\nAll Patients",
    41: "Last\nUpdate",
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

_validate()