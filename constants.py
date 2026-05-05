"""
constants.py — Project-wide file paths and spreadsheet row constants.

All other modules import from here so row numbers and filenames are
defined in exactly one place.
"""

from pathlib import Path

# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------

INPUT_FILE  = Path("Scoreboard Test.xlsx")
OUTPUT_FILE = Path("output.json")

# ---------------------------------------------------------------------------
# Row numbers (1-based, matching what you see in Excel)
# ---------------------------------------------------------------------------

ROW_SECTION_HEADER = 1   # merged-cell section banners (e.g. "PHONE PERFORMANCE")
ROW_METRIC_LABELS  = 2   # main column headers
ROW_FOCUS          = 3   # "Financial", "Marketing", …
ROW_SOURCE         = 4   # "EMR", "CallHero", …
ROW_ROLE           = 5   # "J", "Nicole", "Beth", …
ROW_TARGET_FLAGS   = 6   # "Target" marker row (sparse; mostly empty)
ROW_TARGET_VALUES  = 7   # spot target values / formula notes
FIRST_DATA_ROW     = 8   # weekly data records start here
