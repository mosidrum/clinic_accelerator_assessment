"""
parser.py — Spreadsheet structure parsers.

Responsible for reading the SCOREBOARD worksheet and producing two
artefacts:
  1. col_map  — a dict of column index → metric descriptor, built from
                the header rows (rows 1–7).
  2. records  — a list of flat dicts, one per weekly data row (row 8+).

All knowledge of which row contains what lives in constants.py.
All string/value normalisation lives in utils.py.
"""

import datetime

from constants import (
    ROW_SECTION_HEADER,
    ROW_METRIC_LABELS,
    ROW_FOCUS,
    ROW_SOURCE,
    ROW_ROLE,
    ROW_TARGET_VALUES,
    FIRST_DATA_ROW,
)
from utils import slugify, make_unique_key, coerce_value, cell_val


def parse_section_headers(ws) -> dict[int, str]:
    """Map column index → section name from the merged-cell banner row (row 1).

    The spreadsheet uses merged cells to label groups of columns with a
    section title (e.g. "PHONE PERFORMANCE" spanning AJ:AP).  We expand
    each merged range so every covered column gets the banner text.
    Un-merged values in row 1 are also captured.

    Returns:
        A dict keyed by column index whose value is the section name string.
        Columns with no banner entry are simply absent from the dict.
    """
    section_map: dict[int, str] = {}

    for merge in ws.merged_cells.ranges:
        if merge.min_row == ROW_SECTION_HEADER:
            # Only the top-left cell of a merged range holds data in openpyxl
            value = ws.cell(merge.min_row, merge.min_col).value
            if value:
                for col in range(merge.min_col, merge.max_col + 1):
                    section_map[col] = str(value).strip()

    # Pick up any unmerged values in row 1 not already covered
    for cell in ws[ROW_SECTION_HEADER]:
        if cell.value and cell.column not in section_map:
            section_map[cell.column] = str(cell.value).strip()

    return section_map


def build_column_map(ws, section_map: dict[int, str]) -> dict[int, dict]:
    """Build a descriptor dict for every metric column from the header rows.

    Iterates row 2 (the main header row) left-to-right.  For each cell
    that has a non-empty label it reads the corresponding metadata from
    rows 3–7 and produces a descriptor dict.

    Skips:
        - Column A (index 1) — that is the date column, handled in extract_records.
        - Columns with a blank or whitespace-only header — layout spacers.

    Duplicate slugs (same label used for multiple service lines) are
    disambiguated by appending the Excel column letter.

    Returns:
        A dict keyed by column index.  Each value is::

            {
                "key":     str,   # unique JSON-safe slug
                "label":   str,   # original header text, preserved exactly
                "column":  str,   # Excel column letter e.g. "B", "AJ"
                "section": str|None,
                "focus":   str|None,
                "source":  str|None,
                "role":    str|None,
                "target":  any,
            }
    """
    seen_keys: dict[str, str] = {}  # base_key → first column letter assigned
    columns: dict[int, dict] = {}

    for cell in ws[ROW_METRIC_LABELS]:
        col_idx = cell.column

        if col_idx == 1 or cell.value is None:
            continue

        raw_label = str(cell.value).strip().replace("\n", " ")
        base_key  = slugify(raw_label)

        if not base_key:
            continue  # blank header after normalisation → spacer column

        unique_key = make_unique_key(base_key, cell.column_letter, seen_keys)

        columns[col_idx] = {
            "key":     unique_key,
            "label":   raw_label,
            "column":  cell.column_letter,
            "section": section_map.get(col_idx),
            "focus":   cell_val(ws, ROW_FOCUS,        col_idx),
            "source":  cell_val(ws, ROW_SOURCE,       col_idx),
            "role":    cell_val(ws, ROW_ROLE,          col_idx),
            "target":  cell_val(ws, ROW_TARGET_VALUES, col_idx),
        }

    return columns


def extract_records(ws, col_map: dict[int, dict]) -> list[dict]:
    """Extract each weekly data row as a flat dict keyed by metric slug.

    Iterates from FIRST_DATA_ROW downward.  A row is included only if
    column A contains a datetime object — this skips stray non-date rows
    near the header block (e.g. a row with 6.0 in column A).

    Every metric in col_map gets an entry in each record; cells that are
    empty or formula-errored become None.

    Returns:
        A list of record dicts sorted chronologically (oldest first).
        Callers may re-sort with the functions in sorters.py.
    """
    records = []

    for row in ws.iter_rows(min_row=FIRST_DATA_ROW):
        raw_date = row[0].value  # column A is index 0 in the row tuple

        if not isinstance(raw_date, datetime.datetime):
            continue

        record: dict = {"week_ending": raw_date.date().isoformat()}

        for col_idx, meta in col_map.items():
            cell = ws.cell(row=row[0].row, column=col_idx)
            record[meta["key"]] = coerce_value(cell.value)

        records.append(record)

    records.sort(key=lambda r: r["week_ending"])
    return records
