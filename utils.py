"""
utils.py — Pure helper functions with no project dependencies.

These functions know nothing about spreadsheet structure or file paths;
they only transform individual values or strings.
"""

import re
import datetime


def slugify(text: str) -> str:
    """Convert a header label to a safe, lowercase JSON key.

    Replaces % and # with readable words first so that "NAR % Collected"
    and "NAR Collected" don't collapse to the same slug. Then collapses
    runs of non-alphanumeric characters to a single underscore.

    Examples:
        "Total Revenue \\n- All Services"  →  "total_revenue_all_services"
        "NAR % Collected at Ax"            →  "nar_pct_collected_at_ax"
        "# of Rx (Clinic Wide)"            →  "num_of_rx_clinic_wide"
    """
    text = str(text).strip().replace("\n", " ")
    text = text.replace("%", "_pct").replace("#", "_num")
    text = re.sub(r"[^a-zA-Z0-9]+", "_", text)
    return text.strip("_").lower()


def make_unique_key(base_key: str, col_letter: str, seen: dict[str, str]) -> str:
    """Return a collision-free JSON key, appending the column letter when needed.

    Using the column letter (e.g. "pva_4_wk_avg_df") is more traceable than
    a bare counter ("pva_4_wk_avg_2") — a developer can cross-reference the
    suffix directly against the original spreadsheet or metadata["metrics"].

    Args:
        base_key:   The slug derived from the column header.
        col_letter: The Excel column letter for this cell (e.g. "DF").
        seen:       Mutable dict tracking which base keys have been assigned.

    Examples:
        "pva_4_wk_avg" col CR → "pva_4_wk_avg"     (first occurrence)
        "pva_4_wk_avg" col DF → "pva_4_wk_avg_df"  (collision → append letter)
    """
    if base_key not in seen:
        seen[base_key] = col_letter
        return base_key
    return f"{base_key}_{col_letter.lower()}"


def coerce_value(raw):
    """Turn a raw openpyxl cell value into a clean Python type.

    Conversion rules (applied in order):
        None                  → None
        datetime.datetime     → ISO date string  e.g. "2026-02-09"
        int / float           → kept as-is
        numeric string        → float            e.g. "2.87" → 2.87
        formula error string  → None             e.g. "#REF!"
        empty string          → None
        any other string      → stripped string
    """
    if raw is None:
        return None

    if isinstance(raw, datetime.datetime):
        return raw.date().isoformat()

    if isinstance(raw, (int, float)):
        return raw

    text = str(raw).strip()

    if not text:
        return None

    # Formula errors produced by openpyxl when data_only=True (e.g. "#REF!")
    if text.startswith("#") and text.endswith("!"):
        return None

    # Numbers stored as strings in Excel (e.g. "2.87", "3.26")
    try:
        return float(text)
    except ValueError:
        pass

    return text


def cell_val(ws, row: int, col: int):
    """Read one cell from the worksheet and return its coerced value.

    Convenience wrapper used when extracting single metadata cells
    (Focus, Source, Role, Target rows) by row and column index.
    """
    return coerce_value(ws.cell(row=row, column=col).value)
