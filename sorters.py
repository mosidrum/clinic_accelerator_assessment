"""
sorters.py — Record sorting functions.

Each function takes a list of record dicts (as produced by parser.extract_records)
and returns a new sorted list.  The original list is never mutated.

Sorting operates on two dimensions:
  1. Record order  — which week comes first in the output list.
  2. Key order     — which metric keys appear first inside each record object.

These dimensions are independent and can be combined freely.
"""


def sort_by_date(records: list[dict], order: str) -> list[dict]:
    """Sort records by the 'week_ending' ISO date string.

    Args:
        records: List of record dicts, each containing a 'week_ending' key.
        order:   'asc'  → oldest week first (earliest date at index 0).
                 'desc' → newest week first (most recent date at index 0).

    Returns:
        A new sorted list; the input list is unchanged.
    """
    reverse = order == "desc"
    return sorted(records, key=lambda r: r["week_ending"], reverse=reverse)


def sort_by_metric(records: list[dict], metric_key: str, order: str) -> list[dict]:
    """Sort records by the value of a single numeric metric field.

    Records where the metric value is None (missing data) are always placed
    at the end of the list regardless of sort direction, so sparse columns
    don't push real data out of view.

    Args:
        records:    List of record dicts.
        metric_key: The JSON key to sort by (e.g. 'total_revenue_all_services').
        order:      'asc'  → smallest value first  (min → max).
                    'desc' → largest value first   (max → min).

    Returns:
        A new sorted list; the input list is unchanged.

    Raises:
        ValueError: If metric_key is not present in the records.
                    The error message instructs the user to run --list-metrics.
    """
    if not records:
        return records

    if metric_key not in records[0]:
        raise ValueError(
            f"Unknown metric key: '{metric_key}'.\n"
            f"Run --list-metrics to see all available keys."
        )

    reverse = order == "desc"

    def sort_key(record):
        val = record.get(metric_key)
        # Tuple trick: (0, value) sorts before (1, 0) so None always goes last
        if val is None:
            return (1, 0)
        return (0, val if not reverse else -val)

    return sorted(records, key=sort_key)


def sort_keys_alphabetically(records: list[dict]) -> list[dict]:
    """Re-order the keys inside each record object alphabetically (A → Z).

    'week_ending' is always pinned as the first key so the record identifier
    is immediately visible when scanning the JSON.  All other metric keys
    follow in ascending alphabetical order.

    This does not change the order of records in the list — combine it with
    sort_by_date or sort_by_metric to control both dimensions at once.

    Args:
        records: List of record dicts.

    Returns:
        A new list of records with reordered keys; values are unchanged.
    """
    result = []
    for record in records:
        other_keys = sorted(k for k in record if k != "week_ending")
        ordered = {"week_ending": record["week_ending"]}
        ordered.update({k: record[k] for k in other_keys})
        result.append(ordered)
    return result
