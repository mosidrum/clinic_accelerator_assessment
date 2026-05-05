"""
cli.py — Command-line interface definition for convert.py.

Defines and parses all flags accepted by the converter.  Keeping argument
parsing here means convert.py's main() stays focused on orchestration and
this module can be tested or extended independently.

Flags
-----
    --sort KEY          Sort field: 'date' or any metric key.  Default: date
    --order asc|desc    Sort direction.  Default: asc
    --sort-keys         Sort metric keys A→Z within each record
    --output PATH       Destination file.  Default: output.json
    --list-metrics      Print all metric keys and exit

Examples
--------
    python3 convert.py
    python3 convert.py --sort total_revenue_all_services --order desc
    python3 convert.py --sort date --order desc --output newest_first.json
    python3 convert.py --sort-keys --output alpha_keys.json
    python3 convert.py --list-metrics
"""

import argparse

from constants import OUTPUT_FILE


def parse_args() -> argparse.Namespace:
    """Build the argument parser and return the parsed namespace.

    Called once at the start of main().  All defaults are defined here
    so they are visible in one place alongside the flag descriptions.
    """
    parser = argparse.ArgumentParser(
        description="Convert Scoreboard Test.xlsx to a clean JSON file.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )

    parser.add_argument(
        "--sort",
        default="date",
        metavar="KEY",
        help=(
            "Field to sort records by. "
            "Use 'date' for chronological order or any metric key "
            "(e.g. total_revenue_all_services). Default: date"
        ),
    )

    parser.add_argument(
        "--order",
        choices=["asc", "desc"],
        default="asc",
        help=(
            "Sort direction: "
            "asc (min→max / oldest→newest) or "
            "desc (max→min / newest→oldest). Default: asc"
        ),
    )

    parser.add_argument(
        "--sort-keys",
        action="store_true",
        help="Sort metric keys A→Z within each record (week_ending always stays first).",
    )

    parser.add_argument(
        "--output",
        default=str(OUTPUT_FILE),
        metavar="PATH",
        help=f"Output file path. Default: {OUTPUT_FILE}",
    )

    parser.add_argument(
        "--list-metrics",
        action="store_true",
        help="Print all available metric keys and exit without writing JSON.",
    )

    return parser.parse_args()
