"""
convert.py — Scoreboard Test.xlsx → output.json  (entry point)

Orchestrates the full pipeline by delegating to focused sub-modules:

    constants.py  — file paths and row-number constants
    utils.py      — value coercion and key-slug helpers
    parser.py     — spreadsheet structure parsing
    sorters.py    — record and key sorting
    cli.py        — command-line argument parsing

Run with:
    python3 convert.py [options]

See cli.py or run --help for all available flags.
"""

import json
import sys
import datetime
from pathlib import Path

import openpyxl

from constants import INPUT_FILE
from cli      import parse_args
from parser   import parse_section_headers, build_column_map, extract_records
from sorters  import sort_by_date, sort_by_metric, sort_keys_alphabetically


def main() -> None:
    args = parse_args()

    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"Input file not found: {INPUT_FILE}")

    # data_only=True returns cached formula results instead of formula strings
    wb = openpyxl.load_workbook(INPUT_FILE, data_only=True)
    ws = wb["SCOREBOARD"]

    section_map = parse_section_headers(ws)
    col_map     = build_column_map(ws, section_map)

    # Build the self-describing metadata block — one entry per metric column
    metrics_meta = {
        meta["key"]: {
            "label":   meta["label"],
            "column":  meta["column"],
            "section": meta["section"],
            "focus":   meta["focus"],
            "source":  meta["source"],
            "role":    meta["role"],
            "target":  meta["target"],
        }
        for meta in col_map.values()
    }

    # --list-metrics: print key→label table and exit without touching the output file
    if args.list_metrics:
        print(f"{'KEY':<45}  {'LABEL'}")
        print("-" * 80)
        for key, meta in metrics_meta.items():
            print(f"{key:<45}  {meta['label']}")
        sys.exit(0)

    records = extract_records(ws, col_map)

    # Apply record-level sort
    if args.sort == "date":
        records = sort_by_date(records, args.order)
    else:
        records = sort_by_metric(records, args.sort, args.order)

    # Optionally sort metric keys A→Z within each record
    if args.sort_keys:
        records = sort_keys_alphabetically(records)

    output = {
        "metadata": {
            "source_file":  INPUT_FILE.name,
            "sheet":        ws.title,
            "extracted_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
            "record_count": len(records),
            "metric_count": len(metrics_meta),
            "sort_applied": {
                "sort_by":    args.sort,
                "order":      args.order,
                "keys_alpha": args.sort_keys,
            },
            "metrics": metrics_meta,
        },
        "records": records,
    }

    out_path = Path(args.output)
    out_path.write_text(json.dumps(output, indent=2, ensure_ascii=False))
    print(f"Written {len(records)} records, {len(metrics_meta)} metrics → {out_path}")


if __name__ == "__main__":
    main()
