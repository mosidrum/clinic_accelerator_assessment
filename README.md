# Clinic Scoreboard — Excel to JSON Converter

## How to run it

```bash
# 1. Clone / unzip the project and enter the directory
cd clinic-accelerator

# 2. Create a virtual environment and install dependencies
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# 3. Run the converter (Scoreboard Test.xlsx must be in the same directory)
python3 convert.py
# → output.json is written to the same directory
```

Python 3.10+ required (uses built-in type-hint syntax).

---

## JSON shape

The file has two top-level keys: `metadata` and `records`.

### `metadata`

Documents every metric column so the records are self-describing.

```json
{
  "metadata": {
    "source_file": "Scoreboard Test.xlsx",
    "sheet": "SCOREBOARD",
    "extracted_at": "2026-05-05T10:00:00+00:00",
    "record_count": 3,
    "metric_count": 124,
    "metrics": {
      "total_revenue_all_services": {
        "label":   "Total Revenue - All Services",
        "column":  "B",
        "section": null,
        "focus":   "Financial",
        "source":  "EMR",
        "role":    "J",
        "target":  null
      },
      "answer_rate": {
        "label":   "Answer Rate",
        "column":  "AL",
        "section": "PHONE PERFORMANCE",
        "focus":   "Answer",
        "source":  "CallHero",
        "role":    "J",
        "target":  null
      }
    }
  }
}
```

### `records`

One object per weekly row, sorted oldest-first. Every key is a slug derived
from the column header; `week_ending` is an ISO date string.

```json
{
  "records": [
    {
      "week_ending": "2026-02-02",
      "total_revenue_all_services": 39202.17,
      "total_payments_all_services": 40068.6,
      "revenue_collected_4_wk_avg": 1.022,
      "ar_90_days": 60.0,
      "total_ar": 57239.65,
      "answer_rate": 0.94,
      "booking_rate": 0.78,
      "pt_total_revenue": 30394.0,
      "rmt_total_revenue": 775.5,
      "pva_4_wk_avg": 6.07,
      "pva_4_wk_avg_df": 2.5
    }
  ]
}
```

**Why this shape?**

A flat record object means a developer can query without any pre-processing:

```python
# Top weeks by PT revenue
sorted(data["records"], key=lambda r: r["pt_total_revenue"] or 0, reverse=True)
```

The `metadata.metrics` block makes the records self-describing — you can look
up what `pva_4_wk_avg_df` means (RMT PVA, column DF) without opening Excel.

---

## Messy bits — what we found and what we did

| Problem | What we found | Decision | Trade-off |
|---|---|---|---|
| **Merged cell** | Row 1 has one merged cell (AJ:AP) labelling the "PHONE PERFORMANCE" section | Expanded the merge so every covered column gets the banner text as its `section` field | Other sections (PT, RMT, CHIRO, Pelvic Health) have no merged cell — their grouping is implicit. We left `section: null` for those rather than guessing from position. |
| **Spacer columns** | ~18 columns have no header in row 2 (columns S, X, AI, AQ, BB, BE, BI, CF, CH, CI, DB, DN, DZ, EE, EH, and others) | Skipped entirely — no header means no data worth preserving | If a spacer column ever gains data without a header it would silently disappear. Acceptable for a display-layout spreadsheet. |
| **Duplicate column headers** | "PVA (4 wk avg)", "Utilization", "TP Utilization", and "Conversion" each appear 2–4 times for different service lines (PT / RMT / CHIRO / Pelvic Health) | Appended the Excel column letter to the slug of the later occurrences (e.g. `pva_4_wk_avg_df` for RMT, column DF) | The suffix is stable and directly traceable in `metadata.metrics["pva_4_wk_avg_df"]["column"]`. It does require consulting metadata to tell the service lines apart. |
| **Formula cells** | `data_only=True` in openpyxl returns the last-computed value. One cell (ED10) contained `#REF!` — a broken formula | Treated any `#REF!`-style string as `null` | If the workbook hasn't been recalculated since formulas broke, cached values may be stale. Nothing we can do at parse time. |
| **Numbers stored as strings** | Column AT ("Average Prescribed") stores values like `"2.87"` as text | Attempted `float()` conversion on every string value; falls back to keeping the string if it fails | Silent coercion means a genuinely text value that looks like a number would be converted. No such case exists in this file. |
| **Multi-row header block** | Rows 3–7 contain "Focus", "Source", "Role", target flags, and spot target values rather than data | Rows 3–6 are mapped into `metadata.metrics` per column. Row 7's sparse target values are stored in the `target` field of each metric. | Row 7 also has `6.0` in column A and a backtick in column B — clearly layout artifacts. They land in the `target` field of those two columns, which is harmless but slightly noisy. |
| **Inconsistent header whitespace and newlines** | Headers contain `\n`, leading/trailing spaces, and mixed capitalisation | Normalised before slugifying: strip → replace `%`→`_pct`, `#`→`_num` → collapse non-alphanumeric to `_` → lowercase | The `label` field in metadata keeps the original text exactly as found. |

---

## If I had two more hours

1. **Infer service-line sections from column position.** The PT/RMT/CHIRO/Pelvic Health groups have no merged banner, so their `section` is currently `null`. I'd detect group boundaries by scanning for section-title headers ("PT Total Revenue", "RMT Total Revenue", etc.) and back-fill `section` for every column between them.

2. **Round-trip validation.** Add a small test that re-reads `output.json` and asserts every non-null value matches the raw cell value from openpyxl. Catches off-by-one column mapping bugs silently introduced by future spreadsheet edits.

3. **CLI flags.** Accept `--input` / `--output` / `--sheet` as arguments so the script works on any similarly-structured scoreboard workbook without editing source code.

4. **Percentage normalization.** Cells like `answer_rate: 0.94` are stored as decimals in Excel but labelled as percentages. I'd add a `"is_percentage": true` flag to the metric metadata so a dashboard can format them correctly without parsing the label.
