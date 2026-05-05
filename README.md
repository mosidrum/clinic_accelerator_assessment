# Clinic Scoreboard — Excel → JSON

Reads `Scoreboard Test.xlsx` and writes a clean `output.json` any dashboard or script can consume directly.

---

## How to run it

Requires Python 3.10+. Check with:

```bash
python3 --version
```

If `pip` is not found, install it first:

```bash
# macOS / Linux
python3 -m ensurepip --upgrade

# or, if that fails
curl https://bootstrap.pypa.io/get-pip.py | python3
```

Then install dependencies and run:

```bash
python3 -m pip install -r requirements.txt
python3 convert.py
# → output.json written to the same directory
```

> Use `python3 -m pip` instead of `pip` — it always targets the same Python you're running the script with.

Optional flags:

```bash
python3 convert.py --sort total_revenue_all_services --order desc
python3 convert.py --sort date --order asc --sort-keys
python3 convert.py --list-metrics   # print all sortable key names
```

---

## JSON shape

Two top-level keys: `metadata` and `records`.

**`metadata`** — one entry per metric so records are self-describing:

```json
"answer_rate": {
  "label":   "Answer Rate",
  "column":  "AL",
  "section": "PHONE PERFORMANCE",
  "focus":   "Answer",
  "source":  "CallHero",
  "role":    "J"
}
```

**`records`** — one flat object per week:

```json
{
  "week_ending": "2026-02-09",
  "total_revenue_all_services": 42360.64,
  "answer_rate": 0.94
}
```

**Why flat?** A consumer can sort, filter, or chart any metric in one line without pre-processing. The metadata block means you never need to open Excel to know what a key refers to.

---

## Messy bits — decisions and trade-offs

| Problem | Decision | Trade-off |
|---|---|---|
| **Merged cell** (row 1, cols AJ–AP) | Expanded the merge — every covered column gets `"section": "PHONE PERFORMANCE"` | Other service-line groups (PT, RMT, etc.) have no banner, so their `section` stays `null` rather than guessed |
| **~18 spacer columns** (no row-2 header) | Skipped entirely | A spacer that gains data without a header would silently disappear |
| **Duplicate headers** ("PVA 4 wk avg", "Utilization", etc. repeated per service line) | Appended the Excel column letter to later occurrences — e.g. `pva_4_wk_avg_df` | Requires a metadata lookup to identify which service line a suffixed key belongs to |
| **Broken formula** (`#REF!` in one cell) | Coerced to `null` | Stale cached values from unrecalculated formulas can't be detected at parse time |
| **Numbers stored as text** | Attempted `float()` on every string; kept original on failure | A text value that looks like a number would be silently converted — no such case exists here |

---

## With two more hours

1. **Infer missing section labels.** PT / RMT / CHIRO / Pelvic Health groups have no merged banner, so their `section` is `null`. Detect group boundaries from the revenue header of each service line and back-fill.
2. **Round-trip test.** Re-read `output.json` and assert every value matches the raw openpyxl cell. Catches silent column-mapping drift when the spreadsheet is edited.
3. **Percentage flag.** Add `"is_percentage": true` to metric metadata so dashboards can format `0.94` as `94%` without parsing the label.
