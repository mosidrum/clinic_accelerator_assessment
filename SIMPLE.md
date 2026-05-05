# What Is This Project? (Explained Simply)

---

## The Big Idea

Imagine you have a giant piece of paper with lots and lots of numbers on it.
The paper is about a clinic — like a doctor's office — and it tracks things like:

- How much money did we make this week?
- How many people called us?
- How many people showed up for their appointment?

That giant piece of paper is called **`Scoreboard Test.xlsx`** — it's a spreadsheet,
like a really big table you open in Microsoft Excel.

The problem is — spreadsheets are great for humans to look at, but **computers
find them really hard to read**. They have messy layouts, blank rows, weirdly
merged boxes, and lots of other headaches.

So we wrote a program that **reads that messy spreadsheet and turns it into a
clean, neat file** that any computer program, website, or dashboard can
understand instantly.

That clean file is called **`output.json`**.

---

## What Is a JSON File?

Think of JSON like a really organised toy box.

Instead of just throwing all your toys in a pile (like a spreadsheet does),
JSON puts every toy in a labelled slot:

```
"total_revenue_all_services" = 42360.64
"answer_rate"                = 0.94
"week_ending"                = "2026-02-09"
```

Any computer program in the world can open that box, look at the label, and
instantly grab exactly the toy it needs — no digging through a pile.

---

## The Files — What Each One Does

Think of the project like a kitchen. Everyone has one job.

---

### `Scoreboard Test.xlsx` — The Ingredients

This is the raw spreadsheet. It's the messy pile of data we start with.
Nobody touches this file — we only read it.

---

### `constants.py` — The Recipe Card

Before anyone starts cooking, someone writes down the rules that never change:

- What is the spreadsheet called? (`Scoreboard Test.xlsx`)
- What should the output file be called? (`output.json`)
- Which row in the spreadsheet has the column names? (Row 2)
- Which row has the actual numbers? (Row 8 onwards)

Every other file in the kitchen reads from this one card so nobody has to
remember or guess.

---

### `utils.py` — The Kitchen Helper

This file does the small, boring, useful jobs that everyone else needs done:

**Job 1 — `slugify`**
Takes a messy column name like `"Total Revenue \n- All Services"` and turns
it into a clean label like `total_revenue_all_services` that a computer can
use as a key. No spaces. No weird characters. All lowercase.

**Job 2 — `make_unique_key`**
Sometimes the spreadsheet uses the same name twice for different things
(e.g. "Utilization" is used for the physio section AND the massage section).
This helper makes sure each one gets a unique name by adding the column letter
on the end — so you get `utilization` and `utilization_dw` instead of two
things with the exact same name that would confuse everybody.

**Job 3 — `coerce_value`**
The spreadsheet stores numbers, dates, and text all jumbled together.
This helper looks at each value and cleans it up:
- A date like `Feb 9 2026` becomes `"2026-02-09"`
- A number stored as text like `"2.87"` becomes the actual number `2.87`
- A broken formula like `#REF!` becomes `null` (nothing — it was broken anyway)
- An empty cell becomes `null` too

**Job 4 — `cell_val`**
A tiny shortcut. Instead of writing the same two lines of code over and over
to read one cell, this wraps it up into one neat call.

---

### `parser.py` — The Chef

This is where the real reading of the spreadsheet happens.
It has three jobs, done in order:

**Step 1 — `parse_section_headers`**
The spreadsheet has some big labels that stretch across multiple columns
(like a sign above a shelf that says "PHONE PERFORMANCE" covering 7 columns).
These are called *merged cells*. This function finds all of them and notes
which columns they cover, so we can tag each metric with its section name.

**Step 2 — `build_column_map`**
Reads every column header (row 2) and for each one builds a little info card:
- What is this column's clean name? (e.g. `answer_rate`)
- What does it track? (e.g. "Answer Rate")
- Which Excel column is it in? (e.g. "AL")
- Is it Financial? Marketing? (the Focus row)
- Where does the data come from — Jane software, CallHero, etc.? (the Source row)
- Who is responsible for it? (the Role row)

Blank columns — the ones that are just empty spaces used for layout — are
quietly ignored.

**Step 3 — `extract_records`**
Goes through every data row (row 8 and below) and turns each one into a
neat package:

```json
{
  "week_ending": "2026-02-09",
  "total_revenue_all_services": 42360.64,
  "answer_rate": 0.94,
  ...and so on for all 124 metrics
}
```

Rows that don't start with a proper date are skipped automatically —
this catches stray leftover rows from the header area.

---

### `sorters.py` — The Filing Clerk

Once we have all the weekly records, maybe we want them in a different order.
The filing clerk knows three tricks:

**Trick 1 — `sort_by_date`**
Put the oldest week first, or the newest week first.
Like stacking papers with January on top, or December on top.

**Trick 2 — `sort_by_metric`**
Pick any number column and sort by it.
Want to see the week with the most revenue at the top? Done.
Want the week with the worst answer rate at the top? Done.
Weeks with no data for that column are always pushed to the bottom so they
don't get in the way.

**Trick 3 — `sort_keys_alphabetically`**
Inside each week's record, the metric names can be sorted A to Z.
This doesn't change which week is first — it just makes the keys inside
each record easier to scan, like arranging items on a shelf alphabetically.

---

### `cli.py` — The Order Form

When you run the program, you can give it instructions:

```
python3 convert.py --sort total_revenue_all_services --order desc
```

`cli.py` reads those instructions and turns them into settings the rest of
the program can use. It's like a waiter writing down your order before
passing it to the kitchen.

The instructions you can give:

| What you type | What it does |
|---|---|
| `--sort date` | Sort weeks by date (default) |
| `--sort <any metric name>` | Sort weeks by that number |
| `--order asc` | Smallest / oldest first |
| `--order desc` | Biggest / newest first |
| `--sort-keys` | Sort the metric names A→Z inside each week |
| `--output myfile.json` | Save to a different file name |
| `--list-metrics` | Show me all the metric names I can sort by |

---

### `convert.py` — The Manager

This is the file you actually run. But it doesn't do much work itself —
it just tells everyone else what to do, in the right order:

1. Take the order (`cli.py`)
2. Open the spreadsheet
3. Find the section labels (`parser.py` → `parse_section_headers`)
4. Build the column info cards (`parser.py` → `build_column_map`)
5. Extract all the weekly records (`parser.py` → `extract_records`)
6. Sort them the way you asked (`sorters.py`)
7. Write the final `output.json` file

That's it. The manager delegates — the specialists do the actual work.

---

### `output.json` — The Result

This is the finished product. It has two sections:

**`metadata`** — a library card for every metric:
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

**`records`** — one entry for every week:
```json
{
  "week_ending": "2026-02-09",
  "total_revenue_all_services": 42360.64,
  "answer_rate": 0.94,
  "pt_total_revenue": 36099.38,
  ...
}
```

Any developer, dashboard, or AI can now open this file and ask questions like
"which week had the most revenue?" without touching the original spreadsheet at all.

---

### `requirements.txt` — The Shopping List

Before running the program on a new computer, you need to install one tool
that Python doesn't come with: `openpyxl`. This is the library that knows
how to read Excel files.

The shopping list says exactly which version to buy so everything works the
same on every computer:

```
openpyxl==3.1.5
```

---

### `README.md` — The Instruction Manual

Explains how to set up and run the project for the first time.
Also describes the shape of the JSON, the messy bits found in the spreadsheet,
and ideas for what could be improved with more time.

---

## How It All Fits Together

```
You run:   python3 convert.py --sort total_revenue_all_services --order desc

               ┌─────────────┐
               │  convert.py │  ← The manager. Runs the whole show.
               └──────┬──────┘
                      │ reads settings from
               ┌──────▼──────┐
               │   cli.py    │  ← Understands your --sort / --order flags
               └──────┬──────┘
                      │ uses defaults from
               ┌──────▼──────────┐
               │  constants.py   │  ← File names, row numbers — never changes
               └─────────────────┘

                      │ convert.py then calls
          ┌───────────┼───────────┐
          ▼           ▼           ▼
    ┌──────────┐ ┌──────────┐ ┌──────────┐
    │ parser.py│ │sorters.py│ │ utils.py │
    │          │ │          │ │          │
    │ reads the│ │ puts the │ │ cleans   │
    │spreadsh- │ │ records  │ │ up each  │
    │  eet     │ │ in order │ │  value   │
    └──────────┘ └──────────┘ └──────────┘
          │
          ▼
   Scoreboard Test.xlsx   →   output.json
   (messy spreadsheet)        (clean JSON)
```

---

## What We Actually Did — Step by Step

1. **Started with a messy Excel file** that had merged cells, blank spacer
   columns, rows used as labels instead of data, formula errors, and numbers
   stored as text.

2. **Wrote a program** (`convert.py`) to read it and produce a clean JSON file.

3. **Discovered the spreadsheet had tricky bits** — duplicate column names,
   one merged section header, formula errors (#REF!), and stray rows that
   looked like data but weren't.  We handled each one explicitly.

4. **Added sorting** so you can regenerate the output sorted any way you want
   — by date, by any metric (highest to lowest or lowest to highest), or with
   keys sorted alphabetically.

5. **Split the code into separate files** so each file has one clear job.
   A new developer can open any single file and immediately understand what
   it does without reading everything else.

6. **Tracked every change in git** with a separate commit for each step,
   so you can see exactly what changed and why at every point in history.
