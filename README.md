# Excel Toolkit — Merge, Clean, Split (no data loss, full audit trail)

Three small command-line tools for people who get a folder of messy Excel files and
need one clean table out of it. No GUI to click through, no "merge wizard" that
silently drops your columns, no manual copy-paste that quietly corrupts dates.

Built for the case that keeps showing up in real work:

> "I have 20 monthly exports from our ERP. Every file has the same data but the
> column names are spelled differently, dates come in four formats, some numbers
> are text, and half of them have a title row on top. I need one table."

| | Manual | This toolkit |
|---|---|---|
| 12 files × 500 rows merged | 40–90 min of copy-paste | ~2 s |
| Column names differ per file | you rename them by hand, one by one | auto-aligned, with a list of what it matched |
| Dates in 4 formats | you re-type them | unified to `yyyy-mm-dd`, every change logged |
| Did every row survive? | you hope so | the tool prints rows in vs rows out and fails loudly if they differ |
| What exactly changed? | nobody knows | a `_log` sheet: row, column, before, after, reason |

The design rule behind all three tools: **a tool may transform data, but it may
never transform it silently.** Every cell that changes shows up in a log. Anything
the tool cannot decide — a date that could honestly be read two ways — is left
alone and flagged, not guessed at.

---

## Requirements

- Python 3.8+
- `openpyxl` (required)
- `xlrd` (optional — only needed to read legacy Excel 97-2003 `.xls` files)

```bash
pip install openpyxl
pip install xlrd      # optional: adds .xls support
```

Tested on macOS and Windows with Python 3.9–3.13. No pandas, no compiled
dependencies, no internet access needed.

Without `xlrd` the tools do not crash on `.xls` — that file is skipped with a
one-line reason and everything else still merges.

---

## Quick start

```bash
# 1. merge every .xlsx / .csv in a folder into one sheet
python3 excel_merge.py --src ./monthly_exports

# 2. clean it: whitespace, dates, text-numbers, duplicates
python3 excel_clean.py --src merged_20260101_1200.xlsx --dedupe

# 3. split the clean result into one file per region
python3 excel_split.py --src merged_20260101_1200_clean.xlsx --by Region
```

Each command prints a summary and writes a report sheet. Nothing is overwritten:
output always goes to a new file.

---

## 1. `excel_merge.py`

Merge every `.xlsx` / `.xlsm` / `.xls` / `.csv` in a folder into a single sheet.

```bash
python3 excel_merge.py --src ./invoices --out all_invoices.xlsx --preview
```

Legacy `.xls` (Excel 97-2003) is read directly when `xlrd` is installed, including
real date cells — they come through as dates, not as serial numbers like `46118`.
That matters more than it sounds: half of the ERP exports small businesses still
produce are `.xls`.

| Flag | What it does |
|---|---|
| `--src DIR` | folder to read (required) |
| `--out FILE` | output file (default: `merged_<timestamp>.xlsx`) |
| `--files a.xlsx b.csv` | only these files |
| `--sheet NAME` | read this sheet name from every workbook |
| `--all-sheets` | read every sheet, not just the first |
| `--header-row N` | header is on row N (default: auto-detect) |
| `--alias my.json` | extend the built-in column-alias table |
| `--no-source` | don't add `_source_file` / `_source_sheet` columns |
| `--fuzzy 0.86` | header similarity threshold for auto-alignment |
| `--preview` | print the first 3 output rows |

**What it handles**

- **Different header spellings for the same column.** `Order ID`, `order_id`,
  `Order No.`, `订单号` all land in one column. Every match is reported twice:
  once on the console and once on the `_aliases` sheet, so you can audit it.
- **A title row above the header** (`Exported from ERP — do not edit`). The tool
  picks the first row that is at least half as wide as the widest of the first ten.
- **Blank rows** in the middle of data. Counted and skipped, never shifted.
- **Rows wider than the header.** Extra values become `未命名列_n` / `Column_n`
  instead of being thrown away.
- **A broken file in the folder.** Reported by name, the rest still merge.
  One corrupt workbook does not stop the batch.
- **`.csv` with `;` or tab delimiters**, in UTF-8 or GBK.

**What it deliberately does not do:** change values. Merge only concatenates.
A date written `2026/3/5` in one file and `05.03.2026` in another stays that way
until you run `excel_clean.py`. Separating "combine" from "repair" means that when
something looks wrong you know which of the two steps did it.

Output: `Merged` sheet + `_report` sheet (rows read / blank / written per file,
with a reconciliation line) + `_aliases` sheet.

---

## 2. `excel_clean.py`

Clean one workbook and write a cell-by-cell audit log.

```bash
python3 excel_clean.py --src dirty.xlsx --dedupe --out clean.xlsx
```

| Flag | What it does |
|---|---|
| `--src FILE` | input `.xlsx` or legacy `.xls` (required) |
| `--out FILE` | output (default: `<name>_clean.xlsx`) |
| `--sheet NAME` | sheet to clean |
| `--header-row N` | header is on row N (default: auto-detect) |
| `--date-order auto\|dmy\|mdy` | how to read `05.03.2026` (default: auto, per column) |
| `--dates iso\|keep` | unify to `yyyy-mm-dd` (default) or leave dates alone |
| `--numbers coerce\|keep` | turn text-numbers into real numbers (default) |
| `--dedupe` | drop exact duplicate rows |
| `--dedupe-on A B` | judge duplicates on these columns only |
| `--max-log N` | cap the `_log` sheet at N rows (default 20000). Truncation is written into the sheet and printed — never silent |
| `--keep-empty-rows` | don't drop fully empty rows |

**What it fixes**

- Trailing / double spaces, non-breaking spaces, zero-width characters
- Full-width digits and punctuation → half-width (`１２３` → `123`)
- **Dates** in any of these forms → one real Excel date:
  `2026-03-05`, `2026/3/5`, `2026年3月5日`, `05.03.2026`, `2026-03-05 14:30`,
  `20260305`, and bare Excel serial numbers (`46000`). Time-of-day is kept.
- **Numbers stored as text**: `"1,234.50"`, `"¥ 88 元"`, `" 12 "` → `1234.5`, `88`, `12`
- **IDs keep their leading zeros.** `00123` is never turned into `123`. The tool
  recognises the pattern and skips coercion for that column.
- **Duplicate rows**, optional, with the dropped row numbers listed.
- Empty cells are counted per column and reported. **They are never filled in** —
  inventing a value is not cleaning.

**The one thing worth reading about: ambiguous dates.**
`05.03.2026` is either 5 March or 3 May. A careless tool picks one. This one:

1. looks for proof inside the same column — if any value has a component above 12
   (e.g. `27.01.2026`), the whole column is proven day-first;
2. if there is no proof, falls back to the separator (`.` → day-first, `/` → month-first);
3. if both readings are still legal, **leaves the value untouched** and logs it as
   `AMBIGUOUS` so you can decide with `--date-order dmy|mdy`;
4. and records the assumption in the `_summary` sheet when it does decide.

Output: `Cleaned` sheet + `_log` sheet (Excel row, column, before, after, action) +
`_summary` sheet (rows in / dropped / out, reconciliation, per-column empty counts).

---

## 3. `excel_split.py`

Split one sheet into many — by a column value, or into fixed-size chunks.

```bash
python3 excel_split.py --src all.xlsx --by Region                  # one file per region
python3 excel_split.py --src all.xlsx --by 3 --mode sheets        # one sheet per value
python3 excel_split.py --src all.xlsx --chunk 500                 # 500 rows per file
```

| Flag | What it does |
|---|---|
| `--src FILE` | input `.xlsx` / `.xlsm` / legacy `.xls` (required) |
| `--by COL` | column name or 1-based index to group by |
| `--chunk N` | split into blocks of N rows instead |
| `--mode files\|sheets` | separate workbooks, or sheets inside one workbook |
| `--out PATH` | output folder (`files`) or output file (`sheets`) |
| `--prefix STR` | file-name prefix |
| `--keep-blank` | keep rows whose key cell is empty |

File and sheet names are sanitised for Windows and macOS (`/ \ : * ? " < > |`
replaced), long values truncated, duplicates suffixed. In `sheets` mode an
`_index` sheet lists every group with its row count.

---

## Verifying it yourself

The repo ships with a deliberately dirty test set and a generator for it:

```bash
cd tools
python3 make_test_data.py                 # builds samples/dirty
python3 excel_merge.py --src samples/dirty --out /tmp/m.xlsx
python3 excel_clean.py --src /tmp/m.xlsx --out /tmp/c.xlsx --dedupe
```

`samples/dirty` contains: 8 workbooks with the same data under 5 different header
spellings, 5 date formats including Excel serials, numbers as text with currency
symbols, leading-zero IDs, nbsp padding, full-width characters, blank rows,
duplicate rows, a header that starts on row 3, one semicolon-delimited CSV, one
legacy `.xls`, and one corrupt file.

Expected result: the two bad files are reported and the rest merge. Rows read must
equal rows written — the tool says so itself on the `_report` sheet.

---

## FAQ

**Does it need Excel installed?** No. `openpyxl` reads and writes `.xlsx` directly.

**How fast is it, really?** Measured, not estimated: 60,000 rows spread over 3 files —
merge 6 s, clean 10 s (104,989 cell changes logged), split into 5 files 9 s. On a
slower machine expect minutes, not hours. The tools never load the whole workbook
into a GUI, so file size stops being a reason to do it by hand.

**Will it break my formulas?** Values are read with `data_only=True`, so formulas
come through as their last calculated value. This is a data-extraction tool, not an
editor — run it on exports, not on your live model.

**`.xls`?** Supported directly across all three tools (merge, clean, split), as long
as `xlrd` is installed (`pip install xlrd`). The `.xls` in `samples/dirty` is a
deliberately fake one — a text file renamed to `.xls` — to prove the tool reports a
bad file instead of dying. A real legacy file lives in
`samples/legacy/legacy_export.xls` if you want to see the working path.

**Files with passwords, merged cells, or pivot caches?** Merged cells are read as
"value in the top-left, blanks elsewhere". Password-protected workbooks are not
supported. Save a copy without protection.

**Can it handle 100k rows?** Yes, but memory grows with the file. For very large
workloads, split by month first (`--chunk`).

**Something went wrong.** `MERGE_DEBUG=1 python3 excel_merge.py ...` prints the
full traceback.

---

## License

MIT — use it at work, in client projects, commercially. No attribution required.

Not covered: any guarantee that your upstream export is sane. These tools make
their transformations visible; they cannot make bad source data good.
