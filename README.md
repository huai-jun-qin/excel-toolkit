# Excel Toolkit — Merge, Clean, Split (no data loss, full audit trail)

> **Rather not run scripts yourself?** I do this as a service — you send the files,
> you get back one clean workbook plus a cell-by-cell change log.
> → [fiverr.com/claire999j](https://www.fiverr.com/claire999j)

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
python3 -m pip install --user openpyxl
python3 -m pip install --user xlrd      # optional: adds .xls support
```

**Use `python3 -m pip`, not bare `pip`.** If more than one Python is installed
(macOS ships one, Homebrew adds another, pyenv adds more), a plain `pip install`
often lands in a different interpreter than the one that runs the script — and
you then get `No module named 'openpyxl'` on a machine that "definitely has it".
`python3 -m pip` always targets the interpreter you are about to run.

Check before running:

```bash
python3 -c "import openpyxl; print(openpyxl.__version__)"
```

If pip refuses with `externally-managed-environment`, add `--break-system-packages`.
When the import is still missing, the scripts now print the exact interpreter they
are running under and the command that installs into it, instead of a bare
"pip install openpyxl".

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
| `--all-sheets` | also read pages named like a summary (`汇总`, `Total`, `Pivot`) |
| `--first-sheet` | read only the first data sheet per file (the old behaviour) |
| `--skip-sheet NAME` | never read that sheet (repeatable) |
| `--header-row N` | header is on row N (default: auto-detect) |
| `--strict` | exit code 2 if any sheet was skipped, so a cron job notices |
| `--no-fill-merged` | don't copy a merged cell's value across its block |
| `--alias my.json` | extend the built-in column-alias table |
| `--no-source` | don't add `_source_file` / `_source_sheet` columns |
| `--fuzzy 0.86` | header similarity threshold for auto-alignment |
| `--preview` | print the first 3 output rows |

**What it handles**

- **Different header spellings for the same column.** `Order ID`, `order_id`,
  `Order No.`, `订单号` all land in one column. Every match is reported twice:
  once on the console and once on the `_aliases` sheet, so you can audit it.
- **A title / info line above the header.** `Exported from ERP — do not edit`, or
  `单位：元 | 期间：2026年1-3月 | 制表：李四`. The detector scores the first ten rows
  on how much they look like short column labels rather than data or a `key：value`
  info line, instead of just taking the widest one.
- **A cover sheet.** Workbooks often open with `2026年3月对账单 / 制表：李四`. That
  page is recognised as a cover and skipped — not merged as a one-row table.
- **Many sheets in one file.** `1月`, `2月`, `3月` in a single workbook are all read
  by default. Previously only the first was read, so 11 of 12 months disappeared
  without a warning.
- **A summary sheet** (`汇总`, `合计`, `Total`, `Pivot`). Skipped by default because
  merging it double-counts. `--all-sheets` overrides that.
- **The header printed again inside the data** — common in one-month-per-page
  exports. Those rows are dropped instead of landing as a `Name | Email` data row.
- **Merged cells.** A `地区` cell merged over rows 2–4 means one value shown once;
  reading it naively gives you the value on row 2 and blanks below. The value is
  copied across its block (turn it off with `--no-fill-merged`), and the count is
  reported per sheet.
- **Two columns that want the same name.** A sheet with both `金额` and `合计`
  used to end up with one column holding whichever came last — the values changed,
  the row count stayed correct, and nothing warned. They are now kept apart and
  the collision is listed on the console.
- **Display formats that carry meaning.** An order id stored as `123` with format
  `0000000`, or an amount in `¥#,##0.00`, keeps that format in the output. Date
  formats are deliberately *not* carried: writing one makes the stored serial
  number read as a date, and interpreting dates is `excel_clean.py`'s job.
- **Blank rows** in the middle of data. Counted and skipped, never shifted.
- **Rows wider than the header.** Extra values become `未命名列_n` / `Column_n`
  instead of being thrown away.
- **A broken file in the folder.** Reported by name, the rest still merge.
  One corrupt workbook does not stop the batch.
- **`.csv` with `;` or tab delimiters**, in UTF-8 or GBK.

Nothing is skipped silently: every sheet that was not read is listed on the console
and on the `_report` sheet under `SHEETS NOT READ`, with the reason.

**What it deliberately does not do:** change values. Merge only concatenates.
A date written `2026/3/5` in one file and `05.03.2026` in another stays that way
until you run `excel_clean.py`. Separating "combine" from "repair" means that when
something looks wrong you know which of the two steps did it.

Output: `Merged` sheet + `_report` sheet (rows read / blank / repeated header /
written per file, with a reconciliation line that must balance) + `_aliases` sheet.

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
  A bare number is only read as a date when the **column header says it is one**
  (`日期` / `时间` / `date` / `time`) or the column already holds real dates.
  Without that rule a salary column of `35000` becomes 1995-10-28 — a real bug
  this tool shipped in an earlier version, and the reason `selftest_clean.py`
  exists.
- **Merged blocks.** A `地区` cell merged over `B2:B4` stores its value only in
  `B2`; read plainly, rows 3-4 look like missing data. The value is copied across
  the block and the count is reported (`--no-fill-merged` turns it off).
- **Display formats that change what you see** are carried over: an order id
  `123` keeps its `0000000` format, money keeps its `¥`. Date formats are
  deliberately not carried — the value is already a real date.
- **Numbers stored as text**: `"1,234.50"`, `"¥ 88 元"`, `" 12 "` → `1234.5`, `88`, `12`
- **IDs keep their leading zeros.** `00123` is never turned into `123`. The tool
  recognises the pattern and skips coercion for that column.
- **Duplicate rows**, optional, with the dropped row numbers listed.
- Empty cells are counted per column and reported. **They are never filled in** —
  inventing a value is not cleaning. The one exception is a merged block, where
  the value already exists and Excel merely stores it once (see above).
- **A cover sheet above the data.** Instead of cleaning the title page, it picks the
  first sheet that actually looks like a table. If the workbook holds several
  sheets with data, it says so on the console and names them — cleaning writes one
  sheet, so the rest would otherwise vanish from the output without a word.

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

If the workbook opens with a cover page, the split starts from the first sheet that
is an actual table — splitting a title page by "column" has nothing to group by.

**Rows are never dropped quietly.** A key cell that sits inside a merged block is
filled from that block first: a 3-row table split on a merged `地区` column used to
come out as one row, with two rows gone and a single console line. A row whose key
is genuinely empty is still skipped, but it is counted and printed. A header row
printed again inside the data becomes nothing — it used to become a group named
`地区` holding a copy of the header.

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

The awkward shapes that break most merge scripts have their own test file:

```bash
cd tools
python3 selftest_merge.py     # 13 cases
python3 selftest_clean.py     # 10 cases
python3 selftest_split.py     #  7 cases
```

Each one builds its workbooks on the fly and asserts the output cell by cell.

`selftest_merge.py` covers a cover page above the data, three months inside one
file, a 4-cell `单位：元 | 期间：…` info line sitting above a 7-column header, the
header printed twice, a sheet named `汇总`, a CSV with BOM and semicolons, a folder
holding an instructions-only file, `--first-sheet` reading exactly one month, a
50-row single-column list (must stay data, not be mistaken for a cover), a 2x2
table, a merged `地区` block, one sheet holding both `金额` and `合计`, and an order
id with an `0000000` format.

`selftest_clean.py` covers the case that matters most: a salary column of
35000 / 28000 / 42000 must stay numbers. It used to come out as 1995-10-28 /
1976-08-28 / 2014-12-27, because those integers fall in Excel's date-serial range
and the log called it "date unified" while reconciliation still said OK. It also
covers three date spellings in one column, `¥1,234` text inside a numeric column,
a merged block, a `0000000` order id, a cover sheet, deduplication, and leading
zeros.

`selftest_split.py` covers a merged `地区` block (3 rows in, 3 rows out — it used to
be 1), an empty key, `华东/上海` as a file name, fixed-size chunks, a repeated header
row, a cover sheet, and `35000` vs `35000.0` landing in the same group.

All of them print a diff-friendly `[PASS]/[FAIL]` line per case and exit non-zero
on failure. Point them at an older copy to see what a change actually fixed:

```bash
MERGE_SCRIPT=/tmp/old_merge.py python3 selftest_merge.py
CLEAN_SCRIPT=/tmp/old_clean.py python3 selftest_clean.py
SPLIT_SCRIPT=/tmp/old_split.py python3 selftest_split.py
```

---

## FAQ

**Does it need Excel installed?** No. `openpyxl` reads and writes `.xlsx` directly.

**How fast is it, really?** Measured, not estimated: 60,000 rows spread over 3 files —
merge 6 s, clean 7 s (104,989 cell changes logged), split into 5 files 9 s. On a
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

## Want this done for you?

The scripts are free and always will be. But if the files are due tomorrow and you
would rather not touch a terminal, I run this as a paid service:

- you send the folder as-is (`.xlsx`, `.xls`, `.csv`, GBK, whatever your ERP exports)
- you get back one workbook, a `_log` sheet listing every cell that changed with
  before/after, and a short note on what was wrong with the original data
- if something is ambiguous — a date that could be read two ways — you get a question,
  not a guess

→ **[fiverr.com/claire999j](https://www.fiverr.com/claire999j)**

Not sure whether your file fits? Message me with the column headers and a couple of
sample rows. I will tell you honestly if it is something these tools cannot handle,
instead of taking the order.

---

## License

MIT — use it at work, in client projects, commercially. No attribution required.

Not covered: any guarantee that your upstream export is sane. These tools make
their transformations visible; they cannot make bad source data good.
