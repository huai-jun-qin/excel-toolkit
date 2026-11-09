#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_merge.py - Merge many Excel / CSV files into one workbook, without touching a single value.

Design rule: merge ONLY concatenates. It never rewrites a value.
If a date says "2026/3/5" in one file and "05.03.2026" in another,
you get both, exactly as they were. Fixing them is excel_clean.py's job.

Requires: openpyxl  (pip install openpyxl)
Optional: xlrd      (pip install xlrd)   -- only needed to read legacy .xls files

Usage:
    python3 excel_merge.py --src ./invoices
    python3 excel_merge.py --src ./invoices --out merged.xlsx --preview
    python3 excel_merge.py --src ./a --files b.xlsx c.csv
"""

import argparse
import csv
import difflib
import os
import re
import sys
import traceback
from collections import Counter
from datetime import date, datetime, time

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit(
        "Missing dependency: openpyxl\n"
        "\n"
        "You are running:  %s  (Python %s)\n"
        "\n"
        "Install it for that exact interpreter:\n"
        "    %s -m pip install --user openpyxl\n"
        "\n"
        "If pip says 'externally-managed-environment', add --break-system-packages.\n"
        "If you have more than one Python installed, a plain 'pip install' often\n"
        "lands in a different one - the command above always targets the Python\n"
        "that is running this script."
        % (sys.executable, sys.version.split()[0], sys.executable)
    )

# ----------------------------------------------------------------------------
# Optional alias table: different names that mean the same column.
# Extend it with --alias my_alias.json  {"customer name": ["客户名称", "client"], ...}
# ----------------------------------------------------------------------------
BUILTIN_ALIAS_GROUPS = [
    ["name", "fullname", "customername", "clientname", "姓名", "名字", "客户名称", "客户姓名", "联系人"],
    ["email", "mail", "emailaddress", "邮箱", "电子邮箱", "电子邮件"],
    ["phone", "mobile", "tel", "telephone", "手机", "手机号", "电话", "联系电话"],
    ["date", "day", "orderdate", "createdat", "created", "dt", "日期", "下单日期", "日期时间", "订货日期"],
    ["amount", "total", "sum", "price", "金额", "合计", "总额", "总金额"],
    ["orderid", "orderno", "order", "订单号", "单号", "编号"],
    ["product", "item", "sku", "产品", "商品", "品名"],
    ["qty", "quantity", "num", "count", "数量", "件数"],
    ["region", "area", "city", "地区", "区域", "城市"],
    ["note", "remark", "comment", "memo", "备注", "说明"],
]

# Guess when a normalized header is a typo / variant of an existing one.
FUZZY_THRESHOLD = 0.86


def norm(s):
    """Normalize a header for matching: lowercase, drop spaces / punctuation / brackets."""
    if s is None:
        return ""
    s = str(s).replace("\ufeff", "").strip().lower()
    s = re.sub(r"[\s_\-/\\.,:;、（）()\[\]【】{}:：·'\"`~!@#$%^&*+=|<>?]+", "", s)
    return s


def build_alias_map(extra_path=None):
    """normalized-name -> canonical-key"""
    groups = [list(g) for g in BUILTIN_ALIAS_GROUPS]
    if extra_path:
        import json
        with open(extra_path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        for canon, variants in raw.items():
            groups.append([canon] + list(variants))
    m = {}
    for g in groups:
        keys = [norm(x) for x in g]
        canon = keys[0]
        for k in keys:
            if k:
                m.setdefault(k, canon)
    return m


def is_blank_row(vals):
    for v in vals:
        if v is None:
            continue
        if isinstance(v, str) and v.strip() == "":
            continue
        return False
    return True


def read_csv(path):
    """Return (headers, rows). Tries utf-8-sig, then gbk, then latin-1."""
    for enc in ("utf-8-sig", "utf-8", "gbk", "latin-1"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                sample = f.read(4096)
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
                except Exception:
                    dialect = csv.excel
                table = list(csv.reader(f, dialect))
            if not table:
                return [], []
            return table[0], table[1:]
        except UnicodeDecodeError:
            continue
        except Exception:
            raise
    raise IOError("cannot decode CSV: %s" % path)


def _xls_cell_value(cell, datemode):
    """Convert one xlrd cell into a plain Python value (str / int / float / datetime)."""
    import xlrd
    t, v = cell.ctype, cell.value
    if t in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return None
    if t == xlrd.XL_CELL_DATE:
        try:
            return xlrd.xldate.xldate_as_datetime(v, datemode)
        except Exception:
            return v
    if t == xlrd.XL_CELL_BOOLEAN:
        return bool(v)
    if t == xlrd.XL_CELL_NUMBER:
        return int(v) if float(v).is_integer() else v
    if t == xlrd.XL_CELL_ERROR:
        return None
    return v


def read_xls_sheets(path):
    """Yield (sheet_name, 2D table) for a legacy .xls workbook. Needs xlrd."""
    try:
        import xlrd
    except ImportError:
        raise RuntimeError(
            "读旧版 .xls 需要 xlrd：pip install xlrd　"
            "（或先用 Excel/WPS 打开后『另存为 .xlsx』再跑）"
        )
    try:
        book = xlrd.open_workbook(path, on_demand=True)
    except Exception as e:
        raise RuntimeError("这个 .xls 打不开（可能不是真正的 Excel 97-2003 文件）：%s" % e)
    try:
        for sh in book.sheets():
            table = [[_xls_cell_value(sh.cell(r, c), book.datemode) for c in range(sh.ncols)]
                     for r in range(sh.nrows)]
            yield sh.name, table
    finally:
        try:
            book.release_resources()
        except Exception:
            pass


def classify_sheets(pairs, sheet_opt, all_sheets, first_sheet=False, skip_names=()):
    """Wrapper: honour --skip-sheet before the automatic rules run."""
    forced = set(skip_names or ())
    pre = []
    if forced:
        pre = [(n, "skipped by request (--skip-sheet)")
               for n, _ in pairs if n in forced]
        pairs = [(n, t) for n, t in pairs if n not in forced]
    picked, skipped = _classify_sheets(pairs, sheet_opt, all_sheets, first_sheet)
    return picked, pre + skipped


def _classify_sheets(pairs, sheet_opt, all_sheets, first_sheet=False):
    """Decide which sheets are data and which are skipped, with a reason for each.

    pairs: [(sheet_name, 2D table)] in workbook order.
    Returns: ([(name, table)] to read, [(name, reason)] skipped).

    Default behaviour reads EVERY data sheet, not just the first one. A monthly
    workbook (1月..12月 in one file) used to lose 11 of 12 months without a word.
    """
    names = [n for n, _ in pairs]

    if sheet_opt:
        picked = [(n, t) for n, t in pairs if n == sheet_opt]
        if not picked:
            raise RuntimeError("这个文件里没有名为 '%s' 的工作表（有：%s）"
                               % (sheet_opt, ", ".join(names)))
        skipped = [(n, "not selected (--sheet)") for n in names if n != sheet_opt]
        return picked, skipped

    kinds = []
    for n, t in pairs:
        if not t or all(is_blank_row(r) for r in t):
            kind = "empty"
        elif looks_like_cover(t):
            kind = "cover"
        elif is_summary_sheet(n):
            kind = "summary"
        else:
            kind = "data"
        kinds.append((n, t, kind))

    data = [(n, t) for n, t, k in kinds if k == "data"]

    if all_sheets:
        keep_names = [n for n, t, k in kinds if k in ("data", "summary")]
        if not keep_names:
            keep_names = [n for n, t, k in kinds if k != "empty"]
        if first_sheet and keep_names:
            keep_names = keep_names[:1]
        keep = [(n, t) for n, t, k in kinds if n in keep_names]
        reason_of = {"empty": "empty", "cover": "cover / notes page",
                     "summary": "summary sheet (read because of --all-sheets)", "data": ""}
        skipped = [(n, reason_of[k]) for n, t, k in kinds if n not in keep_names]
        return keep, skipped

    if data:
        keep = data[:1] if first_sheet else data
    else:
        # No page looks like a table. Fall back to a summary-ish page (a one-sheet
        # workbook is often named 汇总 but still holds the rows), never to a cover
        # page - merging a title line as data is worse than reading nothing.
        keep = [(n, t) for n, t, k in kinds if k not in ("empty", "cover")][:1]
    keep_names = [n for n, _ in keep]
    reason_of = {"empty": "empty",
                 "cover": "cover / notes page",
                 "summary": "summary sheet (skipped, would double-count)",
                 "data": "not read (--first-sheet)"}
    skipped = [(n, reason_of[k]) for n, t, k in kinds if n not in keep_names]
    return keep, skipped


def iter_xls(path, sheet_opt, all_sheets, header_row=None, first_sheet=False, skip_sheets=()):
    """Yield (sheet_label, headers, data_rows, note, meta) for a legacy .xls file."""
    sheets = list(read_xls_sheets(path))
    if not sheets:
        return
    picked, skipped = classify_sheets(sheets, sheet_opt, all_sheets, first_sheet, skip_sheets)
    for name, reason in skipped:
        yield name, None, [], "skipped: " + reason, None
    for name, table in picked:
        headers, rows = pick_header(table, header_row)
        if headers is None:
            yield name, None, [], "skipped: empty sheet", None
            continue
        yield name, headers, rows, "", None


def row_width(r):
    n = 0
    for v in r:
        if v is None:
            continue
        if isinstance(v, str) and v.strip() == "":
            continue
        n += 1
    return n


SUMMARY_WORDS = ("汇总", "合计", "总计", "小计", "统计", "summary", "total", "pivot", "dashboard")


def is_summary_sheet(name):
    """Sheet names that usually hold derived numbers, not source rows."""
    n = norm(name)
    return any(norm(w) in n for w in SUMMARY_WORDS)


def looks_like_cover(table):
    """True for a title / notes / instruction page - something that is NOT a table.

    Real workbooks very often start with a cover sheet:
        A1: 2026年3月对账单      A3: 制表：李四    A4: 电话：138...
    Reading that as data silently produces one garbage row and zero real rows.

    Careful with false positives: a 500-row single-column list of emails IS data,
    and so is a 2-row 2-column table. Only a few sparse rows with a sentence in
    them count as a cover.
    """
    rows = [r for r in table if not is_blank_row(r)]
    if not rows:
        return True
    widths = [row_width(r) for r in rows]
    maxw = max(widths)
    if maxw <= 1:
        return len(rows) <= 6          # one column: a title/notes block, not a list
    cells = sum(widths)
    long_text = 0
    for r in rows:
        for v in r:
            if isinstance(v, str) and len(v.strip()) >= 12:
                long_text += 1
    if len(rows) <= 3 and cells <= 5 and long_text >= 1:
        return True
    return False


def header_score(row):
    """Score how much a row looks like a header rather than a title or a data row.

    Real headers are SHORT labels: 姓名 / 金额 / Order ID.
    Info lines above them look like: 单位：元 / 期间：2026年1-3月 / 制表：李四.
    Both are text, so length and the 'key：value' shape are what tell them apart.
    """
    vals = [v for v in row if v is not None and (not isinstance(v, str) or v.strip() != "")]
    if len(vals) < 2:
        return -1.0
    short = mid = numeric = shaped = 0
    for v in vals:
        if isinstance(v, (datetime, date, time)):
            numeric += 1
            continue
        if isinstance(v, bool):
            short += 1
            continue
        if isinstance(v, (int, float)):
            numeric += 1
            continue
        s = str(v).strip()
        if re.fullmatch(r"[-+]?\d+(\.\d+)?", s):
            numeric += 1
            continue
        if re.search(r"[:：=]", s):
            shaped += 1          # 单位：元 -> an info line, not a column name
        if len(s) <= 8:
            short += 1
        elif len(s) <= 24:
            mid += 1
    n = float(len(vals))
    return (short * 1.0 + mid * 0.3 - 0.8 * numeric - 0.5 * shaped) / n


def auto_header_index(table, scan=10):
    """Find the header row.

    Real files carry a title / 'exported from ERP' / unit line above the header.
    Old rule (first row at least half as wide as the widest) breaks when that
    info line has 4 cells and the header has 5 - the info line wins and every
    column name becomes garbage.

    Now: among the first `scan` rows wide enough to be the header, take the one
    that scores most like labels (short text, few numbers). Ties go to the
    earliest row, so a normal file still resolves to row 1.
    """
    head = table[:scan]
    if not head:
        return 0
    widths = [row_width(r) for r in head]
    maxw = max(widths)
    if maxw <= 1:
        for i, w in enumerate(widths):
            if w >= 1:
                return i
        return 0
    need = max(2, maxw * 0.5)
    best_i, best_s = -1, -99.0
    for i, r in enumerate(head):
        if widths[i] < need:
            continue
        s = header_score(r)
        if s > best_s + 0.15:      # clearly better -> take it
            best_s, best_i = s, i
    if best_i >= 0:
        return best_i
    for i, w in enumerate(widths):
        if w >= need:
            return i
    return 0


def pick_header(table, header_row):
    """Split a raw 2D table into (headers, data_rows).

    If header_row is given (1-based, physical), that row is the header.
    Otherwise it is auto-detected.
    """
    if not table:
        return None, []
    if header_row:
        idx = header_row - 1
        if idx >= len(table):
            return None, []
        return list(table[idx]), [list(r) for r in table[idx + 1:]]
    idx = auto_header_index(table)
    return list(table[idx]), [list(r) for r in table[idx + 1:]]


def is_repeated_header(row, headers):
    """True when a data row is the header row printed again.

    Files exported one month per page often repeat the header inside the data
    block, so the merged sheet ends up with 'Name | Email' sitting in the data.
    """
    a = [norm(v) for v in row if not (v is None or (isinstance(v, str) and v.strip() == ""))]
    b = [norm(v) for v in headers if not (v is None or (isinstance(v, str) and v.strip() == ""))]
    return len(a) >= 2 and a == b


# ---------------------------------------------------------------------------
# Merged cells and number formats
#
# Both are invisible to a values-only read:
#   B2:B4 merged = "华东"  -> only B2 has the value, B3/B4 come through as None
#   order id 123 with format "0000000" -> reads as 123, and the leading zeros are
#   gone from the output even though the source file displays 0000123.
# Neither shows up in a row-count check, so both need explicit handling.
# ---------------------------------------------------------------------------
_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
       "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


def sheet_merge_refs(path):
    """{sheet name: ['B2:B4', ...]} read straight from the sheet XML.

    openpyxl's read_only mode does not expose merged_cells, and loading the
    workbook normally would defeat the whole point of streaming. Reading the zip
    entry directly is cheap and exact.
    """
    import zipfile
    import xml.etree.ElementTree as ET

    out = {}
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            if "xl/workbook.xml" not in names:
                return out
            rels = {}
            if "xl/_rels/workbook.xml.rels" in names:
                for rel in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")):
                    rels[rel.get("Id")] = rel.get("Target")
            sheets = ET.fromstring(z.read("xl/workbook.xml")).find("m:sheets", _NS)
            if sheets is None:
                return out
            for sh in sheets:
                target = rels.get(sh.get("{%s}id" % _NS["r"]), "")
                if not target:
                    continue
                target = target.lstrip("/")
                if not target.startswith("xl/"):
                    target = "xl/" + target
                if target not in names:
                    continue
                root = ET.fromstring(z.read(target))
                mc = root.find("m:mergeCells", _NS)
                refs = [c.get("ref") for c in mc if c.get("ref")] if mc is not None else []
                if refs:
                    out[sh.get("name")] = refs
    except Exception:
        return out
    return out


def fill_merged(table, refs, header_idx):
    """Copy each merged block's top-left value across the block. Returns how many.

    A merged range means "these cells are the same value" - Excel just shows it
    once. Leaving the rest as None is how a region column ends up empty from row 2
    down. Ranges that touch the header (a title spanning columns) are left alone:
    filling those would invent columns.
    """
    from openpyxl.utils import range_boundaries

    filled = 0
    for ref in refs:
        try:
            min_col, min_row, max_col, max_row = range_boundaries(ref)
        except Exception:
            continue
        if min_row - 1 <= header_idx:        # inside / above the header row
            continue
        r0, c0 = min_row - 1, min_col - 1
        if r0 >= len(table) or c0 >= len(table[r0]):
            continue
        val = table[r0][c0]
        if val is None:
            continue
        for r in range(r0, min(max_row, len(table))):
            row = table[r]
            if len(row) < max_col:
                row.extend([None] * (max_col - len(row)))
            for c in range(c0, max_col):
                if row[c] is None:
                    row[c] = val
                    filled += 1
    return filled


_DATE_FMT = re.compile(r"[ymdhs]", re.I)
_DIGIT_PLACE = re.compile(r"[0#]")


def pick_number_format(fmts):
    """Most common display format worth carrying over, or None.

    Carried: 0000000 (an order id must keep its zeros), #,##0.00, ¥#,##0.00,
    0.0% - formats that change what the reader SEES.
    Not carried: date/time formats. Writing "yyyy-mm-dd" on a column makes every
    reader (openpyxl included) treat the stored serial 46004 as a date, which
    means merge would quietly start interpreting values. Unifying dates is
    excel_clean.py's job, and its _log says so. Merge only concatenates.
    """
    if not fmts:
        return None
    for fmt, cnt in fmts.most_common():
        if not fmt or fmt == "General":
            continue
        if "@" in fmt or len(fmt) > 60:
            continue          # text format / pathological custom format
        if _DATE_FMT.search(fmt) and not _DIGIT_PLACE.search(fmt):
            continue          # date or time format: leave to clean
        return fmt
    return None


def looks_numeric_text(v):
    """A string that is really a number wearing decoration: '¥1,234', ' 88 ', '1,234.50'."""
    if not isinstance(v, str):
        return False
    s = v.strip()
    if not s:
        return False
    for ch in (",", "¥", "$", "￥", "元", " ", "\u00a0"):
        s = s.replace(ch, "")
    try:
        float(s)
        return True
    except Exception:
        return False


def mixed_type_columns(out_rows, canonical, display, sample=3000):
    """Columns holding both real numbers and number-looking text.

    This is the quiet one: Excel's SUM, pivot tables and date filters skip text
    cells without a word, so the customer's total comes out low while every cell
    looks fine. Merge cannot fix it (it never rewrites a value) - but it can say
    so, and excel_clean.py can coerce the text.
    """
    found = []
    rows = out_rows[:sample]
    for k in canonical:
        if str(k).startswith("_"):
            continue
        n_num = n_txt = 0
        for rec in rows:
            v = rec.get(k)
            if v is None or isinstance(v, bool):
                continue
            if isinstance(v, (int, float, datetime, date)):
                n_num += 1
            elif looks_numeric_text(v):
                n_txt += 1
        if n_num and n_txt:
            found.append((display.get(k, k), n_num, n_txt))
    return found


def iter_sheets(path, sheet_opt, all_sheets, header_row=None, first_sheet=False,
                skip_sheets=(), fill_merged_cells=True):
    """Yield (sheet_label, headers, data_rows, note, meta) for one file.

    meta = {"fmt": {col_index: Counter(number_format)}, "merged": n_filled} for
    .xlsx, or None when there is nothing to carry (csv / legacy .xls).
    """
    ext = os.path.splitext(path)[1].lower()

    if ext in (".csv", ".txt"):
        h, rows = read_csv(path)
        table = ([] if not h else [h] + rows)
        hdr, rows = pick_header(table, header_row)
        yield "(csv)", hdr, rows, "", None
        return

    if ext == ".xls":
        yield from iter_xls(path, sheet_opt, all_sheets, header_row, first_sheet, skip_sheets)
        return

    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        names = list(wb.sheetnames)
        # Cheap probe first (300 rows per sheet) so a 12-sheet workbook is not
        # loaded twice and a broken "first sheet only" rule cannot hide months.
        probe = []
        for n in names:
            ws = wb[n]
            head = []
            for i, r in enumerate(ws.iter_rows(values_only=True)):
                head.append(list(r))
                if i >= 299:
                    break
            probe.append((n, head))
        picked, skipped = classify_sheets(probe, sheet_opt, all_sheets, first_sheet, skip_sheets)

        for n, reason in skipped:
            yield n, None, [], "skipped: " + reason, None

        merges = sheet_merge_refs(path) if fill_merged_cells else {}
        for n, _ in picked:
            ws = wb[n]
            # Values: read raw, exactly as stored. A date kept as serial 46004 stays
            # 46004 here - interpreting it is excel_clean.py's job, and doing it in
            # both places would make the audit log lie about where a change happened.
            table = [list(r) for r in ws.iter_rows(values_only=True)]
            # Formats: a second pass over the first 200 rows (cheap, and a column's
            # display format is almost always uniform). Read via Cell objects because
            # values_only drops the format entirely.
            fmts = {}
            ws_fmt = wb[n]
            for ri, row in enumerate(ws_fmt.iter_rows()):
                if ri >= 200:
                    break
                for ci, c in enumerate(row):
                    nf = c.number_format
                    if nf and nf != "General":
                        fmts.setdefault(ci, Counter())[nf] += 1
            if not table:
                yield n, None, [], "skipped: empty sheet", None
                continue
            hdr_idx = (header_row - 1) if header_row else auto_header_index(table)
            if hdr_idx >= len(table) or hdr_idx < 0:
                yield n, None, [], "skipped: empty sheet", None
                continue
            merged_n = fill_merged(table, merges.get(n, []), hdr_idx) if merges else 0
            headers = list(table[hdr_idx])
            rows = [list(r) for r in table[hdr_idx + 1:]]
            yield n, headers, rows, "", {"fmt": fmts, "merged": merged_n}
    finally:
        wb.close()


def main():
    ap = argparse.ArgumentParser(
        description="Merge every .xlsx/.csv in a folder into one sheet. Values are never modified.")
    ap.add_argument("--src", required=True, help="folder containing the files")
    ap.add_argument("--out", default=None, help="output .xlsx (default: merged_<timestamp>.xlsx)")
    ap.add_argument("--files", nargs="*", default=None,
                    help="only these file names inside --src (default: all)")
    ap.add_argument("--sheet", default=None, help="read this sheet name from every workbook")
    ap.add_argument("--all-sheets", action="store_true",
                    help="read every sheet including pages named like a summary")
    ap.add_argument("--first-sheet", action="store_true",
                    help="read only the first data sheet of each file (old default)")
    ap.add_argument("--skip-sheet", action="append", default=None, metavar="NAME",
                    help="never read a sheet with this name (repeatable)")
    ap.add_argument("--strict", action="store_true",
                    help="exit with code 2 if any sheet was skipped, so a cron job notices")
    ap.add_argument("--no-fill-merged", action="store_true",
                    help="do not copy a merged cell's value across its block")
    ap.add_argument("--header-row", type=int, default=None,
                    help="1-based row index of the header (default: detected)")
    ap.add_argument("--alias", default=None, help="extra alias json, e.g. {\"name\": [\"客户名称\"]}")
    ap.add_argument("--no-source", action="store_true", help="do not add _source_file / _source_sheet columns")
    ap.add_argument("--fuzzy", type=float, default=FUZZY_THRESHOLD, help="header similarity threshold 0-1")
    ap.add_argument("--preview", action="store_true", help="print the first 3 rows of the result")
    args = ap.parse_args()

    src = os.path.abspath(args.src)
    if not os.path.isdir(src):
        sys.exit("Not a folder: %s" % src)

    exts = (".xlsx", ".xlsm", ".csv", ".txt", ".xls")
    if args.files:
        files = [os.path.join(src, f) for f in args.files]
        missing = [f for f in files if not os.path.exists(f)]
        if missing:
            sys.exit("File not found: %s" % ", ".join(missing))
    else:
        files = sorted(
            os.path.join(src, f) for f in os.listdir(src)
            if f.lower().endswith(exts) and not f.startswith("~$")
            and os.path.isfile(os.path.join(src, f))
        )
    if not files:
        sys.exit("No .xlsx / .csv found in %s" % src)

    print("=" * 68)
    print("Input folder : %s" % src)
    print("Files found  : %d" % len(files))
    print("=" * 68)

    alias_map = build_alias_map(args.alias)
    canonical = []          # ordered list of canonical column keys
    display = {}            # canonical key -> header text to show
    seen_variants = {}      # canonical key -> set of raw headers seen
    col_fmt = {}            # canonical key -> Counter(number_format)
    aliased_log = []
    split_log = []          # (file, raw header, the key it wanted to share)
    out_rows = []
    report = []             # (file, sheet, read, blank, dup, written, note)
    skipped_sheets = []     # (file, sheet, reason) - for the report and --strict
    failed = []
    merged_filled = 0

    for path in files:
        fname = os.path.basename(path)
        try:
            for sheet_label, headers, rows, note, meta in iter_sheets(
                    path, args.sheet, args.all_sheets, args.header_row,
                    args.first_sheet, args.skip_sheet or (), not args.no_fill_merged):
                if headers is None:
                    report.append((fname, sheet_label, 0, 0, 0, 0, note or "empty sheet, skipped"))
                    skipped_sheets.append((fname, sheet_label, note or "empty sheet, skipped"))
                    continue
                fmts = (meta or {}).get("fmt") or {}
                merged_n = (meta or {}).get("merged") or 0
                merged_filled += merged_n

                # map raw header index -> canonical key
                col_of = {}
                for i, h in enumerate(headers):
                    raw = "" if h is None else str(h).strip()
                    if raw == "":
                        if any(r[i] not in (None, "") for r in rows if i < len(r)):
                            key = "未命名列_%d" % (i + 1)
                        else:
                            continue
                    else:
                        key = alias_map.get(norm(h), norm(h))
                        if key not in canonical and canonical:
                            hit = difflib.get_close_matches(key, canonical, n=1, cutoff=args.fuzzy)
                            if hit:
                                aliased_log.append((fname, raw, hit[0]))
                                key = hit[0]

                    # Two columns of the SAME sheet on one key. "金额" and "合计"
                    # are both in the built-in alias group, so the second column
                    # would overwrite the first cell by cell: values change, every
                    # row count still matches, and nothing warns. Split instead.
                    if key in col_of.values():
                        base = key
                        suffix = norm(raw) or ("col%d" % (i + 1))
                        key = "%s__%s" % (base, suffix)
                        n = 2
                        while key in col_of.values() or key in canonical:
                            key = "%s__%s_%d" % (base, suffix, n)
                            n += 1
                        split_log.append((fname, raw or ("column %d" % (i + 1)),
                                          display.get(base, base)))
                        display[key] = raw or ("column %d" % (i + 1))

                    if key not in canonical:
                        canonical.append(key)
                        display.setdefault(key, raw if raw else key)
                    seen_variants.setdefault(key, set()).add(str(h))
                    col_of[i] = key
                    if i in fmts:
                        col_fmt.setdefault(key, Counter()).update(fmts[i])

                written = blank = dup = 0
                for r in rows:
                    if is_blank_row(r):
                        blank += 1
                        continue
                    if is_repeated_header(r, headers):
                        dup += 1
                        continue
                    # rows wider than the header: keep the extras instead of losing them
                    rec = {}
                    for i, v in enumerate(r):
                        if i in col_of:
                            rec[col_of[i]] = v
                        else:
                            extra = "未命名列_%d" % (i + 1)
                            if extra not in canonical:
                                canonical.append(extra)
                                display[extra] = extra
                            rec[extra] = v
                    if not args.no_source:
                        rec["_source_file"] = fname
                        rec["_source_sheet"] = sheet_label
                    out_rows.append(rec)
                    written += 1
                bits = []
                if dup:
                    bits.append("repeated header rows dropped: %d" % dup)
                if merged_n:
                    bits.append("merged cells filled: %d" % merged_n)
                report.append((fname, sheet_label, len(rows), blank, dup, written, "; ".join(bits)))
        except Exception as e:
            failed.append((fname, str(e)))
            report.append((fname, "-", 0, 0, 0, 0, "FAILED: %s" % e))
            print("  [!] %s -> %s" % (fname, e))
            if os.environ.get("MERGE_DEBUG"):
                traceback.print_exc()

    if not args.no_source:
        for k in ("_source_file", "_source_sheet"):
            canonical.append(k)
        display["_source_file"] = "Source file"
        display["_source_sheet"] = "Source sheet"

    if not out_rows:
        sys.exit("\nNothing was read. Check the files above. Nothing written.")

    # ---------------- write output ----------------
    out = args.out or ("merged_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M"))
    out = os.path.abspath(out)
    wb = Workbook()
    ws = wb.active
    ws.title = "Merged"

    ws.append([display.get(k, k) for k in canonical])
    for rec in out_rows:
        ws.append([rec.get(k) for k in canonical])

    # keep dates readable and column widths sane
    for ci, key in enumerate(canonical, start=1):
        letter = get_column_letter(ci)
        width = min(max(10, len(str(display.get(key, key))) + 4), 40)
        ws.column_dimensions[letter].width = width
        # Carry the source display format: an order id stored as 123 with format
        # "0000000" must still read 0000123, and a currency column must not turn
        # into a bare 1234. The value itself is never touched.
        fmt = pick_number_format(col_fmt.get(key))
        if fmt:
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=ci).number_format = fmt
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for c in row:
            if not isinstance(c.value, (datetime, date)):
                continue
            midnight = (not isinstance(c.value, datetime)) or (
                c.value.hour == 0 and c.value.minute == 0 and c.value.second == 0)
            fmt = c.number_format or "General"
            if midnight and re.search(r"[hHsS]", fmt):
                c.number_format = "yyyy-mm-dd"      # a date, not a timestamp at 00:00
            elif fmt == "General":
                c.number_format = "yyyy-mm-dd" if midnight else "yyyy-mm-dd hh:mm"
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(canonical)), ws.max_row)

    # ---------------- report sheet ----------------
    rs = wb.create_sheet("_report")
    rs.append(["File", "Sheet", "Rows read", "Blank rows skipped",
               "Repeated header rows dropped", "Rows written", "Note"])
    total_read = total_blank = total_dup = total_written = 0
    for row in report:
        rs.append(list(row))
        total_read += row[2]
        total_blank += row[3]
        total_dup += row[4]
        total_written += row[5]
    rs.append([])
    rs.append(["TOTAL", "", total_read, total_blank, total_dup, total_written, ""])

    accounted = total_read - total_blank - total_dup
    if accounted == total_written:
        check = "OK" if not skipped_sheets else "OK (%d sheet(s) not read - see below)" % len(skipped_sheets)
    else:
        check = "MISMATCH: read-blank-dup=%d but wrote %d" % (accounted, total_written)
    rs.append(["CHECK", "", "", "", "", check, ""])

    if skipped_sheets:
        rs.append([])
        rs.append(["SHEETS NOT READ", "", "", "", "", "", ""])
        for f, sname, why in skipped_sheets:
            rs.append([f, sname, "", "", "", "", why])

    for letter, w in (("A", 34), ("B", 18), ("C", 11), ("D", 20), ("E", 26), ("F", 14), ("G", 46)):
        rs.column_dimensions[letter].width = w

    if aliased_log:
        als = wb.create_sheet("_aliases")
        als.append(["File", "Header in file", "Matched to"])
        for a in aliased_log:
            als.append(list(a))
        als.column_dimensions["A"].width = 34
        als.column_dimensions["B"].width = 26
        als.column_dimensions["C"].width = 26

    wb.save(out)

    # ---------------- console summary ----------------
    print()
    print("%-34s %7s %7s %7s %8s" % ("file", "read", "blank", "dup", "written"))
    print("-" * 68)
    for row in report:
        print("%-34s %7d %7d %7d %8d" % (row[0][:34], row[2], row[3], row[4], row[5]))
    print("-" * 68)
    print("%-34s %7d %7d %7d %8d" % ("TOTAL", total_read, total_blank, total_dup, total_written))

    if skipped_sheets:
        print("\nSheets NOT read (%d) - each one is listed so nothing disappears silently:"
              % len(skipped_sheets))
        for f, sname, why in skipped_sheets:
            print("  - %-30s / %-18s %s" % (f[:30], str(sname)[:18], why))

    print("\nColumns (%d): %s" % (len(canonical), ", ".join(display.get(k, k) for k in canonical)))

    if aliased_log:
        print("\nHeaders auto-aligned (check these):")
        for f, h, k in aliased_log:
            print("  - %-28s %-20s -> %s" % (f[:28], h, k))

    dup_headers = {k: v for k, v in seen_variants.items() if len(v) > 1}
    if dup_headers:
        print("\nSame column, different spellings (merged into one):")
        for k, vs in dup_headers.items():
            print("  - %s  <=  %s" % (display.get(k, k), " | ".join(sorted(vs))))

    if split_log:
        print("\nTwo columns wanted the same name - kept apart instead of overwriting:")
        for f, h, base in split_log:
            print("  - %-24s %-20s would have overwritten '%s'" % (f[:24], h[:20], base))

    if merged_filled:
        print("\nMerged cells: %d empty cells filled with their block's value "
              "(a merged region means one value, not a blank one)." % merged_filled)
        print("             Turn this off with --no-fill-merged if the blanks were meant to stay blank.")

    mixed = mixed_type_columns(out_rows, canonical, display)
    if mixed:
        print("\nMixed types inside one column (Excel's SUM silently skips the text cells):")
        for name, nn, nt in mixed:
            print("  - %-24s %d numbers, %d text values   -> excel_clean.py can coerce them"
                  % (str(name)[:24], nn, nt))

    if failed:
        print("\nFAILED files (%d):" % len(failed))
        for f, e in failed:
            print("  - %s: %s" % (f, e))

    if args.preview and out_rows:
        print("\nFirst 3 rows:")
        keys = [k for k in canonical]
        for rec in out_rows[:3]:
            print("  " + " | ".join("%s=%s" % (display.get(k, k), rec.get(k)) for k in keys[:6]))

    print("\nSaved -> %s" % out)
    print("Check the '_report' sheet: read = blank + repeated header + written, for every line.")
    if skipped_sheets:
        print("%d sheet(s) were not read. If one of them actually holds data, re-run with --all-sheets."
              % len(skipped_sheets))
    if args.strict and (skipped_sheets or failed):
        sys.exit(2)


if __name__ == "__main__":
    main()
