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


def iter_xls(path, sheet_opt, all_sheets, header_row=None):
    """Yield (sheet_label, headers, data_rows) for a legacy .xls file."""
    sheets = list(read_xls_sheets(path))
    if not sheets:
        return
    if sheet_opt:
        picked = [(n, t) for n, t in sheets if n == sheet_opt]
        if not picked:
            raise RuntimeError("这个文件里没有名为 '%s' 的工作表（有：%s）"
                               % (sheet_opt, ", ".join(n for n, _ in sheets)))
    elif all_sheets:
        picked = sheets
    else:
        # same rule as .xlsx: first sheet that actually has more than a header row
        picked = [sheets[0]]
        for n, t in sheets:
            if sum(1 for r in t if not is_blank_row(r)) > 1:
                picked = [(n, t)]
                break
    for name, table in picked:
        headers, rows = pick_header(table, header_row)
        if headers is None:
            continue
        yield name, headers, rows


def row_width(r):
    n = 0
    for v in r:
        if v is None:
            continue
        if isinstance(v, str) and v.strip() == "":
            continue
        n += 1
    return n


def auto_header_index(table, scan=10):
    """Find the header row.

    Real files often carry a title / 'exported from ERP' line above the header.
    Rule: inside the first `scan` rows, take the FIRST row that is at least half as
    wide as the widest of them. That skips a lone title line but still accepts a
    header whose first cells are blank.
    """
    head = table[:scan]
    widths = [row_width(r) for r in head]
    maxw = max(widths) if widths else 0
    if maxw <= 1:
        for i, w in enumerate(widths):
            if w >= 1:
                return i
        return 0
    need = max(2, maxw * 0.5)
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


def iter_sheets(path, sheet_opt, all_sheets, header_row=None):
    """Yield (sheet_label, headers, data_rows) for one file."""
    ext = os.path.splitext(path)[1].lower()

    if ext in (".csv", ".txt"):
        h, rows = read_csv(path)
        table = ([] if not h else [h] + rows)
        hdr, rows = pick_header(table, header_row)
        yield "(csv)", hdr, rows
        return

    if ext == ".xls":
        yield from iter_xls(path, sheet_opt, all_sheets, header_row)
        return

    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        if all_sheets:
            targets = wb.worksheets
        elif sheet_opt:
            if sheet_opt not in wb.sheetnames:
                raise RuntimeError("这个文件里没有名为 '%s' 的工作表（有：%s）"
                                   % (sheet_opt, ", ".join(wb.sheetnames)))
            targets = [wb[sheet_opt]]
        else:
            # first sheet that actually has content
            targets = [wb.worksheets[0]]
            for ws in wb.worksheets:
                if ws.max_row and ws.max_row > 1:
                    targets = [ws]
                    break

        for ws in targets:
            table = [list(r) for r in ws.iter_rows(values_only=True)]
            headers, rows = pick_header(table, header_row)
            if headers is None:
                continue
            yield ws.title, headers, rows
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
    ap.add_argument("--all-sheets", action="store_true", help="read every sheet, not just the first")
    ap.add_argument("--header-row", type=int, default=None,
                    help="1-based row index of the header (default: first non-blank row)")
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
    aliased_log = []
    out_rows = []
    report = []             # (file, sheet, read, blank_skipped, written, note)
    dropped_cols = set()
    failed = []

    for path in files:
        fname = os.path.basename(path)
        try:
            for sheet_label, headers, rows in iter_sheets(path, args.sheet, args.all_sheets, args.header_row):
                if headers is None:
                    report.append((fname, sheet_label, 0, 0, 0, "empty sheet, skipped"))
                    continue
                # map raw header index -> canonical key
                col_of = {}
                for i, h in enumerate(headers):
                    if h is None or str(h).strip() == "":
                        if any(r[i] not in (None, "") for r in rows if i < len(r)):
                            key = "未命名列_%d" % (i + 1)
                        else:
                            continue
                    else:
                        key = alias_map.get(norm(h), norm(h))
                        if key not in canonical and canonical:
                            hit = difflib.get_close_matches(key, canonical, n=1, cutoff=args.fuzzy)
                            if hit:
                                aliased_log.append((fname, str(h), hit[0]))
                                key = hit[0]
                    if key not in canonical:
                        canonical.append(key)
                        display[key] = str(h).strip() if h not in (None, "") else key
                    seen_variants.setdefault(key, set()).add(str(h))
                    col_of[i] = key

                written = blank = 0
                for r in rows:
                    if is_blank_row(r):
                        blank += 1
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
                report.append((fname, sheet_label, len(rows), blank, written, ""))
        except Exception as e:
            failed.append((fname, str(e)))
            report.append((fname, "-", 0, 0, 0, "FAILED: %s" % e))
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
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for c in row:
            if isinstance(c.value, (datetime, date)):
                c.number_format = "yyyy-mm-dd" if not isinstance(c.value, datetime) or (
                    c.value.hour == 0 and c.value.minute == 0) else "yyyy-mm-dd hh:mm"
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(canonical)), ws.max_row)

    # ---------------- report sheet ----------------
    rs = wb.create_sheet("_report")
    rs.append(["File", "Sheet", "Rows read", "Blank rows skipped", "Rows written", "Note"])
    total_read = total_blank = total_written = 0
    for row in report:
        rs.append(list(row))
        total_read += row[2]
        total_blank += row[3]
        total_written += row[4]
    rs.append([])
    rs.append(["TOTAL", "", total_read, total_blank, total_written, ""])
    rs.append(["CHECK", "", "", "", "OK" if total_written == sum(1 for _ in out_rows) else "MISMATCH", ""])
    for letter, w in (("A", 34), ("B", 18), ("C", 12), ("D", 20), ("E", 14), ("F", 50)):
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
    print("%-38s %8s %8s %8s" % ("file", "read", "blank", "written"))
    print("-" * 68)
    for row in report:
        print("%-38s %8d %8d %8d" % (row[0][:38], row[2], row[3], row[4]))
    print("-" * 68)
    print("%-38s %8d %8d %8d" % ("TOTAL", total_read, total_blank, total_written))

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
    print("Check the '_report' sheet: rows read must equal rows written.")


if __name__ == "__main__":
    main()
