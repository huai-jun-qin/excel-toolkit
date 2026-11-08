#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_split.py - Split one sheet into many, by a column value (or into fixed-size chunks).

Two modes:
    --mode files   one .xlsx per group   (default)
    --mode sheets  one sheet per group, all inside a single .xlsx

Requires: openpyxl  (pip install openpyxl)
Optional: xlrd      (pip install xlrd)   -- only needed to read legacy .xls files

Usage:
    python3 excel_split.py --src all.xlsx --by Region
    python3 excel_split.py --src all.xlsx --by 3 --mode sheets --out groups.xlsx
    python3 excel_split.py --src all.xlsx --chunk 500
"""

import argparse
import os
import re
import sys
from datetime import date, datetime

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Missing dependency. Run:  pip install openpyxl")


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


def load_table(src, sheet_opt):
    """Read a single sheet into a 2D list of plain values, from .xlsx or legacy .xls.

    Returns (sheet_title, table). Raises RuntimeError with a readable message.
    """
    ext = os.path.splitext(src)[1].lower()
    if ext == ".xls":
        try:
            import xlrd
        except ImportError:
            raise RuntimeError(
                "读旧版 .xls 需要 xlrd：pip install xlrd　"
                "（或先用 Excel/WPS 打开后『另存为 .xlsx』再跑）"
            )
        try:
            book = xlrd.open_workbook(src, on_demand=True)
        except Exception as e:
            raise RuntimeError("这个 .xls 打不开（可能不是真正的 Excel 97-2003 文件）：%s" % e)
        try:
            sheets = book.sheets()
            if sheet_opt:
                names = [sh.name for sh in sheets]
                if sheet_opt not in names:
                    raise RuntimeError("没有名为 '%s' 的工作表（有：%s）"
                                       % (sheet_opt, ", ".join(names)))
                sh = next(s for s in sheets if s.name == sheet_opt)
            else:
                sh = sheets[0]
            table = [[_xls_cell_value(sh.cell(r, c), book.datemode) for c in range(sh.ncols)]
                     for r in range(sh.nrows)]
            return sh.name, table
        finally:
            try:
                book.release_resources()
            except Exception:
                pass

    wb_in = load_workbook(src, data_only=True)
    try:
        if sheet_opt:
            if sheet_opt not in wb_in.sheetnames:
                raise RuntimeError("No such sheet: %s (have: %s)"
                                   % (sheet_opt, ", ".join(wb_in.sheetnames)))
            ws_in = wb_in[sheet_opt]
        else:
            ws_in = wb_in.worksheets[0]
        return ws_in.title, [list(r) for r in ws_in.iter_rows(values_only=True)]
    finally:
        wb_in.close()

ILLEGAL = re.compile(r'[\\/:*?"<>|\r\n\t]')


def safe_name(s, fallback="blank"):
    """Turn a cell value into something Windows / macOS will both accept as a file name."""
    s = "" if s is None else str(s)
    s = s.strip().replace("\u00a0", " ")
    s = ILLEGAL.sub("_", s)
    s = re.sub(r"\s+", "_", s)
    s = s.strip("._")
    if not s:
        s = fallback
    return s[:60]


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
    """First row in the first `scan` rows that is at least half as wide as the widest one."""
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


def write_block(ws_target, headers, rows):
    ws_target.append(headers)
    for r in rows:
        ws_target.append(r)
    for ci in range(1, len(headers) + 1):
        ws_target.column_dimensions[get_column_letter(ci)].width = min(
            max(10, len(str(headers[ci - 1])) + 4), 34)
    for row in ws_target.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, (datetime, date)):
                c.number_format = "yyyy-mm-dd"
    ws_target.freeze_panes = "A2"
    if ws_target.max_row > 1:
        ws_target.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(headers)), ws_target.max_row)


def main():
    ap = argparse.ArgumentParser(description="Split a sheet into groups.")
    ap.add_argument("--src", required=True, help="input .xlsx")
    ap.add_argument("--sheet", default=None, help="sheet to read (default: first with data)")
    ap.add_argument("--header-row", type=int, default=None, help="1-based header row")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--by", default=None, help="column name or 1-based index to group by")
    g.add_argument("--chunk", type=int, default=None, help="split into blocks of N rows instead")
    ap.add_argument("--mode", choices=["files", "sheets"], default="files")
    ap.add_argument("--out", default=None,
                    help="files mode: output folder (default: <name>_split) | sheets mode: output .xlsx")
    ap.add_argument("--prefix", default=None, help="file name prefix (files mode)")
    ap.add_argument("--keep-blank", action="store_true", help="keep rows whose key is empty")
    args = ap.parse_args()

    src = os.path.abspath(args.src)
    if not os.path.exists(src):
        sys.exit("Not found: %s" % src)
    ext = os.path.splitext(src)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".xls"):
        sys.exit("Only .xlsx / .xlsm / .xls are supported.")

    try:
        _, table = load_table(src, args.sheet)
    except RuntimeError as e:
        sys.exit(str(e))

    if not table:
        sys.exit("Sheet is empty.")

    if args.header_row:
        hrow = args.header_row - 1
    else:
        hrow = auto_header_index(table)
    headers = [("" if h is None else str(h).strip()) or "Column_%d" % (i + 1)
               for i, h in enumerate(table[hrow])]
    body = [r + [None] * (len(headers) - len(r)) if len(r) < len(headers) else r[:len(headers)]
            for r in table[hrow + 1:]]

    # ---- decide groups ----
    groups = []           # list of (label, rows)
    if args.chunk:
        n = max(1, args.chunk)
        for i in range(0, len(body), n):
            groups.append(("part_%03d" % (i // n + 1), body[i:i + n]))
    else:
        key = args.by
        if re.fullmatch(r"\d+", str(key)):
            idx = int(key) - 1
            if idx < 0 or idx >= len(headers):
                sys.exit("--by %s: the sheet has only %d columns" % (key, len(headers)))
        else:
            if key not in headers:
                close = [h for h in headers if key.lower() in str(h).lower()]
                sys.exit("No column named '%s'. Columns are: %s%s"
                         % (key, ", ".join(headers),
                            ("  (did you mean: %s)" % ", ".join(close)) if close else ""))
            idx = headers.index(key)

        order, bucket = [], {}
        skipped = 0
        for r in body:
            v = r[idx]
            if v is None or (isinstance(v, str) and not v.strip()):
                if not args.keep_blank:
                    skipped += 1
                    continue
                label = "(blank)"
            else:
                label = str(v).strip()
            if label not in bucket:
                bucket[label] = []
                order.append(label)
            bucket[label].append(r)
        groups = [(label, bucket[label]) for label in order]
        if skipped:
            print("Skipped %d row(s) with an empty '%s'." % (skipped, headers[idx]))

    if not groups:
        sys.exit("Nothing to split: no data rows.")

    base = os.path.splitext(os.path.basename(src))[0]
    total = sum(len(r) for _, r in groups)

    # ---- write ----
    if args.mode == "sheets":
        out = os.path.abspath(args.out or ("%s_split.xlsx" % base))
        wb = Workbook()
        wb.remove(wb.active)
        used = set()
        for label, rows in groups:
            name = safe_name(label, "(blank)")
            while name.lower() in used:
                name += "_"
            used.add(name.lower())
            ws = wb.create_sheet(title=name[:31])
            write_block(ws, headers, rows)
        # index sheet
        ix = wb.create_sheet(title="_index", index=0)
        ix.append(["Group", "Rows"])
        for label, rows in groups:
            ix.append([label, len(rows)])
        ix.append(["TOTAL", total])
        ix.column_dimensions["A"].width = 32
        ix.column_dimensions["B"].width = 12
        wb.save(out)
        print("\n%-30s %8s" % ("sheet", "rows"))
        print("-" * 40)
        for label, rows in groups:
            print("%-30s %8d" % (label[:30], len(rows)))
        print("-" * 40)
        print("%-30s %8d" % ("TOTAL", total))
        print("\nSaved -> %s  (%d sheets + _index)" % (out, len(groups)))
    else:
        outdir = os.path.abspath(args.out or ("%s_split" % base))
        os.makedirs(outdir, exist_ok=True)
        prefix = args.prefix if args.prefix is not None else ""
        written = []
        for i, (label, rows) in enumerate(groups, start=1):
            fname = "%s%02d_%s.xlsx" % (prefix, i, safe_name(label))
            path = os.path.join(outdir, fname)
            wb = Workbook()
            ws = wb.active
            ws.title = safe_name(label, "(blank)")[:31]
            write_block(ws, headers, rows)
            wb.save(path)
            written.append((fname, len(rows)))

        print("\n%-46s %8s" % ("file", "rows"))
        print("-" * 56)
        for f, n in written:
            print("%-46s %8d" % (f[:46], n))
        print("-" * 56)
        print("%-46s %8d" % ("TOTAL", total))
        print("\nSaved %d files -> %s" % (len(written), outdir))

    # ---- reconciliation ----
    if total == len(body):
        print("Reconcile       : OK  (%d rows in, %d rows out)" % (len(body), total))
    else:
        print("Reconcile       : %d rows in, %d out — blank-key rows were dropped"
              % (len(body), total))


if __name__ == "__main__":
    main()
