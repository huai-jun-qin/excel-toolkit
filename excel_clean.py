#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_clean.py - Clean one workbook and report every single change it made.

Nothing is changed silently. Every coercion lands in the "_log" sheet:

    Excel row | Column | Before | After | Action

and anything the tool cannot decide (e.g. a date that could honestly be read
two ways) is left as-is and flagged instead of being guessed at.

Requires: openpyxl  (pip install openpyxl)

Usage:
    python3 excel_clean.py --src dirty.xlsx
    python3 excel_clean.py --src dirty.xlsx --out clean.xlsx --dedupe
    python3 excel_clean.py --src dirty.xlsx --date-order dmy
"""

import argparse
import os
import re
import sys
import unicodedata
from collections import Counter
from datetime import date, datetime, timedelta

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Missing dependency. Run:  pip install openpyxl")

MAX_LOG_ROWS = 20000

ZERO_WIDTH = "\u200b\u200c\u200d\ufeff"
NBSP = "\u00a0"


# ---------------------------------------------------------------------------
# text
# ---------------------------------------------------------------------------
def tidy_text(s):
    """Trim, collapse inner whitespace, drop nbsp / zero-width / ideographic space."""
    if not isinstance(s, str):
        return s
    s = s.replace(NBSP, " ").replace("\u3000", " ")
    for ch in ZERO_WIDTH:
        s = s.replace(ch, "")
    return re.sub(r"\s+", " ", s).strip()


def halfwidth(s):
    """Full-width digits / letters / punctuation -> half-width."""
    return unicodedata.normalize("NFKC", s)


# ---------------------------------------------------------------------------
# numbers
# ---------------------------------------------------------------------------
INT_RE = re.compile(r"^[+-]?\d+$")
NUM_RE = re.compile(r"^[+-]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$")
MONEY_RE = re.compile(r"^[¥$€£￥]?\s*[+-]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?\s*(元|块|rmb|cny|usd)?$", re.I)


def to_number(s):
    """'1,234.50' / '¥ 88 元' -> 1234.5 / 88 ; None if not numeric."""
    if isinstance(s, bool):
        return None
    if isinstance(s, (int, float)):
        return s
    if not isinstance(s, str):
        return None
    t = tidy_text(halfwidth(s))
    if not t:
        return None
    if NUM_RE.match(t):
        bare = t.replace(",", "")
        return int(bare) if INT_RE.match(bare) else float(bare)
    if MONEY_RE.match(t):
        bare = re.sub(r"[¥$€£￥\s元块]", "", t)
        bare = re.sub(r"(rmb|cny|usd)$", "", bare, flags=re.I).replace(",", "")
        if not bare:
            return None
        try:
            return float(bare) if "." in bare else int(bare)
        except ValueError:
            return None
    return None


def looks_like_id(s):
    """Must stay text: leading zeros, e.g. 00123."""
    return isinstance(s, str) and re.match(r"^0\d+$", tidy_text(halfwidth(s))) is not None


# ---------------------------------------------------------------------------
# dates
# ---------------------------------------------------------------------------
YMD_RE = re.compile(r"^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?$")
YMD_DT_RE = re.compile(r"^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?[ T](\d{1,2}):(\d{2})(?::(\d{2}))?$")
YMD8_RE = re.compile(r"^(\d{4})(\d{2})(\d{2})$")
AMB_RE = re.compile(r"^(\d{1,2})([-/.])(\d{1,2})\2(\d{4})$")     # 05.03.2026 / 5-3-2026
AMB2_RE = re.compile(r"^(\d{1,2})([-/.])(\d{1,2})\2(\d{2})$")    # 05.03.26


def parse_date(v, order="auto"):
    """Return (datetime|None, status, sep_hint).

    status: 'ok'        parsed
            'ambiguous' two readings are both legal, caller must decide
            'no'        not a date
    sep_hint: '.', '/', '-' for ambiguous values, else None
    """
    if isinstance(v, datetime):
        return v, "ok", None
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day), "ok", None

    if isinstance(v, (int, float)) and not isinstance(v, bool):
        # Excel serial date (day 1 = 1900-01-01, offset 1899-12-30)
        if 20000 <= float(v) <= 60000 and float(v) == int(float(v)):
            return datetime(1899, 12, 30) + timedelta(days=int(v)), "ok", None
        return None, "no", None

    if not isinstance(v, str):
        return None, "no", None
    s = tidy_text(halfwidth(v))
    if not s:
        return None, "no", None

    m = YMD_DT_RE.match(s)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4)), int(m.group(5)),
                            int(m.group(6)) if m.group(6) else 0), "ok", None
        except ValueError:
            return None, "no", None

    m = YMD_RE.match(s) or YMD8_RE.match(s)
    if m:
        try:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            return datetime(y, mo, d), "ok", None
        except ValueError:
            return None, "no", None

    m = AMB_RE.match(s)
    if m:
        a, sep, b, y = int(m.group(1)), m.group(2), int(m.group(3)), int(m.group(4))
    else:
        m = AMB2_RE.match(s)
        if not m:
            return None, "no", None
        a, sep, b = int(m.group(1)), m.group(2), int(m.group(3))
        y = 2000 + int(m.group(4))

    if a > 12 and b > 12:
        return None, "no", None
    if a > 12:
        d, mo = a, b
    elif b > 12:
        d, mo = b, a
    else:
        if order == "dmy":
            d, mo = a, b
        elif order == "mdy":
            d, mo = b, a
        else:
            return None, "ambiguous", sep
    try:
        return datetime(y, mo, d), "ok", None
    except ValueError:
        return None, "no", None


def date_evidence(v):
    """For an unambiguous 05.13.2026-style value, say which reading it proved."""
    if isinstance(v, str):
        s = tidy_text(halfwidth(v))
        m = AMB_RE.match(s) or AMB2_RE.match(s)
        if m:
            a, b = int(m.group(1)), int(m.group(3))
            if a > 12 and b <= 12:
                return "dmy"
            if b > 12 and a <= 12:
                return "mdy"
    return None


# ---------------------------------------------------------------------------
# sheet geometry
# ---------------------------------------------------------------------------
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
    """First row in the first `scan` rows that is at least half as wide as the widest one.

    Skips the 'exported from ERP' title line that sits above real headers.
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


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Clean one workbook and log every change.")
    ap.add_argument("--src", required=True, help="input .xlsx")
    ap.add_argument("--out", default=None, help="output .xlsx (default: <name>_clean.xlsx)")
    ap.add_argument("--sheet", default=None, help="sheet to clean (default: first)")
    ap.add_argument("--header-row", type=int, default=None, help="1-based header row (default: auto)")
    ap.add_argument("--dates", choices=["iso", "keep"], default="iso",
                    help="iso = unify every date column to yyyy-mm-dd (default)")
    ap.add_argument("--date-order", choices=["auto", "dmy", "mdy"], default="auto",
                    help="how to read 05.03.2026. auto = decide per column from the column itself")
    ap.add_argument("--numbers", choices=["coerce", "keep"], default="coerce",
                    help="coerce text-numbers into real numbers (default)")
    ap.add_argument("--dedupe", action="store_true", help="drop exact duplicate rows")
    ap.add_argument("--dedupe-on", nargs="*", default=None, help="compare only these columns")
    ap.add_argument("--keep-empty-rows", action="store_true", help="keep fully empty rows")
    ap.add_argument("--max-log", type=int, default=MAX_LOG_ROWS,
                    help="cap the _log sheet at N rows (default %d). Truncation is reported, "
                         "never silent." % MAX_LOG_ROWS)
    args = ap.parse_args()

    src = os.path.abspath(args.src)
    if not os.path.exists(src):
        sys.exit("Not found: %s" % src)
    if os.path.splitext(src)[1].lower() == ".xls":
        # legacy .xls: reuse the reader that excel_merge.py already ships (needs xlrd)
        try:
            from excel_merge import read_xls_sheets
        except ImportError:
            sys.exit("excel_merge.py must sit next to excel_clean.py to read .xls "
                     "(or run excel_merge.py first — it writes a .xlsx)")
        sheets = list(read_xls_sheets(src))
        if not sheets:
            sys.exit("This .xls has no sheets: %s" % src)
        names = [n for n, _ in sheets]
        if args.sheet and args.sheet not in names:
            sys.exit("No such sheet: %s (have: %s)" % (args.sheet, ", ".join(names)))
        sheet_title, table = next((n, t) for n, t in sheets if not args.sheet or n == args.sheet)
        if not table:
            sys.exit("Sheet is empty: nothing to clean.")
    else:
        wb_in = load_workbook(src, data_only=True)
        if args.sheet:
            if args.sheet not in wb_in.sheetnames:
                sys.exit("No such sheet: %s (have: %s)" % (args.sheet, ", ".join(wb_in.sheetnames)))
            ws_in = wb_in[args.sheet]
        else:
            ws_in = wb_in.worksheets[0]
        sheet_title = ws_in.title
        table = [list(r) for r in ws_in.iter_rows(values_only=True)]

    if not table:
        sys.exit("Sheet is empty: nothing to clean.")

    if args.header_row:
        hrow = args.header_row - 1
        if hrow >= len(table):
            sys.exit("--header-row %d is past the end of the sheet" % args.header_row)
    else:
        hrow = auto_header_index(table)

    raw_headers = table[hrow]
    body = table[hrow + 1:]

    headers, seen = [], {}
    for i, h in enumerate(raw_headers):
        name = tidy_text(h) if h is not None else ""
        if name == "":
            name = "Column_%d" % (i + 1)
        if name in seen:
            seen[name] += 1
            name = "%s_%d" % (name, seen[name])
        else:
            seen[name] = 1
        headers.append(name)

    ncol = len(headers)
    body = [r + [None] * (ncol - len(r)) if len(r) < ncol else r[:ncol] for r in body]

    log = []
    log_dropped = 0
    cap = args.max_log

    def note(row_no, col, before, after, action):
        nonlocal log_dropped
        if len(log) < cap:
            log.append((row_no, col, before, after, action))
        else:
            log_dropped += 1

    for i, (raw, new) in enumerate(zip(raw_headers, headers)):
        if raw is not None and str(raw) != new:
            note(hrow + 1, get_column_letter(i + 1), raw, new, "header tidied")

    # ---- decide which columns are dates / numbers, and how to read them ----
    sample = [r for r in body if any(c not in (None, "") for c in r)][:500]
    date_cols, num_cols, col_order = set(), set(), {}

    for ci in range(ncol):
        vals = [r[ci] for r in sample if r[ci] not in (None, "")]
        if not vals:
            continue

        if args.dates == "iso" and not all(
                isinstance(v, str) and looks_like_id(v) for v in vals):
            hits = 0
            ev = Counter()
            seps = Counter()
            for v in vals:
                dt, st, sep = parse_date(v, "auto")
                if dt is not None:
                    hits += 1
                    e = date_evidence(v)
                    if e:
                        ev[e] += 1
                elif st == "ambiguous" and sep:
                    seps[sep] += 1
            if hits + sum(seps.values()) > 0 and \
                    (hits + sum(seps.values())) / len(vals) >= 0.6:
                date_cols.add(ci)
                if args.date_order != "auto":
                    col_order[ci] = args.date_order
                elif ev["dmy"] > ev["mdy"]:
                    col_order[ci] = "dmy"
                elif ev["mdy"] > ev["dmy"]:
                    col_order[ci] = "mdy"
                elif seps.most_common(1):
                    # no proof in the column -> the separator is the only hint
                    col_order[ci] = "dmy" if seps.most_common(1)[0][0] == "." else "mdy"
                else:
                    col_order[ci] = "auto"

        if args.numbers == "coerce":
            textual = [v for v in vals if isinstance(v, str) and not looks_like_id(v)]
            if textual:
                nhit = sum(1 for v in textual if to_number(v) is not None)
                if nhit / len(vals) >= 0.5 and nhit / len(textual) >= 0.8:
                    num_cols.add(ci)

    # ---- cell level cleaning ----
    out_rows = []
    dropped_empty = 0
    ambiguous = 0
    order_used = {}

    for ri, r in enumerate(body, start=hrow + 2):
        if not args.keep_empty_rows and all(
                c is None or (isinstance(c, str) and not c.strip()) for c in r):
            dropped_empty += 1
            continue

        new_r = []
        for ci, v in enumerate(r):
            nv = v
            if isinstance(v, str):
                t = tidy_text(halfwidth(v))
                if t != v:
                    note(ri, headers[ci], v, t, "whitespace / full-width normalised")
                    nv = t

            if ci in date_cols:
                order = col_order.get(ci, "auto")
                # parse strictly first: only a genuinely ambiguous value gets the assumption
                dt, st, sep = parse_date(nv, "auto")
                assumed = False
                if dt is None and st == "ambiguous" and order != "auto":
                    dt, _, _ = parse_date(nv, order)
                    assumed = dt is not None
                if dt is not None:
                    was = nv.strftime("%Y-%m-%d") if isinstance(nv, (datetime, date)) else nv
                    already = (isinstance(nv, (datetime, date)) and
                               (nv.hour, nv.minute) == (0, 0) and
                               nv.strftime("%Y-%m-%d") == dt.strftime("%Y-%m-%d"))
                    if not already:
                        act = "date unified"
                        if assumed:
                            src_txt = "DD.MM.YYYY" if order == "dmy" else "MM/DD/YYYY"
                            act = "date unified (read as %s)" % src_txt
                            order_used[headers[ci]] = "assumed %s" % src_txt
                        note(ri, headers[ci], was, dt.strftime("%Y-%m-%d"), act)
                    nv = dt
                elif st == "ambiguous":
                    ambiguous += 1
                    note(ri, headers[ci], nv, nv,
                         "AMBIGUOUS, left as-is (run with --date-order dmy|mdy)")

            if ci in num_cols and isinstance(nv, str) and not looks_like_id(nv):
                num = to_number(nv)
                if num is not None:
                    note(ri, headers[ci], nv, num, "text -> number")
                    nv = num

            if nv is None and v not in (None, ""):
                note(ri, headers[ci], v, "(empty)", "value could not be parsed")
            new_r.append(nv)
        out_rows.append((ri, new_r))

    # ---- dedupe ----
    dup_rows = []
    if args.dedupe:
        keys = headers if not args.dedupe_on else args.dedupe_on
        idxs = []
        for k in keys:
            if k not in headers:
                sys.exit("--dedupe-on: no such column: %s" % k)
            idxs.append(headers.index(k))
        seen_keys, keep = set(), []
        for ri, r in out_rows:
            k = tuple(str(r[i]) for i in idxs)
            if k in seen_keys:
                dup_rows.append(ri)
                note(ri, "(whole row)", " | ".join(str(r[i]) for i in idxs)[:80], "removed",
                     "duplicate dropped")
                continue
            seen_keys.add(k)
            keep.append((ri, r))
        out_rows = keep

    # ---- write ----
    out = args.out or re.sub(r"\.xlsx$", "", src) + "_clean.xlsx"
    out = os.path.abspath(out)
    wb = Workbook()
    ws = wb.active
    ws.title = "Cleaned"
    ws.append(headers)
    for _, r in out_rows:
        ws.append(r)
    for ci in range(1, ncol + 1):
        ws.column_dimensions[get_column_letter(ci)].width = min(
            max(10, len(str(headers[ci - 1])) + 4), 34)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, datetime):
                c.number_format = "yyyy-mm-dd" if (c.value.hour, c.value.minute) == (0, 0) \
                    else "yyyy-mm-dd hh:mm"
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(ncol), ws.max_row)

    ls = wb.create_sheet("_log")
    ls.append(["Excel row", "Column", "Before", "After", "Action"])
    for entry in log:
        ls.append(list(entry))
    if log_dropped:
        ls.append(["", "TRUNCATED", log_dropped, "",
                   "log capped at %d rows - %d further changes were made but are NOT listed here"
                   % (cap, log_dropped)])
    for letter, w in (("A", 11), ("B", 22), ("C", 30), ("D", 30), ("E", 46)):
        ls.column_dimensions[letter].width = w

    ss = wb.create_sheet("_summary")
    missing = {}
    for ci, h in enumerate(headers):
        n = sum(1 for _, r in out_rows if r[ci] in (None, ""))
        if n:
            missing[h] = n
    rows_in = len(body)
    ss.append(["Metric", "Value"])
    ss.append(["Source file", os.path.basename(src)])
    ss.append(["Sheet", sheet_title])
    ss.append(["Header row", hrow + 1])
    ss.append(["Rows read", rows_in])
    ss.append(["Empty rows dropped", dropped_empty])
    ss.append(["Duplicate rows dropped", len(dup_rows)])
    ss.append(["Rows written", len(out_rows)])
    ss.append(["Cells changed", len(log) + log_dropped])
    if log_dropped:
        ss.append(["Change log truncated", "%d changes not listed (cap %d)" % (log_dropped, cap)])
    ss.append(["Ambiguous dates left alone", ambiguous])
    ss.append(["Reconciliation",
               "OK" if rows_in - dropped_empty - len(dup_rows) == len(out_rows) else "MISMATCH"])
    ss.append([])
    if order_used:
        ss.append(["Date columns read as", ""])
        for h, o in order_used.items():
            ss.append([h, o])
        ss.append([])
    ss.append(["Empty cells per column", ""])
    for h, n in sorted(missing.items(), key=lambda x: -x[1]):
        ss.append([h, n])
    ss.column_dimensions["A"].width = 34
    ss.column_dimensions["B"].width = 26

    wb.save(out)

    # ---- console ----
    print("=" * 68)
    print("Source          : %s  [%s]" % (os.path.basename(src), sheet_title))
    print("Header row      : %d" % (hrow + 1))
    print("Rows read       : %d" % rows_in)
    print("Empty dropped   : %d" % dropped_empty)
    print("Duplicate drop  : %d" % len(dup_rows))
    print("Rows written    : %d" % len(out_rows))
    print("Cells changed   : %d" % (len(log) + log_dropped))
    print("Reconcile       : %s" % ("OK" if rows_in - dropped_empty - len(dup_rows) == len(out_rows)
                                   else "MISMATCH"))
    if log_dropped:
        print("!! LOG TRUNCATED: %d changes not listed (cap %d). "
              "Re-run with --max-log %d to list them all."
              % (log_dropped, cap, len(log) + log_dropped))
    if ambiguous:
        print("Ambiguous dates : %d left untouched (see _log)" % ambiguous)
    print("-" * 68)
    print("Columns         : %s" % ", ".join(headers))
    acts = Counter(e[4] for e in log)
    if acts:
        print("\nWhat was changed:")
        for a, n in acts.most_common():
            print("  %-56s %d" % (a, n))
    if missing:
        print("\nEmpty cells per column:")
        for h, n in sorted(missing.items(), key=lambda x: -x[1])[:10]:
            print("  %-30s %d" % (h, n))
    print("\nSaved -> %s" % out)
    print("Open the '_log' sheet: every cell that changed is listed with before/after.")


if __name__ == "__main__":
    main()
