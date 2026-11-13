#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_profile.py - Read the data before you clean it. Writes a report, never edits the source.

The other three tools change files. This one only looks. It answers the question
you should ask before running any cleanup:

    What is actually wrong with this workbook?

For every sheet it finds the header row, then reports per column:

    how full it is, what types live in it, how many cells are numbers stored as
    text, how many carry stray whitespace / full-width characters, how many date
    spellings appear in one column, whether the column is an ID that must stay
    text, and how many rows are exact duplicates.

and per sheet:

    merged blocks, blank rows, a header printed again inside the data, rows wider
    than the header, duplicate column names, other sheets that look like covers.

The source workbook is opened read-only and never written. Output is a new
workbook with `_summary` / `_columns` / `_issues` sheets, plus optional
Markdown, HTML and JSON copies of the same findings.

Requires: openpyxl  (pip install openpyxl)
Optional: xlrd      (pip install xlrd)   -- only to read legacy .xls

Usage:
    python3 excel_profile.py --src dirty.xlsx
    python3 excel_profile.py --src dirty.xlsx --md report.md --html report.html
    python3 excel_profile.py --src dirty.xlsx --strict      # exit 2 if a high issue exists
    python3 excel_profile.py --src dirty.xlsx --sheet 明细 --header-row 3
"""

import argparse
import csv
import io
import json
import os
import re
import sys
import unicodedata
from collections import Counter
from datetime import date, datetime, time

try:
    from openpyxl import Workbook, load_workbook
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

TOOL_VERSION = "1.0"
MAX_SAMPLES = 4
DEFAULT_MAX_ROWS = 200000

ZERO_WIDTH = "\u200b\u200c\u200d\ufeff"
ODD_SPACE = "\u00a0\u2009\u3000"

# header words that make a bare integer mean "Excel serial date"
DATE_HEADER_RE = re.compile(r"日期|时间|date|time|day|月份|period", re.I)
ID_HEADER_RE = re.compile(
    r"(^|[^a-z])(id|no|num|code|sn|sku)([^a-z]|$)|编号|单号|工号|编码|订单号|手机|电话|电话号|身份证|邮编|税号|账号",
    re.I,
)
MONEY_HEADER_RE = re.compile(r"金额|价|费|额|amount|price|total|cost|revenue|sum|salary|工资", re.I)


# ---------------------------------------------------------------------------
# small value tests
# ---------------------------------------------------------------------------
def is_blank(v):
    if v is None:
        return True
    if isinstance(v, str) and not v.strip():
        return True
    return False


def tidy(s):
    s = s.replace("\u00a0", " ").replace("\u3000", " ")
    for ch in ZERO_WIDTH:
        s = s.replace(ch, "")
    return re.sub(r"\s+", " ", s).strip()


INT_RE = re.compile(r"^[+-]?\d+$")
NUM_RE = re.compile(r"^[+-]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$")
MONEY_RE = re.compile(
    r"^[¥$€£￥]?\s*[+-]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?\s*(元|块|rmb|cny|usd)?$", re.I
)


def to_number(s):
    """'1,234.50' / '¥ 88 元' -> float ; None if not numeric text."""
    if isinstance(s, bool) or not isinstance(s, str):
        return None
    t = tidy(unicodedata.normalize("NFKC", s))
    if not t:
        return None
    compact = re.sub(r"(?<=\d)[\s\u00a0\u2009\u3000]+(?=\d)", "", t)
    if NUM_RE.match(compact):
        bare = compact.replace(",", "")
        return float(bare) if "." in bare else int(bare)
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


def is_leading_zero_id(s):
    """'00123' - an identifier, not a number that lost its zeros."""
    return isinstance(s, str) and re.fullmatch(r"0\d+", s.strip()) is not None


def date_format_of(s):
    """Return a human label for the spelling, or None. Read-only: no parsing to a real date."""
    if not isinstance(s, str):
        return None
    t = s.strip()
    if not t:
        return None
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", t):
        return "2026-03-05"
    if re.fullmatch(r"\d{4}/\d{1,2}/\d{1,2}", t):
        return "2026/3/5"
    if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", t):
        return "3/5/2026 (order unknown)"
    if re.fullmatch(r"\d{1,2}\.\d{1,2}\.\d{4}", t):
        return "05.03.2026 (order unknown)"
    if re.fullmatch(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日?", t):
        return "2026年3月5日"
    if re.fullmatch(r"\d{8}", t):
        return "20260305"
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}[ T]\d{1,2}:\d{2}(:\d{2})?", t):
        return "2026-03-05 14:30"
    if re.fullmatch(r"\d{1,2}\.\d{1,2}\.\d{4}\s+\d{1,2}:\d{2}", t):
        return "05.03.2026 14:30"
    return None


def is_ambiguous_date(s):
    """True when both day-first and month-first readings are legal."""
    if not isinstance(s, str):
        return False
    t = s.strip()
    for pat in (r"(\d{1,2})[./-](\d{1,2})[./-](\d{4})",):
        m = re.fullmatch(pat, t)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a <= 12 and b <= 12:
                return True
            return False
    return False


def is_dirty_text(s):
    """Stray whitespace, full-width forms, zero-width or non-breaking characters."""
    if not isinstance(s, str) or not s:
        return False
    if s != s.strip():
        return True
    if any(ch in s for ch in ZERO_WIDTH) or any(ch in s for ch in ODD_SPACE):
        return True
    if re.search(r"[\u3000\u200b]", s):
        return True
    # full-width digits / latin / punctuation -> normalises to something else
    if unicodedata.normalize("NFKC", s) != s:
        return True
    if re.search(r"[ \t]{2,}", s):
        return True
    return False


# ---------------------------------------------------------------------------
# loading (never writes)
# ---------------------------------------------------------------------------
class SheetData(object):
    def __init__(self, name, rows, merged):
        self.name = name
        self.rows = rows
        self.merged = merged


def decode_bytes(raw):
    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk", "latin-1"):
        try:
            return raw.decode(enc), enc
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", "replace"), "utf-8 (replaced)"


def read_csv_sheet(path, max_rows):
    raw = open(path, "rb").read()
    text, enc = decode_bytes(raw)
    head = text[:8192]
    delim = ","
    for cand in ("\t", ";", ","):
        if head.count(cand) > head.count(delim):
            delim = cand
    rows = []
    for i, row in enumerate(csv.reader(io.StringIO(text), delimiter=delim)):
        if i >= max_rows:
            break
        rows.append([c for c in row])
    return [SheetData(os.path.basename(path), rows, [])], {"encoding": enc, "delimiter": delim}


def read_xls_sheet(path, max_rows):
    try:
        import xlrd
    except ImportError:
        sys.exit(
            "This is a legacy .xls file and xlrd is not installed.\n"
            "    %s -m pip install --user xlrd\n"
            "Or save it as .xlsx and run again." % sys.executable
        )
    book = xlrd.open_workbook(path)
    out = []
    for sh in book.sheets():
        rows = []
        for r in range(min(sh.nrows, max_rows)):
            rows.append([sh.cell_value(r, c) for c in range(sh.ncols)])
        out.append(SheetData(sh.name, rows, []))
    return out, {"encoding": "xls"}


def load_source(path, wanted_sheet=None, max_rows=DEFAULT_MAX_ROWS):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        sheets, meta = read_csv_sheet(path, max_rows)
    elif ext == ".xls":
        sheets, meta = read_xls_sheet(path, max_rows)
    else:
        wb = load_workbook(path, data_only=True, read_only=False)
        sheets = []
        for ws in wb.worksheets:
            if wanted_sheet and ws.title != wanted_sheet:
                continue
            rows = []
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i >= max_rows:
                    break
                rows.append(list(row))
            merged = [str(r) for r in ws.merged_cells.ranges]
            sheets.append(SheetData(ws.title, rows, merged))
        meta = {"encoding": "xlsx", "sheets_total": len(wb.sheetnames),
                "all_sheet_names": list(wb.sheetnames)}
        wb.close()
    if not sheets:
        sys.exit("No sheet matched --sheet %r in %s" % (wanted_sheet, path))
    return sheets, meta


# ---------------------------------------------------------------------------
# header detection (shared idea with excel_merge, kept small on purpose)
# ---------------------------------------------------------------------------
def norm_header(v):
    if v is None:
        return ""
    return re.sub(r"\s+", "", str(v)).strip().lower()


def header_score(row):
    vals = [v for v in row if not is_blank(v)]
    if len(vals) < 2:
        return -1e9
    score = 0.0
    for v in vals:
        if isinstance(v, str):
            t = v.strip()
            if len(t) <= 40:
                score += 2
            else:
                score -= 1
            if to_number(t) is not None:
                score -= 1.5
            if date_format_of(t):
                score -= 1.5
            if re.match(r"^[^:：]{1,12}[:：]", t):
                score -= 1.0
        else:
            score -= 2
    score += min(len(vals), 12) * 0.2
    return score


def detect_header(rows, explicit=None, scan=10):
    if explicit:
        return max(0, explicit - 1)
    best, best_score = 0, -1e9
    for i, row in enumerate(rows[:scan]):
        sc = header_score(row)
        if sc > best_score:
            best, best_score = i, sc
    return best


def looks_like_cover(rows):
    filled = [r for r in rows if any(not is_blank(v) for v in r)]
    if len(filled) <= 3:
        return True
    return False


# ---------------------------------------------------------------------------
# profiling
# ---------------------------------------------------------------------------
def type_of(v):
    if is_blank(v):
        return "empty"
    if isinstance(v, bool):
        return "bool"
    if isinstance(v, (datetime, date, time)):
        return "date"
    if isinstance(v, (int, float)):
        return "number"
    if isinstance(v, str):
        return "text"
    return "other"


def profile_sheet(sheet, header_idx, max_rows):
    rows = sheet.rows
    header_row = rows[header_idx] if header_idx < len(rows) else []
    width = max([len(r) for r in rows] or [0])
    headers = []
    for c in range(width):
        v = header_row[c] if c < len(header_row) else None
        headers.append("" if is_blank(v) else str(v).strip())

    data = rows[header_idx + 1:]
    total_rows = len(data)
    issues = []
    columns = []

    def add(severity, col, where, issue, detail, suggestion):
        issues.append({
            "severity": severity,
            "sheet": sheet.name,
            "column": col,
            "where": where,
            "issue": issue,
            "detail": detail,
            "suggestion": suggestion,
        })

    # ---- sheet level -------------------------------------------------------
    # only merges that sit *inside* the data area are a problem worth reporting.
    # A merged title band above the header is normal in exported reports and
    # flagging it would bury the findings that matter.
    merges_in_data = []
    for rng in sheet.merged:
        m = re.match(r"^[A-Za-z]+(\d+)", str(rng))
        if m and int(m.group(1)) > header_idx + 1:
            merges_in_data.append(rng)
    if merges_in_data:
        add("medium", "", "", "合并单元格 %d 处" % len(merges_in_data),
            "值只存在左上角格子里，按行读取时下面几行看起来是空的："
            + ", ".join(merges_in_data[:MAX_SAMPLES]),
            "合并前先拆开并向下填充，或确认工具会填充合并块")

    blank_rows = 0
    over_width = 0
    repeated_header = 0
    seen_rows = Counter()
    dup_rows = 0
    dup_first = ""
    for offset, row in enumerate(data):
        excel_row = header_idx + 2 + offset
        filled = [v for v in row if not is_blank(v)]
        if not filled:
            blank_rows += 1
            continue
        if len(filled) > len([h for h in headers if h]):
            over_width += 1
        if headers and tuple(norm_header(v) for v in row[:len(headers)]) == tuple(norm_header(h) for h in headers):
            repeated_header += 1
        key = tuple("" if is_blank(v) else str(v) for v in row)
        seen_rows[key] += 1
        if seen_rows[key] == 2:
            dup_rows += 1
            if not dup_first:
                dup_first = "row %d" % excel_row
        elif seen_rows[key] > 2:
            dup_rows += 1

    if blank_rows:
        add("low", "", "", "数据区里有 %d 行是空行" % blank_rows,
            "夹在数据中间的空行会把表格切断，排序、透视、筛选都会出错",
            "清洗时去掉空行（保留其位置记录）")
    if repeated_header:
        add("medium", "", "", "表头在数据里又出现 %d 次" % repeated_header,
            "每页重复打印表头的导出文件很常见；当成数据读会多出一行「姓名|邮箱」",
            "合并/清洗时识别并丢弃这些行")
    if over_width:
        add("low", "", "", "%d 行的列数比表头多" % over_width,
            "多出来的值没有列名，可能被直接丢掉",
            "检查是否漏了列名，或让工具写成 未命名列_n")
    if dup_rows:
        add("medium", "", "", "完全重复的行 %d 行" % dup_rows,
            "第一次出现的位置：%s；注意只有每一列都相同才算重复" % (dup_first or "-"),
            "确认业务上确实该去重后，用 excel_clean.py --dedupe")

    # duplicate column names
    name_count = Counter(norm_header(h) for h in headers if h)
    dup_names = [h for h, n in name_count.items() if n > 1]
    if dup_names:
        add("high", "", "", "表头有重名列 %d 组" % len(dup_names),
            "重名：" + ", ".join(dup_names[:MAX_SAMPLES]) + "；读取时后一列会覆盖前一列，行数还对得上，所以最难发现",
            "改成唯一列名后再合并或清洗")

    empty_headers = [i + 1 for i, h in enumerate(headers) if not h]
    if empty_headers:
        add("low", "", "", "%d 列没有表头" % len(empty_headers),
            "列号：" + ", ".join(str(i) for i in empty_headers[:MAX_SAMPLES]),
            "补列名，否则合并时这些列只能靠位置对齐")

    if looks_like_cover(rows):
        add("info", "", "", "这张表看起来像封面/说明页",
            "有效行数 ≤ 3，通常不是数据表",
            "指定真正的工作表，如 --sheet 明细")

    # ---- column level ------------------------------------------------------
    data_count = 0
    for row in data:
        if any(not is_blank(v) for v in row):
            data_count += 1
    # empty share is measured against every row in the data region, not just the
    # filled ones, otherwise a column of blanks can report more than 100%.
    denom = max(1, total_rows)

    for c in range(len(headers)):
        header = headers[c]
        col_letter = column_letter(c)
        types = Counter()
        empties = 0
        dirty = 0
        dirty_cells = []
        numtext = 0
        numtext_cells = []
        date_labels = Counter()
        date_cells = []
        ambiguous = 0
        leading_zero = 0
        uniq = set()
        maxlen = 0
        numbers = 0
        texts = 0

        for offset, row in enumerate(data):
            v = row[c] if c < len(row) else None
            excel_row = header_idx + 2 + offset
            t = type_of(v)
            types[t] += 1
            if t == "empty":
                empties += 1
                continue
            cell = "%s%d" % (col_letter, excel_row)
            uniq.add("" if is_blank(v) else str(v))
            if isinstance(v, str):
                maxlen = max(maxlen, len(v))
                texts += 1
                if is_dirty_text(v):
                    dirty += 1
                    if len(dirty_cells) < MAX_SAMPLES:
                        dirty_cells.append(cell)
                label = date_format_of(v)
                n = to_number(v)
                # a date spelling such as 20260305 is numeric text but must NOT be
                # "fixed" into a number - it would lose its meaning entirely.
                if n is not None and not is_leading_zero_id(v) and not label:
                    numtext += 1
                    if len(numtext_cells) < MAX_SAMPLES:
                        numtext_cells.append(cell)
                if is_leading_zero_id(v):
                    leading_zero += 1
                if label:
                    date_labels[label] += 1
                    if len(date_cells) < MAX_SAMPLES:
                        date_cells.append(cell)
                    if is_ambiguous_date(v):
                        ambiguous += 1
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                numbers += 1
                if DATE_HEADER_RE.search(header) and float(v).is_integer() and 1 <= v <= 60000:
                    date_labels["Excel 序列号 (number)"] += 1
                    if len(date_cells) < MAX_SAMPLES:
                        date_cells.append(cell)
            elif isinstance(v, (datetime, date)):
                date_labels["真实日期 (datetime)"] += 1

        filled = max(0, total_rows - empties)
        columns.append({
            "column": col_letter,
            "header": header,
            "filled": filled,
            "empty": empties,
            "empty_pct": round(min(100.0, 100.0 * empties / denom), 1),
            "unique": len(uniq),
            "types": ", ".join("%s %d" % (k, n) for k, n in types.most_common() if k != "empty") or "-",
            "numbers_as_text": numtext,
            "dirty_text": dirty,
            "date_formats": len(date_labels),
            "date_labels": " | ".join(date_labels.keys()),
            "max_len": maxlen,
        })

        if numtext:
            add("medium", header, ", ".join(numtext_cells), "%d 个数字被存成了文本" % numtext,
                "例如 " + ", ".join(numtext_cells) + "；SUM/透视表会静默跳过它们",
                "清洗时转成真数字；编号列除外")
        if dirty:
            add("medium", header, ", ".join(dirty_cells), "%d 个格子有多余空格/全角字符/不可见字符" % dirty,
                "例如 " + ", ".join(dirty_cells) + "；VLOOKUP 对不上多半是这里",
                "去空格、半角化、去零宽字符")
        if len(date_labels) > 1:
            add("high", header, ", ".join(date_cells), "同一列有 %d 种日期写法" % len(date_labels),
                "写法：" + " / ".join(date_labels.keys()),
                "统一成 yyyy-mm-dd；先确认日/月顺序")
        if ambiguous:
            add("high", header, ", ".join(date_cells), "%d 个日期日/月无法判定" % ambiguous,
                "如 05.03.2026，两种读法都合法；猜错会整整错两个月",
                "用 --date-order dmy|mdy 明确指定，或让工具标出来问客户")
        if numbers >= 2 and texts >= 2 and not numtext:
            add("medium", header, "", "同一列混了数字和文本",
                "数字 %d 个、文本 %d 个" % (numbers, texts),
                "确认这一列到底该是什么类型")
        if empty_pct_warn(empties, total_rows):
            add("low", header, "", "空值率 %.0f%%" % (100.0 * empties / denom),
                "空 %d / 共 %d 行" % (empties, total_rows),
                "空值不填；确认这些行是不是被错误合并或漏导")
        if leading_zero and ID_HEADER_RE.search(header):
            add("info", header, "", "编号列存为文本（正确）",
                "%d 个值带前导零，如 00123" % leading_zero,
                "保持文本；转成数字会变成 123")
        if not header:
            add("low", "列 %d" % (c + 1), "", "这一列没有表头", "清洗和合并都要靠表头对齐", "补列名")

    # whole-column duplicates within one column are not reported; row dupes are above.

    return {
        "sheet": sheet.name,
        "header_row": header_idx + 1,
        "columns": columns,
        "issues": issues,
        "data_rows": data_count,
        "merged": len(merges_in_data),
        "blank_rows": blank_rows,
        "duplicate_rows": dup_rows,
        "repeated_header": repeated_header,
    }


def empty_pct_warn(empties, total):
    if not total:
        return False
    return (100.0 * empties / total) >= 20.0 and empties >= 1


def column_letter(idx):
    s = ""
    n = idx + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------------------------------------------------------------------------
# scoring / output
# ---------------------------------------------------------------------------
WEIGHT = {"high": 10, "medium": 5, "low": 2, "info": 0}


def score_of(issues):
    penalty = sum(WEIGHT.get(i["severity"], 0) for i in issues)
    return max(0, 100 - penalty)


def verdict(score):
    if score >= 90:
        return "基本干净，可以直接用"
    if score >= 70:
        return "能用，但清洗前要先处理标红的高危项"
    if score >= 40:
        return "问题不少，直接分析会得出错结论"
    return "先别分析，先修数据"


def write_report_workbook(path, profiles, meta, source):
    wb = Workbook()
    ws = wb.active
    ws.title = "_summary"
    all_issues = [i for p in profiles for i in p["issues"]]
    counts = Counter(i["severity"] for i in all_issues)
    score = score_of(all_issues)
    rows = [
        ("source", os.path.abspath(source)),
        ("generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("tool", "excel_profile.py v%s" % TOOL_VERSION),
        ("source modified", "no - opened read-only"),
        ("sheets profiled", ", ".join(p["sheet"] for p in profiles)),
        ("data rows", sum(p["data_rows"] for p in profiles)),
        ("columns", sum(len(p["columns"]) for p in profiles)),
        ("score", "%d / 100" % score),
        ("verdict", verdict(score)),
        ("high", counts.get("high", 0)),
        ("medium", counts.get("medium", 0)),
        ("low", counts.get("low", 0)),
        ("info", counts.get("info", 0)),
        ("merged blocks", sum(p["merged"] for p in profiles)),
        ("blank rows", sum(p["blank_rows"] for p in profiles)),
        ("duplicate rows", sum(p["duplicate_rows"] for p in profiles)),
        ("repeated header rows", sum(p["repeated_header"] for p in profiles)),
        ("source encoding", meta.get("encoding", "")),
        ("csv delimiter", repr(meta.get("delimiter", "")) if meta.get("delimiter") else ""),
    ]
    ws.append(["key", "value"])
    for k, v in rows:
        if v == "" or v is None:
            continue
        ws.append([k, v])
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 78

    wc = wb.create_sheet("_columns")
    cols = ["sheet", "column", "header", "filled", "empty", "empty_pct", "unique",
            "types", "numbers_as_text", "dirty_text", "date_formats", "date_labels", "max_len"]
    wc.append(cols)
    for p in profiles:
        for c in p["columns"]:
            wc.append([p["sheet"]] + [c.get(k, "") for k in cols[1:]])
    for i, w in enumerate([16, 8, 26, 8, 7, 9, 7, 34, 14, 11, 12, 34, 8], start=1):
        wc.column_dimensions[column_letter(i - 1)].width = w

    wi = wb.create_sheet("_issues")
    wi.append(["severity", "sheet", "column", "cells", "issue", "detail", "suggestion"])
    order = {"high": 0, "medium": 1, "low": 2, "info": 3}
    for i in sorted(all_issues, key=lambda x: order.get(x["severity"], 9)):
        wi.append([i["severity"], i["sheet"], i["column"], i["where"],
                   i["issue"], i["detail"], i["suggestion"]])
    for i, w in enumerate([10, 16, 22, 18, 34, 60, 46], start=1):
        wi.column_dimensions[column_letter(i - 1)].width = w
    wi.freeze_panes = "A2"
    wc.freeze_panes = "A2"

    wb.save(path)
    return score, counts, all_issues


def write_markdown(path, profiles, source, score, counts, issues):
    L = []
    L.append("# 数据体检报告 · %s" % os.path.basename(source))
    L.append("")
    L.append("- 文件：`%s`" % os.path.abspath(source))
    L.append("- 生成时间：%s" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    L.append("- 工具：excel_profile.py v%s（只读，未改动源文件）" % TOOL_VERSION)
    L.append("- **体检得分：%d / 100** — %s" % (score, verdict(score)))
    L.append("- 问题数：高危 %d · 中 %d · 低 %d · 提示 %d"
             % (counts.get("high", 0), counts.get("medium", 0),
                counts.get("low", 0), counts.get("info", 0)))
    L.append("")
    for p in profiles:
        L.append("## 工作表 `%s`" % p["sheet"])
        L.append("")
        L.append("- 表头在第 %d 行 · 数据 %d 行 · 列 %d"
                 % (p["header_row"], p["data_rows"], len(p["columns"])))
        L.append("- 合并块 %d · 空行 %d · 重复行 %d · 重复表头 %d"
                 % (p["merged"], p["blank_rows"], p["duplicate_rows"], p["repeated_header"]))
        L.append("")
        L.append("| 列 | 表头 | 有值 | 空 | 唯一 | 类型 | 数字存成文本 | 脏文本 | 日期写法 |")
        L.append("|---|---|---|---|---|---|---|---|---|")
        for c in p["columns"]:
            L.append("| %s | %s | %d | %d | %d | %s | %d | %d | %d |"
                     % (c["column"], c["header"] or "（空）", c["filled"], c["empty"],
                        c["unique"], c["types"] or "-", c["numbers_as_text"],
                        c["dirty_text"], c["date_formats"]))
        L.append("")
    if issues:
        L.append("## 问题清单")
        L.append("")
        L.append("| 级别 | 工作表 | 列 | 单元格 | 问题 | 建议 |")
        L.append("|---|---|---|---|---|---|")
        order = {"high": 0, "medium": 1, "low": 2, "info": 3}
        for i in sorted(issues, key=lambda x: order.get(x["severity"], 9)):
            L.append("| %s | %s | %s | %s | %s | %s |"
                     % (i["severity"], i["sheet"], i["column"] or "-", i["where"] or "-",
                        i["issue"], i["suggestion"]))
        L.append("")
    L.append("---")
    L.append("")
    L.append("下一步：按高危项先修数据，再用 `excel_clean.py` 清洗（每处改动都会写进 `_log`）。")
    open(path, "w", encoding="utf-8").write("\n".join(L) + "\n")


HTML_CSS = """
:root{--bg:#f7f8fa;--card:#fff;--ink:#1c2024;--muted:#6b7280;--line:#e5e7eb;
--high:#b91c1c;--medium:#b45309;--low:#0369a1;--info:#4b5563;--ok:#15803d}
*{box-sizing:border-box}
body{margin:0;padding:32px 20px 64px;background:var(--bg);color:var(--ink);
font:15px/1.65 -apple-system,BlinkMacSystemFont,"Segoe UI","Hiragino Sans GB","Microsoft YaHei",sans-serif}
.wrap{max-width:1000px;margin:0 auto}
h1{font-size:26px;margin:0 0 4px}
.sub{color:var(--muted);font-size:13.5px;margin-bottom:24px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;
padding:20px 22px;margin-bottom:18px}
.score{display:flex;align-items:baseline;gap:12px;margin:0 0 8px}
.score b{font-size:44px;line-height:1}
.pill{display:inline-block;padding:2px 9px;border-radius:999px;font-size:12px;font-weight:600}
.high{background:#fee2e2;color:var(--high)}.medium{background:#fef3c7;color:var(--medium)}
.low{background:#e0f2fe;color:var(--low)}.info{background:#f3f4f6;color:var(--info)}
table{width:100%;border-collapse:collapse;font-size:13.5px}
th,td{text-align:left;padding:7px 9px;border-bottom:1px solid var(--line);vertical-align:top}
th{color:var(--muted);font-weight:600;font-size:12.5px;text-transform:uppercase;letter-spacing:.04em}
td.num{text-align:right;font-variant-numeric:tabular-nums}
h2{font-size:17px;margin:0 0 12px}
.foot{color:var(--muted);font-size:12.5px;margin-top:8px}
"""


def write_html(path, profiles, source, score, counts, issues):
    e = lambda s: (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    H = ["<!doctype html><html lang=\"zh-CN\"><meta charset=\"utf-8\">",
         "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">",
         "<title>数据体检 · %s</title><style>%s</style>" % (e(os.path.basename(source)), HTML_CSS),
         "<div class=\"wrap\">",
         "<h1>数据体检报告</h1>",
         "<div class=\"sub\">%s · %s · excel_profile.py v%s · 只读，未改动源文件</div>"
         % (e(os.path.basename(source)), datetime.now().strftime("%Y-%m-%d %H:%M"), TOOL_VERSION),
         "<div class=\"card\"><div class=\"score\"><b>%d</b><span>/ 100</span>"
         "<span class=\"pill %s\">%s</span></div>"
         % (score, "high" if score < 70 else ("medium" if score < 90 else "info"), e(verdict(score))),
         "<div class=\"sub\" style=\"margin:0\">高危 %d · 中 %d · 低 %d · 提示 %d</div></div>"
         % (counts.get("high", 0), counts.get("medium", 0), counts.get("low", 0), counts.get("info", 0))]

    for p in profiles:
        H.append("<div class=\"card\"><h2>工作表 <code>%s</code></h2>" % e(p["sheet"]))
        H.append("<div class=\"sub\">表头第 %d 行 · 数据 %d 行 · 列 %d · 合并块 %d · 空行 %d · 重复行 %d</div>"
                 % (p["header_row"], p["data_rows"], len(p["columns"]), p["merged"],
                    p["blank_rows"], p["duplicate_rows"]))
        H.append("<table><tr><th>列</th><th>表头</th><th>有值</th><th>空</th><th>唯一</th>"
                 "<th>类型</th><th>数字存文本</th><th>脏文本</th><th>日期写法</th></tr>")
        for c in p["columns"]:
            H.append("<tr><td>%s</td><td>%s</td><td class=\"num\">%d</td><td class=\"num\">%d</td>"
                     "<td class=\"num\">%d</td><td>%s</td><td class=\"num\">%d</td>"
                     "<td class=\"num\">%d</td><td class=\"num\">%d</td></tr>"
                     % (c["column"], e(c["header"] or "（空）"), c["filled"], c["empty"],
                        c["unique"], e(c["types"] or "-"), c["numbers_as_text"],
                        c["dirty_text"], c["date_formats"]))
        H.append("</table></div>")

    if issues:
        order = {"high": 0, "medium": 1, "low": 2, "info": 3}
        H.append("<div class=\"card\"><h2>问题清单（%d）</h2><table>" % len(issues))
        H.append("<tr><th>级别</th><th>工作表</th><th>列</th><th>单元格</th><th>问题</th><th>建议</th></tr>")
        for i in sorted(issues, key=lambda x: order.get(x["severity"], 9)):
            H.append("<tr><td><span class=\"pill %s\">%s</span></td><td>%s</td><td>%s</td>"
                     "<td>%s</td><td>%s<br><span class=\"sub\">%s</span></td><td>%s</td></tr>"
                     % (i["severity"], i["severity"], e(i["sheet"]), e(i["column"] or "-"),
                        e(i["where"] or "-"), e(i["issue"]), e(i["detail"]), e(i["suggestion"])))
        H.append("</table></div>")

    H.append("<div class=\"foot\">下一步：先修高危项，再用 excel_clean.py 清洗（每处改动写进 _log）。</div>")
    H.append("</div></html>")
    open(path, "w", encoding="utf-8").write("".join(H))


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(
        description="Read-only data-quality report for an Excel/CSV file. Never edits the source.")
    ap.add_argument("--src", required=True, help="input .xlsx / .xlsm / .xls / .csv")
    ap.add_argument("--sheet", help="profile only this sheet")
    ap.add_argument("--header-row", type=int, help="header is on row N (default: auto-detect)")
    ap.add_argument("--out", help="report workbook (default: <name>_profile.xlsx)")
    ap.add_argument("--md", help="also write a Markdown report to this path")
    ap.add_argument("--html", help="also write a self-contained HTML report to this path")
    ap.add_argument("--json", help="also write the findings as JSON to this path")
    ap.add_argument("--strict", action="store_true", help="exit 2 if any high-severity issue exists")
    ap.add_argument("--max-rows", type=int, default=DEFAULT_MAX_ROWS, help="cap rows read per sheet")
    args = ap.parse_args()

    if not os.path.exists(args.src):
        sys.exit("No such file: %s" % args.src)

    sheets, meta = load_source(args.src, args.sheet, args.max_rows)

    profiles = []
    for sh in sheets:
        if not sh.rows:
            continue
        hidx = detect_header(sh.rows, args.header_row)
        profiles.append(profile_sheet(sh, hidx, args.max_rows))

    if not profiles:
        sys.exit("Nothing to profile: every sheet is empty.")

    base = os.path.splitext(os.path.abspath(args.src))[0]
    out = args.out or (base + "_profile.xlsx")
    score, counts, issues = write_report_workbook(out, profiles, meta, args.src)

    print("=" * 78)
    print("excel_profile.py v%s  |  %s" % (TOOL_VERSION, os.path.basename(args.src)))
    print("=" * 78)
    print("source untouched: opened read-only")
    for p in profiles:
        print("")
        print("sheet %-18s header row %-3d data rows %-6d columns %d"
              % (p["sheet"], p["header_row"], p["data_rows"], len(p["columns"])))
        for c in p["columns"]:
            flags = []
            if c["numbers_as_text"]:
                flags.append("%d num-as-text" % c["numbers_as_text"])
            if c["dirty_text"]:
                flags.append("%d dirty" % c["dirty_text"])
            if c["date_formats"] > 1:
                flags.append("%d date formats" % c["date_formats"])
            if c["empty_pct"] >= 20:
                flags.append("%.0f%% empty" % c["empty_pct"])
            if flags:
                print("    %-3s %-24s %s" % (c["column"], (c["header"] or "-")[:24], ", ".join(flags)))
    print("")
    print("-" * 78)
    print("score %d / 100  -  %s" % (score, verdict(score)))
    print("high %d  medium %d  low %d  info %d"
          % (counts.get("high", 0), counts.get("medium", 0),
             counts.get("low", 0), counts.get("info", 0)))
    order = {"high": 0, "medium": 1, "low": 2, "info": 3}
    for i in sorted(issues, key=lambda x: order.get(x["severity"], 9))[:12]:
        col = (" [%s]" % i["column"]) if i["column"] else ""
        print("  %-6s %s%s" % (i["severity"], i["issue"], col))
    if len(issues) > 12:
        print("  ... %d more, see the _issues sheet" % (len(issues) - 12))
    print("-" * 78)
    print("report: %s" % os.path.abspath(out))

    if args.md:
        write_markdown(args.md, profiles, args.src, score, counts, issues)
        print("markdown: %s" % os.path.abspath(args.md))
    if args.html:
        write_html(args.html, profiles, args.src, score, counts, issues)
        print("html    : %s" % os.path.abspath(args.html))
    if args.json:
        payload = {
            "source": os.path.abspath(args.src),
            "generated": datetime.now().isoformat(timespec="seconds"),
            "tool": "excel_profile.py v" + TOOL_VERSION,
            "source_modified": False,
            "score": score,
            "verdict": verdict(score),
            "counts": dict(counts),
            "sheets": [
                {"sheet": p["sheet"], "header_row": p["header_row"], "data_rows": p["data_rows"],
                 "columns": p["columns"],
                 "sheet_issues": [i for i in p["issues"] if not i["column"]]}
                for p in profiles
            ],
            "issues": issues,
        }
        open(args.json, "w", encoding="utf-8").write(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
        print("json    : %s" % os.path.abspath(args.json))

    if args.strict and counts.get("high", 0):
        print("")
        print("STRICT: %d high-severity issue(s) - exiting 2" % counts["high"])
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
