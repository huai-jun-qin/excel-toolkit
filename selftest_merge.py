#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""selftest_merge.py - build deliberately dirty workbooks, then assert the merge
handles them. Run it after any change to excel_merge.py:

    python3 selftest_merge.py

Every case is a real-world shape that used to break:
  1. cover page + data page      (cover text was merged as one garbage row)
  2. January..March in one file  (only the first month was read, silently)
  3. header not on row 1         (an info row with 4 cells beat the 5-column header)
  4. header printed twice        (header landed in the data)
  5. CSV with BOM + semicolons
  6. a sheet named 汇总           (double counting if merged)
"""

import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
# MERGE_SCRIPT lets you run the same cases against an older copy of the tool.
MERGE = os.environ.get("MERGE_SCRIPT") or os.path.join(HERE, "excel_merge.py")

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    sys.exit("openpyxl missing. Run: %s -m pip install openpyxl" % sys.executable)


def cover_and_data(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "封面"
    ws["A1"] = "2026年3月对账单"
    ws["A3"] = "制表：李四"
    ws["A4"] = "电话：13800000000"
    ws["A5"] = "本表数据以实际到账为准，如有疑问请在三日内联系财务部"
    d = wb.create_sheet("明细")
    d.append(["客户名称", "金额", "日期"])
    for i in range(3):
        d.append(["客户%d" % (i + 1), 100 + i, "2026-03-%02d" % (i + 1)])
    wb.save(path)


def months(path):
    wb = Workbook()
    first = True
    for m in (1, 2, 3):
        ws = wb.active if first else wb.create_sheet()
        ws.title = "%d月" % m
        first = False
        ws.append(["客户名称", "金额"])
        for i in range(2):
            ws.append(["M%d客户%d" % (m, i + 1), m * 10 + i])
    wb.save(path)


def header_below_info(path):
    """The nasty one: a 4-cell info line sitting right above a 7-column header.

    Width alone cannot tell them apart (4 >= 7*0.5), so the header detector has
    to notice that '单位：元' is an info label, not a column name.
    """
    wb = Workbook()
    ws = wb.active
    ws.append([])
    ws.append(["单位：元", "期间：2026年1-3月", "制表：李四", "第1页"])
    ws.append(["客户名称", "金额", "日期", "备注", "经办人", "地区", "单号"])
    for i in range(3):
        ws.append(["客户%d" % i, i * 5, "2026-01-0%d" % (i + 1), "无", "张三", "华东", "A00%d" % i])
    wb.save(path)


def repeated_header(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户名称", "金额"])
    ws.append(["甲", 10])
    ws.append(["乙", 20])
    ws.append(["客户名称", "金额"])
    ws.append(["丙", 30])
    wb.save(path)


def csv_bom(path):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        f.write("客户名称;金额;日期\n")
        f.write("王五;88;2026-02-01\n")
        f.write("赵六;99;2026-02-02\n")


def summary_sheet(path):
    wb = Workbook()
    d = wb.active
    d.title = "明细"
    d.append(["客户名称", "金额"])
    for i in range(4):
        d.append(["客户%d" % i, 10 + i])
    s = wb.create_sheet("汇总")
    s.append(["项目", "合计"])
    s.append(["总额", 46])
    wb.save(path)


def col0(wb):
    ws = wb["Merged"]
    return [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]


def header_a1(wb):
    return wb["Merged"].cell(row=1, column=1).value


def want_header(wb):
    got = header_a1(wb)
    return [] if got == "客户名称" else ["A1 is %r, expected '客户名称'" % (got,)]


def want_months(wb):
    vals = [str(v) for v in col0(wb)]
    bad = [m for m in ("M1客户1", "M2客户1", "M3客户1") if m not in vals]
    return ["missing %s" % ", ".join(bad)] if bad else []


def want_no_summary(wb):
    vals = [str(v) for v in col0(wb)]
    return ["summary row '总额' was merged"] if "总额" in vals else []


def want_no_header_row(wb):
    vals = [str(v) for v in col0(wb)]
    n = sum(1 for v in vals if v == "客户名称")
    return ["%d data rows repeat the header text" % n] if n else []


def want_one_month(wb):
    vals = [str(v) for v in col0(wb)]
    bad = [m for m in ("M2客户1", "M3客户1") if m in vals]
    return ["--first-sheet still read %s" % ", ".join(bad)] if bad else []


def cover_only_folder(src):
    """A folder holding one real file and one 'template/instructions' file.

    The instructions file has no table at all. It must contribute zero rows -
    a title line merged into the data is worse than reading nothing.
    """
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "填写说明"
    ws["A2"] = "本文件仅为模板说明，实际数据见另一份附件，请勿修改格式"
    wb.save(os.path.join(src, "封面说明.xlsx"))

    good = Workbook()
    g = good.active
    g.append(["客户名称", "金额"])
    g.append(["甲", 1])
    g.append(["乙", 2])
    good.save(os.path.join(src, "数据.xlsx"))


def want_no_cover_text(wb):
    vals = [str(v) for v in col0(wb)]
    bad = [v for v in vals if "说明" in v or "模板" in v]
    return ["cover text merged: %s" % bad] if bad else []


def single_col_list(path):
    """500 emails in one column. This is data, not a cover page."""
    wb = Workbook()
    ws = wb.active
    ws.append(["邮箱"])
    for i in range(50):
        ws.append(["user%d@example.com" % i])
    wb.save(path)


def tiny_table(path):
    """A 2x2 table. Small is not the same as a cover page."""
    wb = Workbook()
    ws = wb.active
    ws.append(["客户名称", "金额"])
    ws.append(["甲", 1])
    wb.save(path)


def want_a1(text):
    def check(wb):
        got = header_a1(wb)
        return [] if got == text else ["A1 is %r, expected %r" % (got, text)]
    return check


CASES = [
    ("cover page + data page", "a.xlsx", cover_and_data, 3, [], [want_header, want_no_header_row]),
    ("3 months in one file", "b.xlsx", months, 6, [], [want_months]),
    ("header below an info row", "c.xlsx", header_below_info, 3, [], [want_header]),
    ("header printed twice", "d.xlsx", repeated_header, 3, [], [want_header, want_no_header_row]),
    ("csv with BOM + semicolons", "e.csv", csv_bom, 2, [], [want_header]),
    ("summary sheet skipped", "f.xlsx", summary_sheet, 4, [], [want_no_summary]),
    ("first-sheet mode reads 1 month", "b.xlsx", months, 2, ["--first-sheet"], [want_one_month]),
    ("instructions file adds 0 rows", None, cover_only_folder, 2, [], [want_header, want_no_cover_text]),
    ("single-column list stays data", "g.xlsx", single_col_list, 50, [], [want_a1("邮箱")]),
    ("2x2 table stays data", "h.xlsx", tiny_table, 1, [], [want_header]),
]


def merged_rows(path):
    wb = load_workbook(path)
    ws = wb["Merged"]
    return ws.max_row - 1, wb


def check_row(wb):
    rs = wb["_report"]
    for row in rs.iter_rows(values_only=True):
        if row and row[0] == "CHECK":
            return " ".join(str(v) for v in row[1:] if v not in (None, ""))
    return ""


def main():
    root = tempfile.mkdtemp(prefix="merge_selftest_")
    fails = 0
    print("selftest: %s   (%s)" % (os.path.basename(MERGE), os.path.basename(sys.executable)))
    print("=" * 78)
    try:
        for i, case in enumerate(CASES, 1):
            name, fname, builder, expect, extra, checks = case
            d = os.path.join(root, "case%02d" % i)
            src = os.path.join(d, "in")
            os.makedirs(src, exist_ok=True)
            if fname is None:
                builder(src)                    # builder fills the whole folder
            else:
                builder(os.path.join(src, fname))
            out = os.path.join(d, "out.xlsx")
            p = subprocess.run(
                [sys.executable, MERGE, "--src", src, "--out", out] + extra,
                capture_output=True, text=True)
            if p.returncode != 0:
                print("[FAIL] %-32s merge exited %d" % (name, p.returncode))
                print(((p.stdout or "") + (p.stderr or ""))[-700:])
                fails += 1
                continue
            got, wb = merged_rows(out)
            problems = []
            if got != expect:
                problems.append("rows %d, expected %d" % (got, expect))
            for fn in checks:
                problems.extend(fn(wb))
            ck = check_row(wb)
            if not ck.startswith("OK"):
                problems.append("report CHECK=%r" % ck[:40])
            if problems:
                print("[FAIL] %-32s %s" % (name, "; ".join(problems)))
                fails += 1
            else:
                print("[PASS] %-32s rows %d" % (name, got))
    finally:
        shutil.rmtree(root, ignore_errors=True)

    print("=" * 78)
    print("%d passed, %d failed" % (len(CASES) - fails, fails))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
