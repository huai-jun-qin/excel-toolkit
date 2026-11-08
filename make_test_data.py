#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_test_data.py - Build a deliberately dirty test set.

Every kind of mess a real client's folder contains is planted on purpose:
  - the same column spelled 4 different ways (incl. one Chinese file)
  - dates as real dates, as text "2026/3/5", "05.03.2026", "2026年3月5日", and as Excel serials
  - numbers stored as text: "1,234.50", "¥ 88 元"
  - IDs with leading zeros that must survive: "00123"
  - trailing spaces, non-breaking spaces, full-width characters
  - blank rows, duplicate rows, and a sheet whose header starts on row 3
  - one .csv, one fake .xls (to prove errors are reported), one corrupt file
  - samples/legacy/: a real .xls, so .xls support can be checked end to end

Run:  python3 make_test_data.py
"""

import csv
import os
import random
import shutil
from datetime import datetime

from openpyxl import Workbook

random.seed(7)

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "samples", "dirty")

REGIONS = ["North", "South", "East", "West", "华东", "华南"]

# the same logical columns, spelled differently per file
HEADER_VARIANTS = [
    ["Order ID", "Customer Name", "Email", "Order Date", "Amount", "Qty", "Region"],
    ["order_id", "customer_name", "e-mail", "date", "amount", "quantity", "region"],
    ["订单号", "客户名称", "邮箱", "日期", "金额", "数量", "地区"],
    ["Order No.", "Client Name", "Email ", "Order Date ", "Total", "QTY", "Region "],
    ["OrderID", "Full Name", "Mail", "Day", "Sum", "Count", "Area"],
]


def dirty_name(i):
    names = ["Alice Chen", "  Bob Li  ", "卡罗尔", "David Wu\u00a0", "Eve Zhang",
             "Ｆｒａｎｋ　Ｚｈｏｕ", "Grace Kim", "Heidi Song", "Ivan Pavlov", "Judy Lin"]
    return names[i % len(names)]


def build_file(path, idx):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    headers = HEADER_VARIANTS[idx % len(HEADER_VARIANTS)]

    if idx == 6:
        # header on row 3, with two junk rows above it
        ws.append(["Exported from the old ERP - do not edit"])
        ws.append([])
        ws.append(headers)
    else:
        ws.append(headers)

    rows = 8 + idx * 3
    for k in range(rows):
        oid = "%05d" % (100 + idx * 100 + k)
        if idx == 2:                       # leading zeros, must stay text
            oid = "00%03d" % (k + 1)
        cust = dirty_name(idx + k)
        mail = "user%d@example.com" % (idx * 10 + k)
        # five different date representations
        kind = (idx + k) % 5
        if kind == 0:
            dt = datetime(2026, 1 + (k % 12), 1 + (k % 27))
        elif kind == 1:
            dt = "2026/%d/%d" % (1 + (k % 12), 1 + (k % 27))
        elif kind == 2:
            dt = "%02d.%02d.2026" % (1 + (k % 27), 1 + (k % 12))   # DD.MM.YYYY
        elif kind == 3:
            dt = "2026年%d月%d日" % (1 + (k % 12), 1 + (k % 27))
        else:
            dt = 46000 + k                 # bare Excel serial

        amount = [1234.5, "1,234.50", "¥ 88 元", 450, "88.8"][(idx + k) % 5]
        qty = ["3", 3, " 3 ", "12", 12][(idx + k) % 5]
        region = REGIONS[(idx + k) % len(REGIONS)]
        ws.append([oid, cust, mail, dt, amount, qty, region])

    # a fully blank row in the middle of the data
    if idx % 2 == 0:
        ws.append([])

    # duplicate the last data row (dedupe test)
    if idx % 3 == 0:
        last = [c.value for c in ws[ws.max_row]]
        ws.append(last)

    wb.save(path)


def main():
    if os.path.isdir(OUT):
        shutil.rmtree(OUT)
    os.makedirs(OUT, exist_ok=True)

    for i in range(8):
        build_file(os.path.join(OUT, "orders_%02d.xlsx" % (i + 1)), i)

    # a csv with a different header spelling + semicolon delimiter
    with open(os.path.join(OUT, "orders_09.csv"), "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Order No", "Client Name", "Email", "Order Date", "Total", "QTY", "Region"])
        for k in range(6):
            w.writerow(["009%02d" % k, "客户%d" % k, "c%d@example.com" % k,
                        "2026-02-%02d" % (k + 1), "1,00%d.00" % k, k + 1, REGIONS[k % 6]])

    # a legacy .xls (really just a text file) -> must produce a readable error, not a crash
    with open(os.path.join(OUT, "old_export.xls"), "w", encoding="utf-8") as f:
        f.write("This is a fake old-format file.\n")

    # a corrupt xlsx -> must be caught and reported, other files still merged
    with open(os.path.join(OUT, "broken.xlsx"), "wb") as f:
        f.write(b"PK\x03\x04 definitely not a zip")

    print("dirty test set -> %s" % OUT)
    for n in sorted(os.listdir(OUT)):
        print("  %-20s %d bytes" % (n, os.path.getsize(os.path.join(OUT, n))))

    build_legacy_xls()

    print("\nExpected outcome: merge reports the fake .xls + broken.xlsx as FAILED and still merges the rest.")
    print("The real .xls in samples/legacy must merge cleanly, dates included.")


def build_legacy_xls():
    """A genuine Excel 97-2003 file, to prove .xls is read and dates survive.

    Needs xlwt (write-only, used here just to generate the sample).
    """
    try:
        import xlwt
    except ImportError:
        print("\nsamples/legacy skipped: pip install xlwt   (only needed to build this sample)")
        return

    out_dir = os.path.join(HERE, "samples", "legacy")
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, "legacy_export.xls")

    wb = xlwt.Workbook()
    ws = wb.add_sheet("Sheet1")
    # a title line above the header - the classic ERP export shape
    ws.write(0, 0, "Exported from the old ERP system - do not edit")
    for c, h in enumerate(["订单号", "客户名称", "日期", "金额", "数量", "地区"]):
        ws.write(2, c, h)

    dstyle = xlwt.XFStyle(); dstyle.num_format_str = "YYYY-MM-DD"
    mstyle = xlwt.XFStyle(); mstyle.num_format_str = "#,##0.00"

    rows = [
        ("00123", "张三", datetime(2026, 3, 5), 1234.5, 3, "华东"),   # real date cell
        ("00124", "李四", "2026/4/6", "1,200.50", 2, "华南"),          # text date + text amount
        ("00125", "Wang Wu", "05.03.2026", "¥ 88 元", "5", "North"),
        ("00126", "赵六", datetime(2026, 6, 7), 300, 1, "华东"),
    ]
    for i, (oid, name, dt, amt, qty, region) in enumerate(rows, start=3):
        ws.write(i, 0, oid)
        ws.write(i, 1, name)
        if isinstance(dt, datetime):
            ws.write(i, 2, dt, dstyle)
        else:
            ws.write(i, 2, dt)
        if isinstance(amt, (int, float)):
            ws.write(i, 3, amt, mstyle)
        else:
            ws.write(i, 3, amt)
        ws.write(i, 4, qty)
        ws.write(i, 5, region)

    wb.save(path)
    print("\nlegacy .xls -> %s  (%d bytes)" % (path, os.path.getsize(path)))


if __name__ == "__main__":
    main()
