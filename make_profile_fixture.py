#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build a deliberately broken workbook so excel_profile.py has something to find."""
import os
from datetime import datetime

from openpyxl import Workbook

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tmp_profile_test")
os.makedirs(OUT_DIR, exist_ok=True)

wb = Workbook()

# --- sheet 1: the messy one ------------------------------------------------
ws = wb.active
ws.title = "明细"

# merged header block at the top, like a report exported from an ERP
ws["A1"] = "2026 年 3 月 订单明细"
ws.merge_cells("A1:F1")
ws["A2"] = "导出时间: 2026-04-01 09:12"

# real header on row 3, with a duplicate column name and one empty name
headers = ["客户编号", "客户名称", "金额", "下单日期", "跟进人", None]
for c, h in enumerate(headers, start=1):
    ws.cell(row=3, column=c, value=h)

rows = [
    # id keeps leading zero; money stored as text with a thousands separator
    ["00123", " 张三 ", "1,280.50", "2026-03-01", "小王", None],
    ["00124", "李四", "980", "2026/3/2", "小李", None],
    ["00125", "王五", "¥ 1,050 元", "05.03.2026", "小王", None],
    # full-width digits and a non-breaking space - classic copy/paste damage
    ["00126", "赵\u00a0六", "７８０", "2026年3月5日", "小张", None],
    # the header printed again in the middle of the data
    ["客户编号", "客户名称", "金额", "下单日期", "跟进人", None],
    ["00127", "钱七", "2,300", "20260307", "小李", None],
    # exactly the same row twice
    ["00128", "孙八", "560", "2026-03-08", "小王", None],
    ["00128", "孙八", "560", "2026-03-08", "小王", None],
    # a blank row splitting the table
    [None, None, None, None, None, None],
    # a row wider than the header
    ["00129", "周九", "3.5", "2026-03-09", "小张", "多出来的一列"],
]

r = 4
for row in rows:
    for c, v in enumerate(row, start=1):
        ws.cell(row=r, column=c, value=v)
    r += 1

# a second header block, to catch "header repeated" and a merged block in the data
ws.merge_cells("B16:C16")
ws["B16"] = "备注合并块"

# --- sheet 2: a cover page -------------------------------------------------
cover = wb.create_sheet("封面")
cover["A1"] = "月度报表"
cover["A2"] = "内部资料"

# --- sheet 3: clean, so the tool has to say something nice ------------------
clean = wb.create_sheet("干净表")
clean.append(["order_id", "amount", "date"])
clean.append(["A-1001", 120.5, datetime(2026, 3, 1)])
clean.append(["A-1002", 88.0, datetime(2026, 3, 2)])
clean.append(["A-1003", 240.0, datetime(2026, 3, 3)])

path = os.path.join(OUT_DIR, "dirty_sample.xlsx")
wb.save(path)
print("wrote", path, os.path.getsize(path), "bytes")
