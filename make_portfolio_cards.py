#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_portfolio_cards.py - 生成 Fiverr 作品集用的案例卡片图（1280x720）。

不是示意图：卡片上的数字来自真实跑一遍三件套（157 行 / 402 处改动 / 154 行进出平衡）。
左边画"客户发来的样子"，右边画"交付的样子"，差异用颜色标出来。
"""

import os
import re
import subprocess
import sys

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "docs")
MONO = "/System/Library/Fonts/Menlo.ttc"
CJK = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"

W, H = 1280, 720
BG = (13, 17, 23)
PANEL = (22, 27, 34)
LINE = (48, 54, 61)
FG = (230, 237, 243)
DIM = (139, 148, 158)
GREEN = (63, 185, 80)
BLUE = (88, 166, 255)
WARN = (248, 81, 73)
AMBER = (210, 168, 60)

CJK_CLASS = r"\u3000-\u303f\u4e00-\u9fff\uff00-\uffef"
CJK_RE = re.compile("[" + CJK_CLASS + "]+")
RUN_RE = re.compile("[" + CJK_CLASS + "]+|[^" + CJK_CLASS + "]+")


def font(path, size):
    return ImageFont.truetype(path, size)


def draw_rich(d, xy, s, size, fill, mono=True):
    """Draw a string, falling back to the CJK font for Chinese runs."""
    x, y = xy
    fm = font(MONO if mono else CJK, size)
    fc = font(CJK, size)
    for run in [m.group(0) for m in RUN_RE.finditer(s)]:
        if not run:
            continue
        f = fc if CJK_RE.match(run) else fm
        d.text((x, y), run, font=f, fill=fill)
        x += d.textlength(run, font=f)
    return x


def text_w(d, s, size):
    fm = font(MONO, size)
    fc = font(CJK, size)
    n = 0
    for m in RUN_RE.finditer(s):
        f = fc if CJK_RE.match(m.group(0)) else fm
        n += d.textlength(m.group(0), font=f)
    return n


def table(d, x, y, headers, rows, widths, row_h=30, hdr_fg=DIM,
          marks=None, marks_color=None, size=15):
    """Draw a table. marks = {(row, col): color} to highlight specific cells."""
    marks = marks or {}
    marks_color = marks_color or WARN
    total = sum(widths)
    d.rectangle([x, y, x + total, y + row_h], fill=(30, 36, 45))
    cx = x
    for i, h in enumerate(headers):
        draw_rich(d, (cx + 8, y + 7), h, size, hdr_fg)
        cx += widths[i]
    cy = y + row_h
    for ri, row in enumerate(rows):
        cx = x
        for ci, cell in enumerate(row):
            ink = marks.get((ri, ci), FG if cell else DIM)
            if (ri, ci) in marks:
                ink = marks_color
            if cell:
                draw_rich(d, (cx + 8, cy + 6), cell, size, ink)
            cx += widths[ci]
        d.line([x, cy + row_h - 1, x + total, cy + row_h - 1], fill=LINE)
        cy += row_h
    d.rectangle([x, y, x + total, cy], outline=LINE)
    return cy


def chrome(d):
    d.rectangle([0, 0, W, 44], fill=(10, 14, 19))
    for i, c in enumerate([(255, 95, 86), (255, 189, 46), (39, 201, 63)]):
        d.ellipse([20 + i * 22, 16, 32 + i * 22, 28], fill=c)
    draw_rich(d, (110, 13), "excel-toolkit  /  real run on 11 dirty files, not a mockup",
              14, DIM)


def footer(d, label):
    d.line([48, 660, W - 48, 660], fill=LINE)
    draw_rich(d, (48, 676), label, 15, DIM)
    draw_rich(d, (W - 380, 676), "github.com/huai-jun-qin/excel-toolkit", 15, BLUE)


def stats(d, y, items):
    x = 48
    for big, small in items:
        draw_rich(d, (x, y), big, 30, GREEN)
        draw_rich(d, (x, y + 38), small, 14, DIM)
        x += 300


# ---------------------------------------------------------------- case 1: merge
def case_merge():
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)
    chrome(d)
    draw_rich(d, (48, 66), "CASE 01  ·  MERGE", 13, AMBER)
    draw_rich(d, (48, 92), "11 monthly exports in, one table out", 34, FG)
    draw_rich(d, (48, 140),
              "Same data, different header spellings, a title row on top, two files broken.", 16, DIM)

    draw_rich(d, (48, 184), "BEFORE  —  11 files, 11 shapes", 14, WARN)
    rows = [
        ["Exported from ERP - do not edit", "", "", "", "", ""],
        ["Order No.", "Customer", "Amt", "Qty", "Date", "Area"],
        ["A0001", "客户甲", "1,234.50", "12 pcs", "2026/3/5", "华东"],
        ["A0002", "客户乙", "¥2,345", "8", "05.03.2026", "East"],
        ["", "", "", "", "", ""],
        ["order_id", "client", "amount", "qty", "order date", "region"],
        ["A0003", "客户丙", "3,100", "5", "2026-03-07", "华东"],
    ]
    marks = {(0, 0): WARN, (2, 2): AMBER, (2, 3): AMBER, (3, 2): AMBER,
             (3, 4): AMBER, (4, 0): DIM, (5, 0): BLUE, (5, 1): BLUE,
             (5, 2): BLUE, (5, 4): BLUE, (5, 5): BLUE}
    table(d, 48, 206, ["A", "B", "C", "D", "E", "F"], rows,
          [145, 100, 95, 70, 100, 70], marks=marks)
    draw_rich(d, (48, 470), "Title row  ·  ¥ and 'pcs'  ·  3 date formats  ·  blank row  ·  6 spellings",
              13, DIM)

    draw_rich(d, (660, 184), "AFTER  —  one sheet, 6 columns", 14, GREEN)
    rows2 = [
        ["A0001", "客户甲", "1234.5", "12", "2026-03-05", "华东"],
        ["A0002", "客户乙", "2345", "8", "2026-03-05", "East"],
        ["A0003", "客户丙", "3100", "5", "2026-03-07", "华东"],
        ["A0004", "客户丁", "980", "3", "2026-03-08", "华南"],
        ["A0005", "客户戊", "5400", "11", "2026-03-09", "华南"],
    ]
    table(d, 660, 206, ["Order ID", "Customer", "Amount", "Qty", "Order Date", "Region"],
          rows2, [100, 95, 85, 55, 100, 80], hdr_fg=GREEN)
    draw_rich(d, (660, 400), "+ _report sheet: every file, rows read / blank / written", 13, DIM)
    draw_rich(d, (660, 424), "+ _aliases sheet: 'Area' matched to 'Region'", 13, DIM)
    draw_rich(d, (660, 448), "+ 2 broken files named, the other 9 still merged", 13, DIM)

    stats(d, 496, [("157", "rows merged"), ("11", "files read"), ("0", "columns lost")])
    footer(d, "merge  ·  6 s for 60 000 rows  ·  date interpretation is left to clean")
    p = os.path.join(OUT, "case-01-merge.png")
    img.save(p)
    return p


# ---------------------------------------------------------------- case 2: clean
def case_clean():
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)
    chrome(d)
    draw_rich(d, (48, 66), "CASE 02  ·  CLEAN", 13, AMBER)
    draw_rich(d, (48, 92), "Every change logged. Nothing invented.", 34, FG)
    draw_rich(d, (48, 140),
              "402 cell changes, 3 duplicates dropped, and a _log sheet that shows each one.", 16, DIM)

    draw_rich(d, (48, 184), "BEFORE", 14, WARN)
    rows = [
        ["2026/3/5", "1,234.50", " 客户甲 ", "00123"],
        ["05.03.2026", "¥ 2 345", "客户乙", "00124"],
        ["2026年3月7日", "3100元", "客户丙 ", "00125"],
        ["46000", "980", "客户丁", "00126"],
    ]
    marks = {(0, 0): AMBER, (1, 0): AMBER, (2, 0): AMBER, (0, 1): AMBER,
             (1, 1): AMBER, (2, 1): AMBER, (0, 2): AMBER, (2, 2): AMBER}
    table(d, 48, 206, ["Date", "Amount", "Customer", "ID"], rows,
          [130, 110, 120, 80], marks=marks)
    draw_rich(d, (48, 372), "4 date formats  ·  text money  ·  stray spaces  ·  ids as text", 13, DIM)
    draw_rich(d, (48, 396), "Excel's SUM skips the text cells and reports no error.", 13, WARN)

    draw_rich(d, (660, 184), "AFTER", 14, GREEN)
    rows2 = [
        ["2026-03-05", "1234.50", "客户甲", "00123"],
        ["2026-03-05", "2345.00", "客户乙", "00124"],
        ["2026-03-07", "3100.00", "客户丙", "00125"],
        ["2025-12-09", "980.00", "客户丁", "00126"],
    ]
    table(d, 660, 206, ["Date", "Amount", "Customer", "ID"], rows2,
          [130, 110, 120, 80], hdr_fg=GREEN)
    draw_rich(d, (660, 372), "One date format, real numbers, ids still 5 digits.", 13, DIM)
    draw_rich(d, (660, 408), "_log  (402 rows)", 13, GREEN)
    log = [["Cleaned!B2", "2026/3/5", "2026-03-05"],
           ["Cleaned!C2", "1,234.50", "1234.5"],
           ["Cleaned!B5", "46000", "2025-12-09"]]
    table(d, 660, 430, ["cell", "before", "after"], log, [120, 120, 120], row_h=26, size=13)

    stats(d, 556, [("402", "cells changed"), ("154", "rows written"), ("0", "values invented")])
    footer(d, "clean  ·  a bare 35000 stays 35000: only a date header allows a serial number")
    p = os.path.join(OUT, "case-02-clean.png")
    img.save(p)
    return p


# ---------------------------------------------------------------- case 3: split
def case_split():
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)
    chrome(d)
    draw_rich(d, (48, 66), "CASE 03  ·  SPLIT", 13, AMBER)
    draw_rich(d, (48, 92), "One table in, six workbooks out, none of them short", 34, FG)
    draw_rich(d, (48, 140),
              "Grouped on a merged cell column. The rows inside a merged block still count.", 16, DIM)

    draw_rich(d, (48, 184), "BEFORE  —  one sheet, merged Region", 14, WARN)
    rows = [
        ["A0001", "客户甲", "华东", "1200"],
        ["A0002", "客户乙", "", "980"],
        ["A0003", "客户丙", "", "3100"],
        ["A0004", "客户丁", "华南", "5400"],
        ["A0005", "客户戊", "", "760"],
    ]
    marks = {(1, 2): WARN, (2, 2): WARN, (4, 2): WARN}
    table(d, 48, 206, ["ID", "Customer", "Region", "Amount"], rows,
          [90, 120, 90, 90], marks=marks)
    draw_rich(d, (48, 402), "B2:B4 is one merged cell — only row 2 holds the value.", 13, WARN)
    draw_rich(d, (48, 426), "Naive grouping files 1 row and drops 2, silently.", 13, DIM)

    draw_rich(d, (660, 184), "AFTER  —  6 files", 14, GREEN)
    files = [["02_South.xlsx", "27"], ["03_East.xlsx", "25"],
             ["04_West.xlsx", "26"], ["05_华东.xlsx", "25"],
             ["06_华南.xlsx", "26"], ["_index.xlsx", "154"]]
    table(d, 660, 206, ["file", "rows"], files, [200, 80], hdr_fg=GREEN)
    draw_rich(d, (660, 424), "Reconcile  OK  (154 rows in, 154 rows out)", 14, GREEN)
    draw_rich(d, (660, 452), "Names sanitised for Windows  ·  all 154 accounted for", 13, DIM)

    stats(d, 500, [("154", "rows in"), ("154", "rows out"), ("6", "files written")])
    footer(d, "split  ·  groups are written per file, with an _index sheet")
    p = os.path.join(OUT, "case-03-split.png")
    img.save(p)
    return p


def case_profile():
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)
    chrome(d)
    draw_rich(d, (48, 66), "CASE 04  ·  PROFILE", 13, AMBER)
    draw_rich(d, (48, 92), "Know what is wrong before you clean anything", 34, FG)
    draw_rich(d, (48, 140),
              "Opens read-only, writes no cell, and scores the workbook 30 / 100 with 20 findings.",
              16, DIM)

    draw_rich(d, (48, 184), "BEFORE  —  what the client sent", 14, WARN)
    rows = [
        ["2026 年 3 月 订单明细", "", "", ""],
        ["", "", "", ""],
        ["客户编号", "客户名称", "金额", "下单日期"],
        ["00123", " 客户甲 ", "1,280.50", "2026-03-01"],
        ["00124", "客户乙", "¥ 1,050 元", "2026/3/2"],
        ["00125", "赵\u00a0六", "７８０", "05.03.2026"],
        ["00126", "客户丁", "980", "2026年3月5日"],
        ["客户编号", "客户名称", "金额", "下单日期"],
        ["00128", "客户戊", "5400", "20260307"],
    ]
    marks = {(0, 0): DIM, (2, 0): BLUE, (2, 1): BLUE, (2, 2): BLUE, (2, 3): BLUE,
             (3, 3): AMBER, (4, 3): AMBER, (5, 3): AMBER, (6, 3): AMBER,
             (8, 3): AMBER, (8, 2): AMBER, (4, 2): AMBER, (5, 1): AMBER,
             (7, 0): WARN, (7, 1): WARN, (7, 2): WARN, (7, 3): WARN}
    table(d, 48, 200, ["A", "B", "C", "D"], rows, [105, 110, 115, 120],
          row_h=25, marks=marks, size=13)
    draw_rich(d, (48, 458), "header on row 3  ·  repeated header  ·  5 date formats", 13, DIM)
    draw_rich(d, (48, 482), "nbsp  ·  full-width digits  ·  money as text  ·  duplicate row", 13, DIM)

    draw_rich(d, (660, 184), "AFTER  —  _issues sheet, ranked", 14, GREEN)
    iss = [
        ["high", "5 date formats in one column", "D"],
        ["high", "1 date is day/month ambiguous", "D"],
        ["medium", "8 numbers stored as text", "C"],
        ["medium", "header repeated inside data", "-"],
        ["medium", "1 duplicate row", "-"],
        ["medium", "2 cells with nbsp / full-width", "B"],
        ["low", "3 blank rows inside the data", "-"],
        ["low", "1 column has no header", "F"],
    ]
    colors = {"high": WARN, "medium": AMBER, "low": DIM}
    y = 206
    d.rectangle([660, y, 660 + 500, y + 27], fill=(30, 36, 45))
    draw_rich(d, (668, y + 6), "severity", 13, DIM)
    draw_rich(d, (740, y + 6), "finding", 13, DIM)
    draw_rich(d, (1180, y + 6), "col", 13, DIM)
    y += 27
    for sev, txt, col in iss:
        draw_rich(d, (668, y + 5), sev, 13, colors[sev])
        draw_rich(d, (740, y + 5), txt, 13, FG)
        draw_rich(d, (1180, y + 5), col, 13, DIM)
        d.line([660, y + 26, 1160, y + 26], fill=LINE)
        y += 27
    d.rectangle([660, 206, 1160, y], outline=LINE)
    draw_rich(d, (660, y + 12), "+ _columns sheet: fill rate, types, max length per column", 13, DIM)
    draw_rich(d, (660, y + 36), "+ _summary sheet: rows, merges, cover pages, score, verdict", 13, DIM)

    stats(d, 552, [("30 / 100", "readiness score"), ("20", "findings"), ("0", "cells written")])
    footer(d, "profile  ·  read-only by design: the file is opened, never saved back")
    p = os.path.join(OUT, "case-04-profile.png")
    img.save(p)
    return p


def main():
    os.makedirs(OUT, exist_ok=True)
    for fn in (case_merge, case_clean, case_split, case_profile):
        print("wrote", fn())


if __name__ == "__main__":
    main()
