#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
selftest_clean.py - 用真脏表验证 excel_clean.py，逐格断言，不是"跑完没报错就算过"。

每个用例都会现造一个工作簿，跑一次 clean，然后把输出逐格读回来比对。
把 CLEAN_SCRIPT 指向另一个版本即可对照（用来证明这些用例真的能抓出问题）：

    python3 selftest_clean.py
    CLEAN_SCRIPT=/tmp/old_clean.py python3 selftest_clean.py
"""

import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime

from openpyxl import Workbook, load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
CLEAN = os.environ.get("CLEAN_SCRIPT") or os.path.join(HERE, "excel_clean.py")


# ---------------------------------------------------------------------------
# builders
# ---------------------------------------------------------------------------
def b_salary(path):
    """整数金额落在 Excel 序列号区间 (20000-60000) 里。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "工资"
    ws.append(["姓名", "工资", "奖金"])
    ws.append(["甲", 35000, 1200])
    ws.append(["乙", 28000, 980])
    ws.append(["丙", 42000, 5400])
    wb.save(path)


def b_dates(path):
    """同一列三种写法，都该归一成同一种日期。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["订单号", "下单日期"])
    ws.append(["A1", "2026/3/1"])
    ws.append(["A2", "2026-3-2"])
    ws.append(["A3", "2026年3月3日"])
    wb.save(path)


def b_text_numbers(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "金额"])
    ws.append(["甲", "1,234"])
    ws.append(["乙", "¥2,345.50"])
    ws.append(["丙", "3500"])
    wb.save(path)


def b_merged(path):
    """地区 B2:B4 合并，只有左上角有值。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "明细"
    ws.append(["客户", "地区", "金额"])
    ws.append(["甲", "华东", 100])
    ws.append(["乙", None, 200])
    ws.append(["丙", None, 300])
    ws.merge_cells("B2:B4")
    wb.save(path)


def b_format(path):
    """订单号 123 靠 0000000 显示成 0000123。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["订单号", "金额"])
    ws.append([123, 10])
    ws["A2"].number_format = "0000000"
    wb.save(path)


def b_mixed(path):
    """58 个真数字 + 2 个文本数字：文本占比 3%，低于原来的 50% 门槛。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["金额"])
    for i in range(58):
        ws.append([100 + i])
    ws.append(["¥1,234"])
    ws.append(["¥2,345"])
    wb.save(path)


def b_cover(path):
    wb = Workbook()
    cover = wb.active
    cover.title = "封面"
    cover["A1"] = "2026年3月对账单"
    cover["A3"] = "制表：李四"
    cover["A4"] = "本表数据以实际到账为准，请勿外传"
    d = wb.create_sheet("明细")
    d.append(["客户", "金额"])
    d.append(["甲", 100])
    d.append(["乙", 200])
    wb.save(path)


def b_dupes(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "金额"])
    ws.append(["甲", 100])
    ws.append(["甲", 100])
    ws.append(["乙", 200])
    wb.save(path)


def b_ids(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["订单号", "数量"])
    ws.append(["00100", "12"])
    ws.append(["00200", "34"])
    wb.save(path)


def b_ambiguous(path):
    """13.05.2026 证明这批是 DD.MM.YYYY，于是 05.03.2026 必须读成 3 月 5 日。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["日期"])
    ws.append(["13.05.2026"])
    ws.append(["05.03.2026"])
    wb.save(path)


def b_space_money(path):
    """千分位用空格/NBSP 写的金额，是文本但必须转成数字，否则 SUM 少算。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "金额"])
    ws.append(["甲", "1\u00a0234.50"])
    ws.append(["乙", "2 345"])
    ws.append(["丙", "3100"])
    wb.save(path)


def b_ymd_int(path):
    """20260305 存成整数：列名叫日期就转，列名叫编号就原样。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["日期", "编号", "数量"])
    for i in range(3):
        ws.append([20260305 + i, 20260305 + i, 12])
    wb.save(path)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def eq(fails, got, want, label):
    if got != want or type(got) is not type(want):
        fails.append("%s: 得到 %r (%s)，应为 %r (%s)"
                     % (label, got, type(got).__name__, want, type(want).__name__))


def col(ws, idx):
    return [ws.cell(row=r, column=idx).value for r in range(2, ws.max_row + 1)]


# ---------------------------------------------------------------------------
# checks  (返回问题列表，空 = 通过)
# ---------------------------------------------------------------------------
def c_salary(ws, out):
    f = []
    for row, want in ((2, 35000), (3, 28000), (4, 42000)):
        eq(f, ws.cell(row=row, column=2).value, want, "工资 第%d行" % row)
    eq(f, ws.cell(row=2, column=3).value, 1200, "奖金 第2行")
    if "date unified" in out:
        f.append("终端把整数金额报成了 date unified")
    return f


def c_dates(ws, out):
    f = []
    for row, want in ((2, datetime(2026, 3, 1)), (3, datetime(2026, 3, 2)), (4, datetime(2026, 3, 3))):
        eq(f, ws.cell(row=row, column=2).value, want, "日期 第%d行" % row)
    return f


def c_text_numbers(ws, out):
    f = []
    eq(f, ws.cell(row=2, column=2).value, 1234, "金额 第2行")
    eq(f, ws.cell(row=3, column=2).value, 2345.5, "金额 第3行")
    eq(f, ws.cell(row=4, column=2).value, 3500, "金额 第4行")
    return f


def c_merged(ws, out):
    f = []
    for row in (2, 3, 4):
        eq(f, ws.cell(row=row, column=2).value, "华东", "地区 第%d行" % row)
    return f


def c_format(ws, out):
    f = []
    eq(f, ws.cell(row=2, column=1).value, 123, "订单号 值")
    eq(f, ws.cell(row=2, column=1).number_format, "0000000", "订单号 显示格式")
    return f


def c_mixed(ws, out):
    f = []
    vals = col(ws, 1)
    eq(f, len(vals), 60, "行数")
    eq(f, vals[-2], 1234, "文本数字 第59行")
    eq(f, vals[-1], 2345, "文本数字 第60行")
    return f


def c_cover(ws, out):
    f = []
    eq(f, ws.cell(row=1, column=1).value, "客户", "表头 A1")
    eq(f, ws.cell(row=1, column=2).value, "金额", "表头 B1")
    eq(f, ws.max_row, 3, "行数(含表头)")
    return f


def c_dupes(ws, out):
    f = []
    eq(f, ws.max_row, 3, "行数(含表头)")
    eq(f, ws.cell(row=2, column=1).value, "甲", "第一行")
    eq(f, ws.cell(row=3, column=1).value, "乙", "第二行")
    return f


def c_ids(ws, out):
    f = []
    eq(f, ws.cell(row=2, column=1).value, "00100", "订单号 第2行")
    eq(f, ws.cell(row=2, column=2).value, 12, "数量 第2行")
    return f


def c_ambiguous(ws, out):
    f = []
    eq(f, ws.cell(row=2, column=1).value, datetime(2026, 5, 13), "13.05.2026")
    eq(f, ws.cell(row=3, column=1).value, datetime(2026, 3, 5), "05.03.2026")
    return f


def c_space_money(ws, out):
    f = []
    eq(f, ws.cell(row=2, column=2).value, 1234.5, "NBSP 千分位金额")
    eq(f, ws.cell(row=3, column=2).value, 2345, "空格千分位金额")
    eq(f, ws.cell(row=4, column=2).value, 3100, "普通金额")
    return f


def c_ymd_int(ws, out):
    f = []
    eq(f, ws.cell(row=2, column=1).value, datetime(2026, 3, 5), "日期列的 8 位整数")
    eq(f, ws.cell(row=2, column=2).value, 20260305, "编号列必须原样不动")
    return f


CASES = [
    ("整数金额不能被当成日期序列号", b_salary, [], c_salary),
    ("三种日期写法统一", b_dates, [], c_dates),
    ("文本金额转真数字", b_text_numbers, [], c_text_numbers),
    ("合并单元格的值要铺满整块", b_merged, [], c_merged),
    ("订单号 0000000 显示格式保留", b_format, [], c_format),
    ("零星文本数字也要转（不能静默）", b_mixed, [], c_mixed),
    ("封面页跳过，清洗明细页", b_cover, [], c_cover),
    ("完全重复行去掉", b_dupes, ["--dedupe"], c_dupes),
    ("前导零订单号保持文本", b_ids, [], c_ids),
    ("日期歧义按列内证据统一", b_ambiguous, [], c_ambiguous),
    ("空格千分位金额要转数字", b_space_money, [], c_space_money),
    ("8位整数：日期列转/编号列不转", b_ymd_int, [], c_ymd_int),
]


def main():
    root = tempfile.mkdtemp(prefix="clean_selftest_")
    fails = 0
    print("selftest: excel_clean.py   (%s)" % CLEAN)
    print("python  : %s %s" % (sys.executable, sys.version.split()[0]))
    print("=" * 78)
    try:
        for i, (name, build, extra, check) in enumerate(CASES, 1):
            d = os.path.join(root, "case%02d" % i)
            os.makedirs(d, exist_ok=True)
            src = os.path.join(d, "in.xlsx")
            out_path = os.path.join(d, "out.xlsx")
            build(src)
            p = subprocess.run(
                [sys.executable, CLEAN, "--src", src, "--out", out_path] + extra,
                capture_output=True, text=True)
            if p.returncode != 0:
                print("[FAIL] %-30s clean 退出码 %d" % (name, p.returncode))
                print((p.stdout or "")[-600:])
                print((p.stderr or "")[-600:])
                fails += 1
                continue
            wb = load_workbook(out_path)
            ws = wb["Cleaned"]
            problems = check(ws, (p.stdout or "") + (p.stderr or ""))
            ok = not problems
            print("[%s] %-30s %s" % ("PASS" if ok else "FAIL", name,
                                     "" if ok else problems[0]))
            for extra_msg in problems[1:]:
                print("       %s" % extra_msg)
            if not ok:
                fails += 1
    finally:
        shutil.rmtree(root, ignore_errors=True)

    print("=" * 78)
    print("%d passed, %d failed" % (len(CASES) - fails, fails))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
