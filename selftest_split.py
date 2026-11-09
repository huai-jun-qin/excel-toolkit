#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
selftest_split.py - 用真脏表验证 excel_split.py，逐文件逐行断言。

重点只有一个：分出去的行加起来必须等于进来的行。少一行就是丢客户数据。

    python3 selftest_split.py
    SPLIT_SCRIPT=/tmp/old_split.py python3 selftest_split.py
"""

import os
import shutil
import subprocess
import sys
import tempfile

from openpyxl import Workbook, load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
SPLIT = os.environ.get("SPLIT_SCRIPT") or os.path.join(HERE, "excel_split.py")


# ---------------------------------------------------------------------------
def b_merged(path):
    """地区合并 B2:B4：只有第一行存着值，另两行是空的。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "明细"
    ws.append(["客户", "地区", "金额"])
    ws.append(["甲", "华东", 100])
    ws.append(["乙", None, 200])
    ws.append(["丙", None, 300])
    ws.merge_cells("B2:B4")
    wb.save(path)


def b_blank_key(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "地区"])
    ws.append(["甲", "华东"])
    ws.append(["乙", ""])
    ws.append(["丙", "华南"])
    ws.append(["丁", "华南"])
    wb.save(path)


def b_illegal(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "地区"])
    ws.append(["甲", "华东/上海"])
    ws.append(["乙", "华东/上海"])
    wb.save(path)


def b_chunk(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["序号", "金额"])
    for i in range(10):
        ws.append([i + 1, 100 + i])
    wb.save(path)


def b_repeat_header(path):
    """数据中间又打了一遍表头，是导出工具常见行为。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "地区"])
    ws.append(["甲", "华东"])
    ws.append(["客户", "地区"])
    ws.append(["乙", "华南"])
    wb.save(path)


def b_cover(path):
    wb = Workbook()
    c = wb.active
    c.title = "封面"
    c["A1"] = "2026年3月对账单"
    c["A3"] = "制表：李四"
    c["A4"] = "本表数据以实际到账为准，请勿外传"
    d = wb.create_sheet("明细")
    d.append(["客户", "地区"])
    d.append(["甲", "华东"])
    d.append(["乙", "华南"])
    wb.save(path)


def b_number_key(path):
    """35000 和 35000.0 在 Excel 里是同一个数。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "编号"])
    ws.append(["甲", 35000])
    ws.append(["乙", 35000.0])
    wb.save(path)


# ---------------------------------------------------------------------------
def read_dir(d):
    out = {}
    for f in sorted(os.listdir(d)):
        if not f.endswith(".xlsx"):
            continue
        ws = load_workbook(os.path.join(d, f)).active
        out[f] = [list(r) for r in ws.iter_rows(values_only=True)]
    return out


def eq(fails, got, want, label):
    if got != want or type(got) is not type(want):
        fails.append("%s: 得到 %r，应为 %r" % (label, got, want))


# ---------------------------------------------------------------------------
def c_merged(files, out, body_rows):
    f = []
    eq(f, len(files), 1, "文件数")
    rows = list(files.values())[0]
    eq(f, len(rows) - 1, 3, "数据行数")
    if len(rows) >= 4:
        eq(f, rows[1][1], "华东", "第1条地区")
        eq(f, rows[2][1], "华东", "第2条地区")
        eq(f, rows[3][1], "华东", "第3条地区")
    return f


def c_blank(files, out, body_rows):
    f = []
    total = sum(len(v) - 1 for v in files.values())
    eq(f, total, 3, "数据行数合计")
    if "Skipped 1 row" not in out:
        f.append("终端没有说明跳过了 1 行空 key")
    return f


def c_illegal(files, out, body_rows):
    f = []
    eq(f, len(files), 1, "文件数")
    name = list(files)[0] if files else ""
    if "华东_上海" not in name:
        f.append("文件名没有把 / 换成 _：%r" % name)
    return f


def c_chunk(files, out, body_rows):
    f = []
    eq(f, len(files), 4, "文件数 (10 行 / 每块 3)")
    eq(f, sum(len(v) - 1 for v in files.values()), 10, "数据行数合计")
    return f


def c_repeat(files, out, body_rows):
    f = []
    eq(f, len(files), 2, "文件数")
    bad = [n for n in files if "地区" in n]
    if bad:
        f.append("重复表头被当成一个分组，生成了 %r" % bad[0])
    eq(f, sum(len(v) - 1 for v in files.values()), 2, "数据行数合计")
    return f


def c_cover(files, out, body_rows):
    f = []
    eq(f, len(files), 2, "文件数")
    for name, rows in files.items():
        if rows and rows[0] and rows[0][0] == "2026年3月对账单":
            f.append("拆的是封面页：%r" % name)
    return f


def c_number(files, out, body_rows):
    f = []
    eq(f, len(files), 1, "文件数 (35000 与 35000.0 应同组)")
    eq(f, sum(len(v) - 1 for v in files.values()), 2, "数据行数合计")
    return f


CASES = [
    ("合并单元格不能丢行", b_merged, ["--by", "地区"], c_merged),
    ("空 key 跳过要说明", b_blank_key, ["--by", "地区"], c_blank),
    ("非法文件名自动替换", b_illegal, ["--by", "地区"], c_illegal),
    ("按每 3 行分块", b_chunk, ["--chunk", "3"], c_chunk),
    ("重复表头不该成为一个分组", b_repeat_header, ["--by", "地区"], c_repeat),
    ("封面页不参与拆分", b_cover, ["--by", "地区"], c_cover),
    ("35000 与 35000.0 同一组", b_number_key, ["--by", "编号"], c_number),
]


def main():
    root = tempfile.mkdtemp(prefix="split_selftest_")
    fails = 0
    print("selftest: excel_split.py   (%s)" % SPLIT)
    print("python  : %s %s" % (sys.executable, sys.version.split()[0]))
    print("=" * 76)
    try:
        for i, (name, build, args, check) in enumerate(CASES, 1):
            d = os.path.join(root, "case%02d" % i)
            os.makedirs(d, exist_ok=True)
            src = os.path.join(d, "in.xlsx")
            outdir = os.path.join(d, "out")
            build(src)
            p = subprocess.run(
                [sys.executable, SPLIT, "--src", src, "--out", outdir] + args,
                capture_output=True, text=True)
            if p.returncode != 0:
                print("[FAIL] %-26s split 退出码 %d" % (name, p.returncode))
                print((p.stdout or "")[-500:])
                print((p.stderr or "")[-500:])
                fails += 1
                continue
            files = read_dir(outdir)
            problems = check(files, (p.stdout or "") + (p.stderr or ""), None)
            ok = not problems
            print("[%s] %-26s %s" % ("PASS" if ok else "FAIL", name,
                                     "" if ok else problems[0]))
            for m in problems[1:]:
                print("       %s" % m)
            if not ok:
                fails += 1
    finally:
        shutil.rmtree(root, ignore_errors=True)

    print("=" * 76)
    print("%d passed, %d failed" % (len(CASES) - fails, fails))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
