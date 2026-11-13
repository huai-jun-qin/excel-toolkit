#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
selftest_profile.py - 用真脏表验证 excel_profile.py，逐条断言体检结论。

这个工具的承诺有两条，测试就盯这两条：
    1. 源文件一个字节都不能改（只读体检）。
    2. 报出来的问题必须真的存在，没报的要真的没有。

所以每个用例都造一张「只坏了这一处」的表，再看报告里有没有对应条目 ——
不是看终端有没有打印漂亮的字。

    python3 selftest_profile.py
    PROFILE_SCRIPT=/tmp/old_profile.py python3 selftest_profile.py
"""

import hashlib
import os
import shutil
import subprocess
import sys
import tempfile

from openpyxl import Workbook, load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE = os.environ.get("PROFILE_SCRIPT") or os.path.join(HERE, "excel_profile.py")


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def eq(fails, got, want, label):
    if got != want:
        fails.append("%s: 期望 %r，实际 %r" % (label, want, got))


def sha1(path):
    h = hashlib.sha1()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def run(src, extra=None, out=None):
    """Run the tool on src; return (stdout+stderr, report workbook path)."""
    rep = out or (os.path.splitext(src)[0] + "_profile.xlsx")
    cmd = [sys.executable, PROFILE, "--src", src, "--out", rep] + (extra or [])
    p = subprocess.run(cmd, capture_output=True, text=True)
    return (p.stdout or "") + (p.stderr or ""), rep, p.returncode


def issues_of(report):
    """Read the _issues sheet that the tool wrote."""
    wb = load_workbook(report, data_only=True)
    try:
        ws = wb["_issues"]
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    if not rows:
        return []
    head = [str(c) if c is not None else "" for c in rows[0]]
    out = []
    for r in rows[1:]:
        out.append({head[i]: ("" if v is None else str(v)) for i, v in enumerate(r)})
    return out


def summary_of(report):
    wb = load_workbook(report, data_only=True)
    try:
        ws = wb["_summary"]
        d = {}
        for k, v in ws.iter_rows(values_only=True):
            if k is not None:
                d[str(k)] = "" if v is None else str(v)
    finally:
        wb.close()
    return d


def hits(issues, needle):
    return [i for i in issues if needle in i.get("issue", "")]


def col_hits(issues, column, needle):
    return [i for i in issues if i.get("column") == column and needle in i.get("issue", "")]


# ---------------------------------------------------------------------------
# fixtures: each one breaks exactly one thing
# ---------------------------------------------------------------------------
def b_clean(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["order_id", "amount", "date"])
    ws.append(["A-1001", 120.5, "2026-03-01"])
    ws.append(["A-1002", 88, "2026-03-02"])
    ws.append(["A-1003", 240, "2026-03-03"])
    wb.save(path)


def b_header_offset(path):
    """标题占两行、表头在第 3 行 —— 导出报表最常见的形状。"""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "2026 年 3 月 订单明细"
    ws.merge_cells("A1:C1")
    ws["A2"] = "导出时间: 2026-04-01"
    ws.append([])
    ws.append(["客户", "金额", "日期"])
    ws.append(["甲", 100, "2026-03-01"])
    ws.append(["乙", 200, "2026-03-02"])
    wb.save(path)


def b_numtext(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "金额"])
    ws.append(["甲", "1,280.50"])
    ws.append(["乙", "980"])
    ws.append(["丙", "¥ 1,050 元"])
    ws.append(["丁", "2,300"])
    wb.save(path)


def b_leading_zero(path):
    """编号 00123 必须保持文本，不能被建议转数字。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["客户编号", "客户名称"])
    ws.append(["00123", "甲"])
    ws.append(["00124", "乙"])
    ws.append(["00125", "丙"])
    wb.save(path)


def b_date_spellings(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "下单日期"])
    ws.append(["甲", "2026-03-01"])
    ws.append(["乙", "2026/3/2"])
    ws.append(["丙", "2026年3月3日"])
    ws.append(["丁", "20260304"])
    wb.save(path)


def b_eight_digit_date(path):
    """20260305 是日期不是数字，绝不能报成「数字存成文本」。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "日期"])
    ws.append(["甲", "20260305"])
    ws.append(["乙", "20260306"])
    wb.save(path)


def b_ambiguous(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "日期"])
    ws.append(["甲", "05.03.2026"])
    ws.append(["乙", "06.03.2026"])
    wb.save(path)


def b_dupes(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "金额"])
    ws.append(["甲", 100])
    ws.append(["乙", 200])
    ws.append(["甲", 100])
    ws.append(["丙", 300])
    wb.save(path)


def b_blank_rows(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "金额"])
    ws.append(["甲", 100])
    ws.append([None, None])
    ws.append(["乙", 200])
    ws.append([None, None])
    ws.append(["丙", 300])
    wb.save(path)


def b_merged(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "地区", "金额"])
    ws.append(["甲", "华东", 100])
    ws.append(["乙", None, 200])
    ws.merge_cells("B2:B3")
    wb.save(path)


def b_repeat_header(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "金额"])
    ws.append(["甲", 100])
    ws.append(["乙", 200])
    ws.append(["客户", "金额"])
    ws.append(["丙", 300])
    wb.save(path)


def b_dup_column(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "金额", "金额"])
    ws.append(["甲", 100, 150])
    ws.append(["乙", 200, 250])
    wb.save(path)


def b_dirty_text(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "备注"])
    ws.append(["甲", "  张三  "])
    ws.append(["乙", "赵\u00a0六"])
    ws.append(["丙", "李\u200b四"])
    wb.save(path)


def b_cover(path):
    """第一张是封面，真正数据在第二张。"""
    wb = Workbook()
    cover = wb.active
    cover.title = "封面"
    cover["A1"] = "月度报表"
    cover["A2"] = "内部资料"
    data = wb.create_sheet("明细")
    data.append(["客户", "金额"])
    data.append(["甲", 100])
    data.append(["乙", 200])
    wb.save(path)


# ---------------------------------------------------------------------------
# assertions
# ---------------------------------------------------------------------------
def c_clean(issues, summary, out, code):
    f = []
    eq(f, len(issues), 0, "干净表的问题数")
    eq(f, summary.get("score"), "100 / 100", "得分")
    if "基本干净" not in summary.get("verdict", ""):
        f.append("结论不对：%r" % summary.get("verdict"))
    eq(f, code, 0, "退出码")
    return f


def c_header_offset(issues, summary, out, code):
    f = []
    eq(f, summary.get("score"), "100 / 100", "表头在后两行时仍应判干净")
    eq(f, len(issues), 0, "问题数")
    return f


def c_numtext(issues, summary, out, code):
    f = []
    h = col_hits(issues, "金额", "数字被存成了文本")
    eq(f, len(h), 1, "应报 1 条数字存文本")
    if h and "4" not in h[0].get("issue", ""):
        f.append("条数不对：%r" % h[0].get("issue"))
    return f


def c_leading_zero(issues, summary, out, code):
    f = []
    bad = col_hits(issues, "客户编号", "数字被存成了文本")
    if bad:
        f.append("把带前导零的编号报成了数字存文本：%r" % bad[0].get("issue"))
    ok = col_hits(issues, "客户编号", "编号列存为文本")
    if not ok:
        f.append("没有肯定「编号列存为文本（正确）」")
    return f


def c_date_spellings(issues, summary, out, code):
    f = []
    h = col_hits(issues, "下单日期", "种日期写法")
    eq(f, len(h), 1, "应报日期写法不统一")
    if h and "4" not in h[0].get("issue", ""):
        f.append("写法数量不对：%r" % h[0].get("issue"))
    bad = col_hits(issues, "下单日期", "数字被存成了文本")
    if bad:
        f.append("把 8 位日期报成了数字存文本：%r" % bad[0].get("issue"))
    return f


def c_eight_digit_date(issues, summary, out, code):
    f = []
    bad = [i for i in issues if "数字被存成了文本" in i.get("issue", "")]
    if bad:
        f.append("20260305 被建议转成数字：%r" % bad[0].get("issue"))
    return f


def c_ambiguous(issues, summary, out, code):
    f = []
    h = col_hits(issues, "日期", "日/月无法判定")
    eq(f, len(h), 1, "应报日月歧义")
    return f


def c_dupes(issues, summary, out, code):
    f = []
    h = hits(issues, "完全重复的行")
    eq(f, len(h), 1, "应报重复行")
    if h and "1" not in h[0].get("issue", ""):
        f.append("重复行数不对：%r" % h[0].get("issue"))
    return f


def c_blank_rows(issues, summary, out, code):
    f = []
    h = hits(issues, "空行")
    if not h:
        f.append("没报空行")
    elif "2" not in h[0].get("issue", ""):
        f.append("空行数不对：%r" % h[0].get("issue"))
    return f


def c_merged(issues, summary, out, code):
    f = []
    h = hits(issues, "合并单元格")
    eq(f, len(h), 1, "应报合并单元格")
    return f


def c_repeat_header(issues, summary, out, code):
    f = []
    h = hits(issues, "表头在数据里又出现")
    eq(f, len(h), 1, "应报重复表头")
    return f


def c_dup_column(issues, summary, out, code):
    f = []
    h = hits(issues, "重名列")
    eq(f, len(h), 1, "应报重名列")
    return f


def c_dirty_text(issues, summary, out, code):
    f = []
    h = col_hits(issues, "备注", "多余空格")
    eq(f, len(h), 1, "应报脏文本")
    return f


def c_cover(issues, summary, out, code):
    f = []
    h = hits(issues, "封面")
    if not h:
        f.append("没认出封面页")
    if h and h[0].get("sheet") != "封面":
        f.append("封面判定打在了错的表上：%r" % h[0].get("sheet"))
    return f


CASES = [
    ("干净表得 100 分", b_clean, c_clean, []),
    ("表头在第 3 行仍认得出", b_header_offset, c_header_offset, []),
    ("数字存成文本", b_numtext, c_numtext, []),
    ("前导零编号不误报", b_leading_zero, c_leading_zero, []),
    ("日期写法不统一", b_date_spellings, c_date_spellings, []),
    ("8 位日期不当数字", b_eight_digit_date, c_eight_digit_date, []),
    ("日/月歧义要报出来", b_ambiguous, c_ambiguous, []),
    ("完全重复的行", b_dupes, c_dupes, []),
    ("夹在中间的空行", b_blank_rows, c_blank_rows, []),
    ("合并单元格", b_merged, c_merged, []),
    ("数据里重复表头", b_repeat_header, c_repeat_header, []),
    ("表头重名列", b_dup_column, c_dup_column, []),
    ("多余空格/不可见字符", b_dirty_text, c_dirty_text, []),
    ("封面页识别", b_cover, c_cover, []),
]


def main():
    root = tempfile.mkdtemp(prefix="profile_selftest_")
    fails = 0
    print("selftest: excel_profile.py   (%s)" % PROFILE)
    print("python  : %s %s" % (sys.executable, sys.version.split()[0]))
    print("=" * 76)
    try:
        for i, (name, build, check, extra) in enumerate(CASES, 1):
            d = os.path.join(root, "case%02d" % i)
            os.makedirs(d, exist_ok=True)
            src = os.path.join(d, "in.xlsx")
            build(src)
            before = sha1(src)
            out, rep, code = run(src, extra)
            after = sha1(src)
            if not os.path.exists(rep):
                print("[FAIL] %-24s 没有生成报告" % name)
                print(out[-400:])
                fails += 1
                continue
            problems = check(issues_of(rep), summary_of(rep), out, code)
            if before != after:
                problems.append("源文件被改动了（sha1 变了）")
            ok = not problems
            print("[%s] %-24s %s" % ("PASS" if ok else "FAIL", name,
                                     "" if ok else problems[0]))
            for m in problems[1:]:
                print("       %s" % m)
            if not ok:
                fails += 1

        # --- 只读承诺：另外单独验一次 --------------------------------
        d = os.path.join(root, "readonly")
        os.makedirs(d, exist_ok=True)
        src = os.path.join(d, "in.xlsx")
        b_numtext(src)
        before = (sha1(src), os.path.getmtime(src))
        run(src)
        after = (sha1(src), os.path.getmtime(src))
        ok = before == after
        print("[%s] %-24s %s" % ("PASS" if ok else "FAIL", "源文件只读不改",
                                 "" if ok else "内容或修改时间变了"))
        if not ok:
            fails += 1

        # --- --strict 遇到高危要退 2 --------------------------------
        d = os.path.join(root, "strict")
        os.makedirs(d, exist_ok=True)
        src = os.path.join(d, "in.xlsx")
        b_ambiguous(src)
        _out, _rep, code = run(src, ["--strict"])
        ok = code == 2
        print("[%s] %-24s %s" % ("PASS" if ok else "FAIL", "--strict 高危退 2",
                                 "" if ok else "退出码是 %d 不是 2" % code))
        if not ok:
            fails += 1

        # --- CSV 也要能体检 ----------------------------------------
        d = os.path.join(root, "csv")
        os.makedirs(d, exist_ok=True)
        src = os.path.join(d, "in.csv")
        with open(src, "w", encoding="utf-8", newline="") as fh:
            fh.write("客户,金额\n甲,\"1,280.50\"\n乙,980\n")
        _out, rep, code = run(src)
        ok = code == 0 and os.path.exists(rep)
        detail = ""
        if not ok:
            detail = "CSV 没跑通"
        else:
            h = col_hits(issues_of(rep), "金额", "数字被存成了文本")
            if not h:
                ok, detail = False, "CSV 里的数字存文本没报"
        print("[%s] %-24s %s" % ("PASS" if ok else "FAIL", "CSV 也能体检", detail))
        if not ok:
            fails += 1

        total = len(CASES) + 3
    finally:
        shutil.rmtree(root, ignore_errors=True)

    print("=" * 76)
    print("%d passed, %d failed" % (total - fails, fails))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
