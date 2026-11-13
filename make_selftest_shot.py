"""Render docs/selftest.png from the real output of the four self-tests.

Nothing here is typed by hand: the script runs each selftest, keeps the
[PASS]/[FAIL] lines and the final tally, and draws them on the same dark
background as the other cards. Latin text uses Menlo; CJK runs are drawn with
Hiragino Sans GB because Menlo has no Chinese glyphs.
"""

import os
import re
import subprocess
import sys

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "docs", "selftest.png")

W, H = 1280, 720
BG = (22, 26, 32)
FG = (231, 236, 241)
DIM = (130, 142, 155)
GREEN = (86, 196, 126)
RED = (232, 94, 94)
AMBER = (232, 176, 78)
BLUE = (110, 170, 240)
LINE = (52, 60, 70)

MONO = "/System/Library/Fonts/Menlo.ttc"
CJK = "/System/Library/Fonts/Hiragino Sans GB.ttc"

SUITES = [
    ("selftest_merge.py", "merge"),
    ("selftest_clean.py", "clean"),
    ("selftest_split.py", "split"),
    ("selftest_profile.py", "profile"),
]

_cache = {}


def font(path, size, index=0):
    key = (path, size, index)
    if key not in _cache:
        try:
            _cache[key] = ImageFont.truetype(path, size, index=index)
        except TypeError:
            _cache[key] = ImageFont.truetype(path, size)
    return _cache[key]


def is_cjk(ch):
    o = ord(ch)
    return (0x2E80 <= o <= 0x9FFF) or (0xF900 <= o <= 0xFAFF) or \
           (0xFF00 <= o <= 0xFF60) or (0x3000 <= o <= 0x303F)


def draw_mixed(d, xy, text, size, fill):
    """Draw ASCII in Menlo and CJK runs in Hiragino, advancing manually."""
    x, y = xy
    fm = font(MONO, size)
    fc = font(CJK, size)
    run, cur = [], None
    for ch in text:
        want = is_cjk(ch)
        if cur is None or want == cur:
            run.append(ch)
            cur = want
        else:
            f = fc if cur else fm
            s = "".join(run)
            d.text((x, y), s, font=f, fill=fill)
            x += d.textlength(s, font=f)
            run = [ch]
            cur = want
    if run:
        f = fc if cur else fm
        s = "".join(run)
        d.text((x, y), s, font=f, fill=fill)
        x += d.textlength(s, font=f)
    return x


def run(script):
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    p = subprocess.run([sys.executable, script], cwd=HERE, env=env,
                       stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    return p.stdout.decode("utf-8", "replace")


def parse(text):
    out = []
    for raw in text.splitlines():
        s = raw.rstrip()
        if s.startswith("[PASS]") or s.startswith("[FAIL]"):
            out.append((s.startswith("[PASS]"), s[6:].strip()))
    return out


def measure(text, size):
    fm = font(MONO, size)
    fc = font(CJK, size)
    tot = 0.0
    for ch in text:
        tot += (fc if is_cjk(ch) else fm).getlength(ch)
    return tot


def fit(text, size, width):
    """Truncate until it fits `width` px, with an ellipsis."""
    fm = font(MONO, size)
    fc = font(CJK, size)
    def w(t):
        tot = 0.0
        for ch in t:
            tot += (fc if is_cjk(ch) else fm).getlength(ch) if False else 0
        return tot
    # simpler: measure by drawing logic
    def measure(t):
        tot = 0.0
        for ch in t:
            f = fc if is_cjk(ch) else fm
            tot += f.getlength(ch)
        return tot
    if measure(text) <= width:
        return text
    while text and measure(text + "…") > width:
        text = text[:-1]
    return text + "…"


def main():
    data = []
    total_p = total_f = 0
    for script, name in SUITES:
        if not os.path.exists(os.path.join(HERE, script)):
            print("missing", script)
            return 1
        rows = parse(run(script))
        p = sum(1 for ok, s in rows if ok)
        f = len(rows) - p
        total_p += p
        total_f += f
        data.append((name, rows, p, f))
        print("%-22s %2d passed %d failed" % (script, p, f))

    f_title = font(MONO, 30)
    f_head = font(MONO, 15)
    f_body = font(MONO, 12)
    f_note = font(MONO, 12)

    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)

    d.rectangle([0, 0, W, 32], fill=(30, 35, 42))
    for i, c in enumerate([(236, 96, 92), (232, 176, 78), (86, 196, 126)]):
        d.ellipse([16 + i * 20, 11, 26 + i * 20, 21], fill=c)
    d.text((96, 9), "excel-toolkit  /  four self-tests, asserted cell by cell",
           font=font(MONO, 14), fill=DIM)

    d.text((40, 58), "%d checks, no mocking" % (total_p + total_f), font=f_title, fill=FG)
    d.text((40, 100), "Each case builds a workbook that is wrong in exactly one way, runs the real tool,",
           font=f_note, fill=DIM)
    d.text((40, 120), "and asserts the reported findings, not the console output. Source files are hashed before and after.",
           font=f_note, fill=DIM)

    col_w = 306
    x0 = 32
    y0 = 172
    step = 17
    for idx, (name, rows, p, f) in enumerate(data):
        x = x0 + idx * (col_w + 8)
        y = y0
        d.text((x, y), name.upper(), font=f_head, fill=AMBER)
        d.text((x + col_w - 44, y + 2), str(p), font=f_body, fill=GREEN)
        if f:
            d.text((x + col_w - 28, y + 2), str(f), font=f_body, fill=RED)
        y += 21
        d.line([x, y, x + col_w - 12, y], fill=LINE)
        y += 9
        for ok, label in rows:
            size = 11
            while size > 9 and measure(label, size) > col_w - 22:
                size -= 1
            draw_mixed(d, (x, y), fit(label, size, col_w - 22), size,
                       FG if ok else RED)
            if not ok:
                d.rectangle([x - 6, y - 2, x + col_w - 12, y + 13], outline=RED)
            y += step
        d.line([x, y + 3, x + col_w - 12, y + 3], fill=LINE)
        draw_mixed(d, (x, y + 9), "%d passed, %d failed" % (p, f), 12,
                   GREEN if f == 0 else RED)

    f_small = font(MONO, 11)
    notes = [
        "merge   13 cases · cover pages, repeated headers, merged cells, order ids",
        "clean   12 cases · 3 date spellings, text money, dedupe, leading zeros",
        "split     7 cases · merged keys, empty keys, chunking, bad file names",
        "profile 17 cases · one broken workbook per finding, plus a clean one",
    ]
    d.line([32, 508, W - 32, 508], fill=LINE)
    for i, line in enumerate(notes):
        d.text((32, 522 + i * 17), line, font=f_small, fill=DIM)

    d.line([32, H - 78, W - 32, H - 78], fill=LINE)
    verdict = "%d passed, %d failed" % (total_p, total_f)
    d.text((32, H - 58), verdict, font=font(MONO, 22),
           fill=GREEN if total_f == 0 else RED)
    d.text((W - 32 - 302, H - 52), "github.com/huai-jun-qin/excel-toolkit",
           font=f_note, fill=DIM)
    d.text((W - 32 - 250, H - 34), "no network  ·  no pandas  ·  python 3.9",
           font=f_note, fill=(90, 100, 112))

    img.save(OUT)
    print("wrote", OUT)
    return 0 if total_f == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
