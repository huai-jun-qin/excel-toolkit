#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_gig_steps.py - 补 Gig 页面的第 4、5 张图：流程 与 交付物。

沿用 intro 视频和案例卡片的深色风格，1280x720。
"""

import os
import re

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.expanduser("~/Desktop/gig-images")
MONO = "/System/Library/Fonts/Menlo.ttc"
CJK = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"

W, H = 1280, 720
BG = (13, 17, 23)
CARD = (22, 27, 34)
LINE = (48, 54, 61)
FG = (230, 237, 243)
DIM = (139, 148, 158)
GREEN = (63, 185, 80)
BLUE = (88, 166, 255)
AMBER = (210, 168, 60)

CJK_CLASS = r"\u3000-\u303f\u4e00-\u9fff\uff00-\uffef"
RUN_RE = re.compile("[" + CJK_CLASS + "]+|[^" + CJK_CLASS + "]+")


def rich(d, xy, s, size, fill):
    x, y = xy
    for run in [m.group(0) for m in RUN_RE.finditer(s)]:
        f = ImageFont.truetype(CJK if re.match("[" + CJK_CLASS + "]", run) else MONO, size)
        d.text((x, y), run, font=f, fill=fill)
        x += d.textlength(run, font=f)
    return x


def chrome(d, title):
    d.rectangle([0, 0, W, 44], fill=(10, 14, 19))
    for i, c in enumerate([(255, 95, 86), (255, 189, 46), (39, 201, 63)]):
        d.ellipse([20 + i * 22, 16, 32 + i * 22, 28], fill=c)
    rich(d, (110, 13), title, 14, DIM)


def shot(name):
    os.makedirs(OUT, exist_ok=True)
    img = Image.new("RGB", (W, H), BG)
    return img, ImageDraw.Draw(img), os.path.join(OUT, name)


# ------------------------------------------------------------ 4: how it works
def steps():
    img, d, path = shot("gig-step-4-how.png")
    chrome(d, "how it works  /  3 steps, no software to install")
    rich(d, (48, 74), "Send files. Get one clean workbook.", 36, FG)
    rich(d, (48, 126), "You do nothing technical. No add-in, no login, nothing uploaded.", 16, DIM)

    items = [
        ("1", "You send the folder", "A zip, a Drive link, or\n20 attachments. Messy\nis fine - that is the point."),
        ("2", "I run it and check it", "Row counts must reconcile.\nI send you the count\nbefore the file."),
        ("3", "You get the workbook", "One sheet, plus a tab\nshowing every change and\nwhere each row came from."),
    ]
    x = 56
    for num, head, body in items:
        d.rounded_rectangle([x, 210, x + 372, 486], radius=10, fill=CARD, outline=LINE)
        d.ellipse([x + 28, 236, x + 76, 284], outline=GREEN, width=2)
        rich(d, (x + 44, 248), num, 22, GREEN)
        rich(d, (x + 28, 304), head, 22, FG)
        yy = 344
        for line in body.split("\n"):
            rich(d, (x + 28, yy), line, 15, DIM)
            yy += 26
        x += 404

    rich(d, (48, 528), "Typical turnaround: same day for one folder, 24h for a batch.", 16, DIM)
    d.line([48, 600, W - 48, 600], fill=LINE)
    rich(d, (48, 618), "Before you pay: send one sample file, I confirm it is doable for the price.",
         16, AMBER)
    rich(d, (48, 664), "fiverr.com/claire999j", 15, BLUE)
    img.save(path)
    return path


# ------------------------------------------------------------ 5: what you get
def delivery():
    img, d, path = shot("gig-step-5-delivery.png")
    chrome(d, "what you get  /  every number can be checked by you")
    rich(d, (48, 74), "Four files, and a reason for each one.", 36, FG)
    rich(d, (48, 126), "The change log is the part clients do not expect. It is also the reason they come back.",
         16, DIM)

    tree = [
        ("merged.xlsx", FG, 0),
        ("Merged", GREEN, 1),
        ("_report    rows read / blank / written, per file", DIM, 2),
        ("_aliases   every header that was matched", DIM, 2),
        ("cleaned.xlsx", FG, 0),
        ("Cleaned", GREEN, 1),
        ("_log       cell, value before, value after", DIM, 2),
        ("_summary   counts: dates fixed, numbers fixed, rows dropped", DIM, 2),
        ("delivery-notes.pdf", FG, 0),
        ("what I changed / what I did NOT change / how to verify", DIM, 1),
        ("README-zh.pdf", FG, 0),
        ("same notes in Chinese, for the colleague who has to use it", DIM, 1),
    ]
    y = 200
    d.rounded_rectangle([48, 182, 760, 596], radius=10, fill=CARD, outline=LINE)
    for text, colour, indent in tree:
        if indent == 0:
            rich(d, (72, y), "▸ " + text, 17, colour)
        elif indent == 1 and not text.startswith("Merged") and not text.startswith("Cleaned"):
            rich(d, (96, y), text, 15, DIM)
        else:
            rich(d, (96, y), text, 15, colour if indent == 1 else DIM)
        y += 26 if indent < 2 else 22

    rich(d, (72, 520), "Files are named after your data,", 14, DIM)
    rich(d, (72, 542), "never 'Sheet1' or 'copy of copy'.", 14, DIM)
    d.rounded_rectangle([800, 182, 1232, 596], radius=10, fill=CARD, outline=LINE)
    rich(d, (828, 202), "The two questions every", 17, DIM)
    rich(d, (828, 226), "client asks:", 17, DIM)
    rich(d, (828, 268), "Did you lose any rows?", 18, FG)
    rich(d, (828, 296), "No. Rows read = blank +", 15, GREEN)
    rich(d, (828, 318), "repeated + written. The", 15, GREEN)
    rich(d, (828, 340), "sheet says so itself.", 15, GREEN)
    rich(d, (828, 386), "Did you invent a value?", 18, FG)
    rich(d, (828, 414), "No. Empty cells stay empty.", 15, GREEN)
    rich(d, (828, 436), "Every change has a before", 15, GREEN)
    rich(d, (828, 458), "and an after on _log.", 15, GREEN)
    rich(d, (828, 512), "Last delivery:", 15, DIM)
    rich(d, (828, 536), "157 rows in, 157 out", 15, FG)
    rich(d, (828, 558), "402 changes logged", 15, FG)

    d.line([48, 630, W - 48, 630], fill=LINE)
    rich(d, (48, 648), "If a number does not reconcile, that is my problem, not yours - I fix it free.",
         16, AMBER)
    rich(d, (48, 682), "fiverr.com/claire999j", 15, BLUE)
    img.save(path)
    return path


if __name__ == "__main__":
    for fn in (steps, delivery):
        print("wrote", fn())
