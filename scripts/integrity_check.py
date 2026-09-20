# -*- coding: utf-8 -*-
"""Compare V17.docx against V17_base.docx for comments, Paperpile fields, and revisions."""
from __future__ import print_function

import os
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import V17, V17_BASE, SUPP_V17, SUPP_ORIGINAL, require

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def stats(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        root = etree.fromstring(z.read("word/document.xml"))
        settings = etree.fromstring(z.read("word/settings.xml"))
        n_comments = 0
        c41 = False
        if "word/comments.xml" in names:
            croot = etree.fromstring(z.read("word/comments.xml"))
            comments = list(croot.iter(W + "comment"))
            n_comments = len(comments)
            for c in comments:
                if c.get(W + "id") == "41":
                    txt = "".join(c.itertext())
                    c41 = "Figure 1d" in txt or "ppt" in txt
        track = settings.find(W + "trackRevisions") is not None
        n_ins = len(root.findall(".//" + W + "ins"))
        n_del = len(root.findall(".//" + W + "del"))
        n_cref = len(root.findall(".//" + W + "commentReference"))
        n_instr = len(root.findall(".//" + W + "instrText"))
        n_fld = len(root.findall(".//" + W + "fldChar"))
        n_b = sum(
            1
            for rpr in root.findall(".//" + W + "rPr")
            if rpr.find(W + "b") is not None
        )
        texts = []
        for t in root.iter(W + "t"):
            if t.text:
                texts.append(t.text)
        joined = "".join(texts)
        nums = [int(x) for x in re.findall(r"\((\d+)\)", joined)]
        return dict(
            ins=n_ins,
            delete=n_del,
            commentRef=n_cref,
            comments=n_comments,
            c41=c41,
            instr=n_instr,
            fld=n_fld,
            bold=n_b,
            track=track,
            size=os.path.getsize(path),
            n_cite_marks=len(nums),
            first20=nums[:20],
        )


def list_sheets(xlsx):
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(xlsx) as z:
        wb = etree.fromstring(z.read("xl/workbook.xml"))
        return [s.get("name") for s in wb.findall(".//" + ns + "sheet")]


def main():
    doc = require(V17, "V17.docx")
    base = require(V17_BASE, "V17_base.docx")
    print("BASE", stats(base))
    print("V17 ", stats(doc))

    xlsx = SUPP_V17 if SUPP_V17.is_file() else SUPP_ORIGINAL
    if xlsx.is_file():
        print("SHEETS", xlsx.name, list_sheets(xlsx))
    else:
        print("SHEETS (missing workbook)")


if __name__ == "__main__":
    main()
