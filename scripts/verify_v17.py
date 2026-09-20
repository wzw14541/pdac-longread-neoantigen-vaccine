# -*- coding: utf-8 -*-
"""Check that V17.docx contains the intended phrases and not the rejected ones."""
from __future__ import print_function

import sys
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import V17, require

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
qn = lambda t: "{%s}%s" % (W, t)

CHECKS = [
    "52 vs. 21 days",
    "23.8",
    "HLA-A*24:02, A*02:01",
    "Pathway enrichment analysis against MSigDB",
    "GSVA",
    "experimental validate",
    "experimentally validate",
    "boasts unparalleled",
    "well-characterized clinical safety profile",
    "become tumor-specific",
    "up to 30% relative to matched normal tissues",
    "In hepatocellular carcinoma, splicing-derived",
    "not directly comparable to the approximately 50%",
    "63/869,610",
    "Chinese (ZCH)",
    "and 20 neoepitopes",
    "and 12 neoepitopes",
    "MYC targets, apoptosis, and angiogenesis",
    "Xenobiotic metabolism",
    "(i) long-read",
    "(ii) novel isoform-derived",
    "(iv) tumor-specific",
    "IC\u2085\u2080",
    "IC50",
    "IFN-\u03b3 ELISpot",
    "IFN- ELISpot",
    "10\u2078",
    "(GGGGS)\u2083",
    "empty LNP",
]


def load_text(path):
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    paras = list(root.iter(qn("p")))
    return paras, "".join(
        "".join((t.text or "") for t in p.iter(qn("t"))) for p in paras
    )


def main():
    doc = require(V17, "V17.docx")
    paras, full = load_text(doc)

    print("=== IFN ords in early paras ===")
    for pi, p in enumerate(paras[:15]):
        joined = "".join((t.text or "") for t in p.iter(qn("t")))
        if "IFN" in joined:
            i = joined.find("IFN")
            chunk = joined[i:i + 20]
            print(pi, repr(chunk), [hex(ord(c)) for c in chunk])

    print("=== verify key phrases ===")
    for s in CHECKS:
        print(("Y" if s in full else "N"), repr(s[:60]), full.count(s))


if __name__ == "__main__":
    main()
