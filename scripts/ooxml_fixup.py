# -*- coding: utf-8 -*-
"""Fix remaining IFN / enum-i / double-period after the main V17 pass."""
import os
import sys
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from ooxml_revise import (
    replace_once_in_tree, count_marks, qn, log, DOC,
)

def main():
    with zipfile.ZipFile(DOC, "r") as z:
        doc_xml = z.read("word/document.xml")
        settings_xml = z.read("word/settings.xml")
        others = {name: z.read(name) for name in z.namelist()
                  if name not in ("word/document.xml", "word/settings.xml")}
        names_order = z.namelist()

    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(doc_xml, parser=parser)
    log("BEFORE " + str(count_marks(root)))

    max_id = 0
    for el in root.xpath(".//*[@w:id]", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}):
        try:
            max_id = max(max_id, int(el.get(qn("id"))))
        except (TypeError, ValueError):
            pass
    id_counter = [max_id]

    extra = [
        ("P4-IFN-split", "validated by IFN-", "validated by IFN-\u03b3"),
        ("P3-enum-i-narrow", "from our work: (1) long-read sequencing", "from our work: (i) long-read sequencing"),
        ("P3-Rojas-dperiod", ".. This enrichment", ". This enrichment"),
    ]
    for label, old, new in extra:
        replace_once_in_tree(root, old, new, id_counter, label)

    log("AFTER " + str(count_marks(root)))
    new_doc_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    tmp = DOC + ".tmp"
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for name in names_order:
            if name == "word/document.xml":
                z.writestr(name, new_doc_xml)
            elif name == "word/settings.xml":
                z.writestr(name, settings_xml)
            else:
                z.writestr(name, others[name])
    os.replace(tmp, DOC)
    log("Wrote " + DOC)


if __name__ == "__main__":
    main()
