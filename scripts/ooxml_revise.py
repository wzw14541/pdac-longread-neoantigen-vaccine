# -*- coding: utf-8 -*-
"""Surgical OOXML revisions for V17.docx.

Does not rebuild paragraphs. Only splits/wraps the runs that contain a
matched plain-text span, and writes w:del / w:ins so Word shows Track Changes.
Paperpile fields (fldChar/instrText) and comment anchors are not rewritten.
"""
from __future__ import print_function

import copy
import io
import os
import shutil
import zipfile
from datetime import datetime, timezone

try:
    from lxml import etree
except ImportError:
    raise SystemExit("lxml is required")

from pathlib import Path as _Path
import sys as _sys
_sys.path.insert(0, str(_Path(__file__).resolve().parent))

from paths import V17 as _V17, LOGS as _LOGS, ROOT as _ROOT

DIR = str(_ROOT)
DOC = str(_V17)
LOG = str(_LOGS / "revise_log.txt")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
NSMAP_W = {"w": W}

AUTHOR = "V17 revision"
DATE = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

REPLACEMENTS = [
    # Phase 0
    (
        "P0-abstract-survival",
        "Vaccine sq4 reduced tumor volumes by 50% versus PBS and significantly prolonged median survival.",
        "Vaccine sq4 reduced tumor volumes by approximately 50% versus PBS controls at Day 25 and significantly prolonged median survival (52 vs. 21 days; HR = 0.35, P = 0.003).",
    ),
    (
        "P0-abstract-massive",
        "induced massive tumor-infiltrating CD8+ T-cell expansion",
        "induced tumor-infiltrating CD8+ T-cell expansion (CD8+ among CD3+ lymphocytes increased from 3.6\u20133.9% to as high as 58.5%)",
    ),
    (
        "P0-abstract-massive-alt",
        "induced massive tumor-infiltrating CD8-cell expansion",
        "induced tumor-infiltrating CD8+ T-cell expansion (CD8+ among CD3+ lymphocytes increased from 3.6\u20133.9% to as high as 58.5%)",
    ),
    (
        "P0-CD86",
        "showed dose-dependent CD86 upregulation, confirming",
        "showed dose-dependent CD86 upregulation (23.8\u201333.2% CD86+ versus 2.64% in controls), confirming",
    ),
    (
        "P0-in-vivo-vol",
        "as reflected by both tumor volume and weight (Figure 6g). Tumor weights were reduced by about 50% compared with controls.",
        "as reflected by both tumor volume and weight, with tumor weights reduced by approximately 50% compared with controls (Figure 6g).",
    ),
    (
        "P0-in-vivo-surv",
        "Survival analysis further showed significantly prolonged median survival in vaccinated groups, with sq4 achieving the most pronounced benefit (Figure 6h).",
        "Kaplan\u2013Meier analysis demonstrated significantly prolonged median survival: 52 days versus 21 days for PBS controls (HR = 0.35, P = 0.003; n = 8 per group, log-rank test), a 2.5-fold extension, with sq4 achieving the most pronounced benefit (Figure 6h).",
    ),
    (
        "P0-HLA-8alleles",
        "HLA-A24:02, A02:01, A11:01, B13:02, B40:01, C01:02, C03:04, C07:02",
        "HLA-A*24:02, A*02:01, A*11:01, B*13:02, B*40:01, C*01:02, C*03:04, C*07:02",
    ),
    (
        "P0-Fig4-HLA-binders",
        "high-affinity HLA-A24:02 binders",
        "high-affinity HLA-A*24:02 binders",
    ),
    (
        "P0-Fig4-HLA-presented",
        "three HLA-A24:02-presented neoepitopes",
        "three HLA-A*24:02-presented neoepitopes",
    ),
    (
        "P0-Fig2-GSVA",
        "GSVA pathway enrichment analysis",
        "Pathway enrichment analysis",
    ),
    # Phase 4 language
    (
        "P4-IFN-gamma-abstract",
        "validated by IFN- ELISpot",
        "validated by IFN-\u03b3 ELISpot",
    ),
    (
        "P4-IFN-gamma-abstract-alt",
        "by IFN- ELISpot",
        "by IFN-\u03b3 ELISpot",
    ),
    (
        "P4-experimentally",
        "systematically discover and experimental validate",
        "systematically discover and experimentally validate",
    ),
    (
        "P4-boasts",
        "the mRNA-LNP infrastructure boasts unparalleled safety",
        "the mRNA-LNP infrastructure has a well-characterized clinical safety profile",
    ),
    (
        "P4-double-space",
        "lymphocytes  were also significantly elevated",
        "lymphocytes were also significantly elevated",
    ),
    (
        "P4-GGGGS-sub",
        "(GGGGS)3 linkers",
        "(GGGGS)\u2083 linkers",
    ),
    (
        "P4-1e8-sup",
        "A total of 1 \u00d7 108 pancreatic cancer cells",
        "A total of 1 \u00d7 10\u2078 pancreatic cancer cells",
    ),
    # Phase 3
    (
        "P3-Kahles-30pct",
        "drives over 30% of splicing events to become tumor-specific",
        "can increase the number of alternative splicing events by up to 30% relative to matched normal tissues",
    ),
    (
        "P3-59fold-bg",
        "Transcripts generated from AS occur at frequencies > 59-fold higher than mutational counterparts",
        "In hepatocellular carcinoma, splicing-derived neoantigens have been reported at frequencies >59-fold higher than mutation-derived counterparts",
    ),
    (
        "P3-59fold-disc",
        "transcripts generated from AS events occur at frequencies > 59-fold higher than mutational counterparts in independent immunopeptidome analyses",
        "splicing-derived antigens constitute a substantially larger numerical reservoir than mutational neoantigens in independent immunopeptidome analyses of other tumor types",
    ),
    (
        "P3-Rojas-fold",
        "represents a 2.5- to 3.6-fold enrichment over comparably pre-selected mutation-derived neoantigen panels from PDAC (7% for equivalent binding-affinity and immunogenicity-filtered sets)",
        "is a peptide-level hit rate and is not directly comparable to the approximately 50% patient-level T-cell response rate reported for personalized mRNA vaccination in PDAC",
    ),
    (
        "P3-Rojas-gap",
        ", with a larger gap (5- to 8-fold) compared to unselected candidate pools (3%)",
        ".",
    ),
    (
        "P3-COI",
        "The authors declare no competing interests.",
        "Xiangeng Wang is affiliated with AuroraEpitope Limited. A U.S. provisional patent application (No. 63/869,610) related to the technology described in this manuscript has been filed. The remaining authors declare no competing interests.",
    ),
    (
        "P3-enum-i",
        "Three key findings emerge from our work: (1) long-read sequencing",
        "Three key findings emerge from our work: (i) long-read sequencing",
    ),
    (
        "P3-enum-ii",
        "(2) novel isoform-derived peptides show strong enrichment",
        "(ii) novel isoform-derived peptides show strong enrichment",
    ),
    (
        "P3-enum-iii",
        "(3) mRNA-LNP vaccines encoding validated non-canonical neoepitopes",
        "(iii) mRNA-LNP vaccines encoding validated non-canonical neoepitopes",
    ),
    (
        "P3-enum-mech-i",
        "(1) novel isoforms may generate peptides with enhanced affinity",
        "(i) novel isoforms may generate peptides with enhanced affinity",
    ),
    (
        "P3-enum-mech-ii",
        "(2) novel splice junctions create unique epitope sequences",
        "(ii) novel splice junctions create unique epitope sequences",
    ),
    (
        "P3-enum-mech-iii",
        "(3) reduced central tolerance to AS-derived antigens",
        "(iii) reduced central tolerance to AS-derived antigens",
    ),
    (
        "P3-enum-mech-iv",
        "(4) tumor-specific splicing factors may preferentially upregulate immunogenic isoforms",
        "(iv) tumor-specific splicing factors may preferentially upregulate immunogenic isoforms",
    ),
    (
        "P3-global-FDR",
        "peptide length 8\u201314 amino acids, PSM-level FDR 1%",
        "peptide length 8\u201314 amino acids, a global PSM-level FDR of 1% across the combined canonical and non-canonical LRDB",
    ),
]

IC50_OLD = "IC50"
IC50_NEW = "IC\u2085\u2080"

_logs = []


def log(msg):
    _logs.append(msg)
    print(msg, flush=True)


def qn(tag):
    return "{%s}%s" % (W, tag)


def local(el):
    return el.tag.split("}")[-1] if isinstance(el.tag, str) else ""


def is_field_char(el):
    return local(el) == "fldChar"


def field_type(el):
    return el.get(qn("fldCharType"))


def collect_t_in_run(r):
    return [c for c in r.iter(qn("t"))]


def paragraph_plain_items(p):
    """Yield dicts for visible w:t nodes that are NOT inside a field."""
    field_depth = 0
    items = []
    for el in p.iter():
        if is_field_char(el):
            t = field_type(el)
            if t == "begin":
                field_depth += 1
            elif t == "end" and field_depth:
                field_depth -= 1
            continue
        if local(el) == "t" and field_depth == 0:
            parent_r = el.getparent()
            while parent_r is not None and local(parent_r) != "r":
                parent_r = parent_r.getparent()
            items.append({"t": el, "r": parent_r, "text": el.text or ""})
    return items


def joined_text(items):
    return "".join(it["text"] for it in items)


def set_t_text(t_el, text):
    t_el.text = text
    if text.startswith(" ") or text.endswith(" ") or (text != text.strip()):
        t_el.set(XML_SPACE, "preserve")
    elif XML_SPACE in t_el.attrib and not (text.startswith(" ") or text.endswith(" ")):
        # keep existing preserve if originally set; harmless
        pass


def split_run_t(item, offset):
    """Split item's run so that item.t keeps text[:offset] and a new run holds the rest.

    Returns the new item (right half). If offset is 0 or len, no split.
    """
    t_el = item["t"]
    r = item["r"]
    text = t_el.text or ""
    if offset <= 0 or offset >= len(text):
        return None
    left, right = text[:offset], text[offset:]
    set_t_text(t_el, left)
    new_r = copy.deepcopy(r)
    # deepcopy copies ALL children; keep rPr and a single t with right text.
    # Remove extra t / delText from new_r, set the first t.
    ts = collect_t_in_run(new_r)
    if not ts:
        # create t
        new_t = etree.SubElement(new_r, qn("t"))
        set_t_text(new_t, right)
    else:
        set_t_text(ts[0], right)
        for extra in ts[1:]:
            extra.getparent().remove(extra)
    parent = r.getparent()
    idx = list(parent).index(r)
    parent.insert(idx + 1, new_r)
    new_t = collect_t_in_run(new_r)[0]
    return {"t": new_t, "r": new_r, "text": right}


def next_id(counter):
    counter[0] += 1
    return str(counter[0])


def wrap_runs_with_revision(runs, new_text, id_counter):
    """Replace consecutive sibling runs with del(old runs) + ins(new_text)."""
    if not runs:
        return False
    first_r = runs[0]
    parent = first_r.getparent()
    for r in runs:
        if r.getparent() is not parent:
            log("    SKIP wrap: runs do not share parent")
            return False
    children = list(parent)
    idxs = [children.index(r) for r in runs]
    if idxs != list(range(idxs[0], idxs[0] + len(idxs))):
        log("    SKIP wrap: runs not consecutive (intervening nodes %s)" % idxs)
        return False

    del_id = next_id(id_counter)
    ins_id = next_id(id_counter)
    del_el = etree.Element(qn("del"))
    del_el.set(qn("id"), del_id)
    del_el.set(qn("author"), AUTHOR)
    del_el.set(qn("date"), DATE)
    ins_el = etree.Element(qn("ins"))
    ins_el.set(qn("id"), ins_id)
    ins_el.set(qn("author"), AUTHOR)
    ins_el.set(qn("date"), DATE)

    # Move original runs under del, converting t -> delText
    for r in runs:
        for t in collect_t_in_run(r):
            t.tag = qn("delText")
            if t.text and (t.text.startswith(" ") or t.text.endswith(" ")):
                t.set(XML_SPACE, "preserve")
        del_el.append(r)

    # New run copies rPr from the first deleted run if present
    new_r = etree.Element(qn("r"))
    rpr = del_el[0].find(qn("rPr"))
    if rpr is not None:
        new_r.append(copy.deepcopy(rpr))
    new_t = etree.SubElement(new_r, qn("t"))
    set_t_text(new_t, new_text)
    ins_el.append(new_r)

    parent.insert(idxs[0], del_el)
    parent.insert(idxs[0] + 1, ins_el)
    return True


def replace_once_in_tree(root, old, new, id_counter, label):
    if not old:
        return False
    for p in root.iter(qn("p")):
        items = paragraph_plain_items(p)
        joined = joined_text(items)
        pos = joined.find(old)
        if pos < 0:
            continue
        end = pos + len(old)

        # Map to item indices / offsets
        char_i = 0
        start_item = start_off = end_item = end_off = None
        for i, it in enumerate(items):
            n = len(it["text"])
            if start_item is None and char_i + n > pos:
                start_item, start_off = i, pos - char_i
            if start_item is not None and char_i + n >= end:
                end_item, end_off = i, end - char_i
                break
            char_i += n
        if start_item is None or end_item is None:
            log("  [%s] mapped poorly" % label)
            return False

        # Split end first (so start indices remain valid if same item)
        if end_off < len(items[end_item]["text"]):
            split_run_t(items[end_item], end_off)
            items = paragraph_plain_items(p)
            # recompute mapping after split
            joined = joined_text(items)
            pos = joined.find(old)
            end = pos + len(old)
            char_i = 0
            start_item = start_off = end_item = end_off = None
            for i, it in enumerate(items):
                n = len(it["text"])
                if start_item is None and char_i + n > pos:
                    start_item, start_off = i, pos - char_i
                if start_item is not None and char_i + n >= end:
                    end_item, end_off = i, end - char_i
                    break
                char_i += n

        if start_off > 0:
            split_run_t(items[start_item], start_off)
            items = paragraph_plain_items(p)
            joined = joined_text(items)
            pos = joined.find(old)
            end = pos + len(old)
            char_i = 0
            start_item = start_off = end_item = end_off = None
            for i, it in enumerate(items):
                n = len(it["text"])
                if start_item is None and char_i + n > pos:
                    start_item, start_off = i, pos - char_i
                if start_item is not None and char_i + n >= end:
                    end_item, end_off = i, end - char_i
                    break
                char_i += n

        # After splits, matched items should start at off 0 and end at full length
        matched = items[start_item:end_item + 1]
        # Trim last if end_off == 0 (shouldn't)
        runs = []
        seen = set()
        for it in matched:
            rid = id(it["r"])
            if rid not in seen:
                runs.append(it["r"])
                seen.add(rid)
        ok = wrap_runs_with_revision(runs, new, id_counter)
        log("  [%s] %s  para_match old_len=%d new_len=%d runs=%d" % (
            "OK" if ok else "FAIL", label, len(old), len(new), len(runs)))
        return ok
    log("  [MISS] %s" % label)
    return False


def enable_track_revisions(settings_xml):
    root = etree.fromstring(settings_xml)
    existing = root.find(qn("trackRevisions"))
    if existing is None:
        el = etree.Element(qn("trackRevisions"))
        # insert near the top after w:compat or at end of settings children
        root.append(el)
        log("Enabled w:trackRevisions in settings.xml")
    else:
        log("w:trackRevisions already present")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def count_marks(root):
    n_ins = len(root.findall(".//" + qn("ins")))
    n_del = len(root.findall(".//" + qn("del")))
    n_cref = len(root.findall(".//" + qn("commentReference")))
    n_instr = len(root.findall(".//" + qn("instrText")))
    n_fld = len(root.findall(".//" + qn("fldChar")))
    n_b = 0
    for rpr in root.findall(".//" + qn("rPr")):
        if rpr.find(qn("b")) is not None:
            n_b += 1
    return dict(ins=n_ins, delete=n_del, commentRef=n_cref, instrText=n_instr, fldChar=n_fld, bold_rPr=n_b)


def main():
    log("Editing " + DOC)
    with zipfile.ZipFile(DOC, "r") as z:
        doc_xml = z.read("word/document.xml")
        settings_xml = z.read("word/settings.xml")
        others = {name: z.read(name) for name in z.namelist() if name not in ("word/document.xml", "word/settings.xml")}
        names_order = z.namelist()

    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(doc_xml, parser=parser)
    before = count_marks(root)
    log("BEFORE " + str(before))

    # revision ids: start above any existing
    max_id = 0
    for el in root.xpath(".//*[@w:id]", namespaces=NSMAP_W):
        try:
            max_id = max(max_id, int(el.get(qn("id"))))
        except (TypeError, ValueError):
            pass
    id_counter = [max(max_id, 20000)]

    done = []
    missed = []
    skip_alt = set()
    for label, old, new in REPLACEMENTS:
        if label in skip_alt:
            log("  [SKIP] %s (primary already applied)" % label)
            continue
        ok = replace_once_in_tree(root, old, new, id_counter, label)
        if ok:
            done.append(label)
            if label == "P0-abstract-massive":
                skip_alt.add("P0-abstract-massive-alt")
            if label == "P4-IFN-gamma-abstract":
                skip_alt.add("P4-IFN-gamma-abstract-alt")
        else:
            if label.endswith("-alt"):
                missed.append(label)
            else:
                # alt may still succeed
                if label + "-alt" not in [x[0] for x in REPLACEMENTS]:
                    missed.append(label)

    # IC50 replace-all via the same engine
    n_ic = 0
    while replace_once_in_tree(root, IC50_OLD, IC50_NEW, id_counter, "P4-IC50-%d" % (n_ic + 1)):
        n_ic += 1
        if n_ic > 20:
            break
    log("IC50 replacements: %d" % n_ic)
    if n_ic:
        done.append("P4-IC50 x%d" % n_ic)
    else:
        missed.append("P4-IC50")

    after = count_marks(root)
    log("AFTER  " + str(after))

    new_doc_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    new_settings = enable_track_revisions(settings_xml)

    tmp = DOC + ".tmp"
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for name in names_order:
            if name == "word/document.xml":
                z.writestr(name, new_doc_xml)
            elif name == "word/settings.xml":
                z.writestr(name, new_settings)
            else:
                z.writestr(name, others[name])
    os.replace(tmp, DOC)
    log("Wrote " + DOC)
    log("DONE %d  MISS %d" % (len(done), len(missed)))
    for m in missed:
        log("  missed: " + m)

    _LOGS.mkdir(parents=True, exist_ok=True)
    with io.open(LOG, "w", encoding="utf-8") as f:
        f.write("\n".join(_logs) + "\n")


if __name__ == "__main__":
    main()
