# -*- coding: utf-8 -*-
"""Apply V17 Phase-0 / text-only edits in Microsoft Word with Track Changes.

Do NOT rewrite the docx with python-docx. All edits go through Word COM
so Paperpile fields, comments, and run-level formatting stay intact.
"""
from __future__ import print_function

import os
import sys
import time
from pathlib import Path

import win32com.client as win32
from win32com.client import constants

sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import V17
DOC_PATH = str(V17)

WD_FIND_CONTINUE = 1
WD_REPLACE_ONE = 1
WD_REPLACE_ALL = 2
WD_ALERTS_NONE = 0


def log(msg):
    print(msg, flush=True)


def find_execute(doc, old, replace_mode=WD_REPLACE_ONE, new=None, match_case=True):
    rng = doc.Content
    f = rng.Find
    f.ClearFormatting()
    f.Replacement.ClearFormatting()
    f.Text = old
    if new is not None:
        f.Replacement.Text = new
    f.Forward = True
    f.Wrap = WD_FIND_CONTINUE
    f.Format = False
    f.MatchCase = match_case
    f.MatchWholeWord = False
    f.MatchWildcards = False
    f.MatchSoundsLike = False
    f.MatchAllWordForms = False
    if new is None:
        ok = bool(f.Execute())
        return ok, rng if ok else None
    ok = bool(f.Execute(Replace=replace_mode))
    return ok, rng if ok else None


def replace_one(doc, old, new, label):
    ok, _ = find_execute(doc, old, WD_REPLACE_ONE, new, match_case=True)
    log("  [{}] {}  {}".format("OK" if ok else "MISS", label, "found" if ok else "NOT FOUND"))
    if not ok:
        log("      FIND: " + old[:120].replace("\n", " "))
    return ok


def replace_all(doc, old, new, label):
    # Count by looping ReplaceOne to know how many hits.
    n = 0
    while True:
        ok, _ = find_execute(doc, old, WD_REPLACE_ONE, new, match_case=True)
        if not ok:
            break
        n += 1
        if n > 50:
            log("  [STOP] {} hit safety cap".format(label))
            break
    log("  [{}] {}  n={}".format("OK" if n else "MISS", label, n))
    return n


def add_comment(doc, find_text, comment, label):
    ok, rng = find_execute(doc, find_text, new=None)
    if not ok:
        log("  [MISS] comment {} — anchor not found".format(label))
        return False
    doc.Comments.Add(rng, comment)
    log("  [OK] comment {}".format(label))
    return True


def main():
    if not os.path.isfile(DOC_PATH):
        log("Missing " + DOC_PATH)
        sys.exit(1)

    import pythoncom
    pythoncom.CoInitialize()
    log("Opening Word: " + DOC_PATH)
    word = None
    own_word = False
    try:
        word = win32.GetActiveObject("Word.Application")
        log("Attached to running Word instance")
    except Exception as e:
        log("GetActiveObject failed: {}".format(e))
        try:
            word = win32.Dispatch("Word.Application")
            own_word = True
            log("Dispatch Word.Application")
        except Exception as e2:
            log("Dispatch failed: {}".format(e2))
            word = win32.gencache.EnsureDispatch("Word.Application")
            own_word = True
            log("EnsureDispatch Word.Application")
    word.Visible = True
    word.DisplayAlerts = WD_ALERTS_NONE
    try:
        word.Options.CheckGrammarAsYouType = False
        word.Options.CheckSpellingAsYouType = False
    except Exception:
        pass

    doc = None
    n_ok = 0
    n_miss = 0
    try:
        doc = word.Documents.Open(DOC_PATH, ReadOnly=False, AddToRecentFiles=False)
        doc.TrackRevisions = True
        try:
            doc.ShowRevisions = True
        except Exception:
            pass
        log("TrackRevisions = {}".format(doc.TrackRevisions))

        log("\n=== Phase 0 backfills ===")
        items = [
            (
                "P0-abstract-survival",
                "Vaccine sq4 reduced tumor volumes by 50% versus PBS and significantly prolonged median survival.",
                "Vaccine sq4 reduced tumor volumes by approximately 50% versus PBS controls at Day 25 and significantly prolonged median survival (52 vs. 21 days; HR = 0.35, P = 0.003).",
            ),
            (
                "P0-abstract-massive",
                "induced massive tumor-infiltrating CD8+ T-cell expansion",
                "induced tumor-infiltrating CD8+ T-cell expansion (CD8+ among CD3+ lymphocytes increased from 3.6–3.9% to as high as 58.5%)",
            ),
            (
                "P0-CD86",
                "showed dose-dependent CD86 upregulation, confirming",
                "showed dose-dependent CD86 upregulation (23.8–33.2% CD86+ versus 2.64% in controls), confirming",
            ),
            (
                "P0-in-vivo-vol",
                "as reflected by both tumor volume and weight (Figure 6g). Tumor weights were reduced by about 50% compared with controls.",
                "as reflected by both tumor volume and weight, with tumor weights reduced by approximately 50% compared with controls (Figure 6g).",
            ),
            (
                "P0-in-vivo-surv",
                "Survival analysis further showed significantly prolonged median survival in vaccinated groups, with sq4 achieving the most pronounced benefit (Figure 6h).",
                "Kaplan–Meier analysis demonstrated significantly prolonged median survival: 52 days versus 21 days for PBS controls (HR = 0.35, P = 0.003; n = 8 per group, log-rank test), a 2.5-fold extension, with sq4 achieving the most pronounced benefit (Figure 6h).",
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
        ]

        log("\n=== Phase 4 language ===")
        items += [
            (
                "P4-IFN-gamma-abstract",
                "validated by IFN- ELISpot",
                "validated by IFN-\u03b3 ELISpot",
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
        ]

        log("\n=== Phase 3 citation / COI ===")
        items += [
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
                "",
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
                "peptide length 8–14 amino acids, PSM-level FDR 1%",
                "peptide length 8–14 amino acids, a global PSM-level FDR of 1% across the combined canonical and non-canonical LRDB",
            ),
        ]

        for label, old, new in items:
            if replace_one(doc, old, new, label):
                n_ok += 1
            else:
                n_miss += 1

        log("\n=== IC50 subscript (replace all) ===")
        n_ic = replace_all(doc, "IC50", "IC\u2085\u2080", "P4-IC50")
        if n_ic:
            n_ok += 1
        else:
            n_miss += 1

        log("\n=== Pending comments (do not invent missing data) ===")
        comments = [
            (
                "Pathway enrichment analysis against MSigDB Hallmark gene sets",
                "Phase 1 pending: keep Hallmark / FDR < 0.25 until Figure 2d panel is finalized; then align legend, body, and pathway names with the chosen panel (scheme A = v3).",
                "Fig2d-pending",
            ),
            (
                "MaxQuant (v2.5.2)",
                "Phase 1: unify MaxQuant version. Methods = v2.5.2; Figure 4 legend = v2.1. Confirm which version was actually used before changing the other occurrence.",
                "MaxQuant-version",
            ),
            (
                "5 \u00d7 10^7 cells per line",
                "Phase 1: unify cell input. Methods = 1 \u00d7 10^8 total; legend = 5 \u00d7 10^7 per line. State whether the number is total or per line after checking the lab record.",
                "cell-input",
            ),
            (
                "16 (25.0%) elicited robust responses",
                "Phase 2 / M1: confirm whether Donor1/2/3 are three donors or three assays; provide anti-CD3 positive-control values; freeze a single responder definition; then update 16/25.0% if the denominator changes.",
                "ELISpot-M1",
            ),
            (
                "OR = 3.47, 95% CI: 3.28–3.67",
                "Phase 2 / M2: subgroup CI is 6.6-fold narrower than the overall CI and must be recomputed from the 2x2 tables. If the interval cannot be supported, drop the subgroup CI and report the OR with n only.",
                "Fisher-M2",
            ),
            (
                "empty LNP, and PBS groups",
                "Phase 2 / M6: Results compare efficacy vs PBS only. Either add empty-LNP tumor/survival/CD8 panels, or delete this arm from Methods and state in Discussion that adjuvant effects were not separated.",
                "empty-LNP-M6",
            ),
            (
                "encoding tandem minigenes with 4, 8, 8, and 12 neoepitopes",
                "Phase 1 / M3f: do not change 12 to 20 until a sq1–sq4 composition table exists. Figure 5 legend currently says 4 to 20 — keep body at 12 until counted from the table.",
                "sq4-peptide-n",
            ),
        ]
        for find_text, comment, label in comments:
            add_comment(doc, find_text, comment, label)

        # Accepting revisions / deleting comments is Phase 4-4 (submission only). Do not do it now.
        doc.Save()
        log("\nSaved. replacements_ok={} miss={}".format(n_ok, n_miss))
        time.sleep(1)
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=True)
            except Exception as e:
                log("Close warning: {}".format(e))
        if own_word:
            try:
                word.Quit()
            except Exception:
                pass
            log("Word closed.")
        else:
            log("Left existing Word instance running.")

    if n_miss:
        log("WARNING: {} replacements missed. Inspect log above.".format(n_miss))
        sys.exit(2)


if __name__ == "__main__":
    main()
