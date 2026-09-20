# PDAC V17 revision log

Started from the original `2026-08-20_pdac_V16.docx`, not from `pdac_V16_修改版.docx`.

## Files

- `manuscript/V17_base.docx`: read-only original. Do not edit. Do not click Paperpile Update.
- `manuscript/V17.docx`: working copy with Track Changes (34 insertions / 34 deletions).
- `tables/Supplementary_Tables_V17.xlsx`: current supplementary workbook (Table S2 rebuilt as DEAS).
- `tables/Table_S2.xlsx`: standalone Table S2 (Differential_AS + Event_PSI_JP).
- `figures/2026-08-20_PDAC_mainfigure_v3.pptx` / `v4.pptx`: figure copies; panels not yet locked.

Open `V17.docx` in Word and check the Review pane. Do not refresh Paperpile.

## Integrity vs V17_base

- commentReference 65 = 65
- comments.xml 65 = 65
- Wang teacher comment id 41 (Figure 1d / ppt numbering) preserved
- Paperpile instrText/fldChar 77/231 = 77/231
- Bold runs 235 = 235
- Revisions: ins 34 / del 34; track revisions ON

Paragraphs were not rebuilt with python-docx. Citation fields and comment anchors match the original.

## Phase 0 backfills (done)

1. Abstract survival: 52 vs. 21 days; HR = 0.35, P = 0.003 (Methods already has n = 8 and log-rank; HR 95% CI and PH test still missing).
2. Abstract massive replaced with CD8 fractions (3.6-3.9% up to 58.5%).
3. CD86: 23.8-33.2% vs 2.64%.
4. In vivo survival sentence aligned with the abstract; n = 8, log-rank stated.
5. Eight HLA alleles in Results now have asterisks.
6. Two HLA asterisks in the Figure 4 legend.
7. Figure 2 legend: GSVA removed. Hallmark / FDR < 0.25 kept pending Figure 2d panel lock.

## Explicitly NOT copied from the modified Word file

- Figure 2d body still names Xenobiotic / KRAS / EMT (not MYC / apoptosis / angiogenesis).
- sq4 peptide count remains 12 (not 20) until a construct table exists.
- No ZCH coverage paragraph.

## Phase 3 / 4 text done

- Abstract IFN-gamma; experimentally validate; mRNA-LNP safety wording neutralized; one double space removed.
- IC50, (GGGGS)3, and 1 x 10^8 now use sub/superscripts.
- Kahles 30%: up to 30% more AS events vs normal, not 30% of events become tumor-specific.
- 59-fold: Background now attributes this to HCC; Discussion no longer uses the 59-fold point estimate on PDAC.
- Rojas: removed the 7% denominator and 2.5-3.6-fold claim; 25.0% peptide-level is not comparable to Rojas ~50% patient-level.
- Competing Interests: AuroraEpitope Limited and US provisional 63/869,610 (full inventor list still needs confirmation).
- Discussion enumerations (1)(2)(3)(4) -> (i)(ii)(iii)(iv).
- Methods now states MaxQuant used a global 1% FDR on the mixed LRDB.

## Phase 2 still blocked on raw records

1. ELISpot (M1): donor vs replicate collision; empty positive control; 16 vs 7.
2. Subgroup CI (M2): 3.28-3.67 is 6.6-fold narrower than the overall interval.
3. Class-specific FDR (M4) for the 18.4% non-canonical fraction.
4. CD8 baseline gating (M5) and >90% Tem vs durability language.
5. Empty LNP arm (M6) is in Methods but not in efficacy figures.
6. Survival HR 95% CI and PH assumption (M7).

## Phase 1 still open

- Figure 2b: body 48 vs panel 195/184.
- Figure 2d: scheme A (v3 panel) matches current body text.
- MaxQuant Methods v2.5.2 vs Figure 4 legend v2.1.
- Cell input 1e8 total vs 5e7 per line.
- Table_S5 is ELISpot and Table_S6 is scores (swapped vs body); no sq1-sq4 sheet; Chinese headers remain.

## Do not do yet

- Do not Accept All or delete comments.
- Do not click Paperpile Update.
- Do not continue on `pdac_V16_修改版.docx`.
