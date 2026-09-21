# Companion code for the PDAC long-read neoantigen / mRNA-LNP vaccine manuscript

**Paper:** Long-Read Sequencing-Enhanced Immunopeptidomics Unveils Non-Canonical Neoepitopes for mRNA-LNP Vaccination against Pancreatic Cancer

**Authors:** Xiangeng Wang, Hualiang Yao, Ming Zhong, Ziwei Wang, Tianli Luo, Yi Shuai, Tao Jiang, Hang Jiang, Xin Wang

This repository is the **manuscript companion packet** (version 17): the Word draft, main figures, supplementary tables, and the scripts used to revise the text and rebuild Table S2. It is the code/data folder that goes with the article, not a generic dump.

本仓库是论文配套材料（V17）：正文稿、主图、补充表，以及修订文稿、重建 Table S2 的脚本。

Corresponding author: Xin Wang (xinwang@cuhk.edu.hk)

## What maps to the paper

| Paper item | In this repo |
| --- | --- |
| Main text (V17) | manuscript/V17.docx |
| Original V16 source (read-only) | manuscript/V17_base.docx |
| Main figures | igures/2026-08-20_PDAC_mainfigure_v3.pptx (body-matched); v4 for comparison |
| Supplementary tables | 	ables/Supplementary_Tables_V17.xlsx |
| Table S2 (DEAS, Figure 2b/2c) | 	ables/Table_S2.xlsx and 	ables/Table_S2_Differential_AS.csv |
| Track-changes / verification scripts | scripts/ |

Table S2: JP / GSE196009, 13 tumor vs 6 normal; 7,434 tested events; 48 pass |dPSI| >= 0.2 and P <= 0.01. This DEAS table does **not** replace Figure 2a landscape counts (23,514 events).

## Keywords

PDAC; long-read sequencing; novel isoforms; non-canonical neoantigens; immunopeptidomics; mRNA-LNP vaccine; alternative splicing

## Layout

`
manuscript/   article text (V17.docx)
figures/      main-figure PowerPoint
tables/       supplementary tables + Table S2
scripts/      revision and Table S2 rebuild tools
`

Open manuscript/V17.docx in Word. Do not click Paperpile Update. Do not Accept All or delete comments.

## Scripts

`
pip install -r requirements.txt
python scripts/integrity_check.py
python scripts/verify_v17.py
python scripts/rebuild_table_s2.py   # needs local RDS via PDAC_RDS_PATH
python scripts/make_table_s2.py
`

The long-read / DEAS / immunopeptidome analysis pipeline is **not** in this repository. This folder is the article packet only.

## License

Unpublished manuscript materials. All rights reserved.
