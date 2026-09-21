# Companion code for PDAC long-read neoantigen / mRNA-LNP vaccination

**Paper:** Long-Read Sequencing-Enhanced Immunopeptidomics Unveils Non-Canonical Neoepitopes for mRNA-LNP Vaccination against Pancreatic Cancer

**Authors:** Xiangeng Wang, Hualiang Yao, Ming Zhong, Ziwei Wang, Tianli Luo, Yi Shuai, Tao Jiang, Hang Jiang, Xin Wang

This repository holds the **scripts and supplementary tables** that accompany the article.

本仓库放的是论文配套的脚本与补充表。

Corresponding author: Xin Wang (xinwang@cuhk.edu.hk)

## What maps to the paper

| Paper item | In this repo |
| --- | --- |
| Supplementary tables | 	ables/Supplementary_Tables.xlsx |
| Table S2 (DEAS; Figure 2b/2c) | 	ables/Table_S2.xlsx, 	ables/Table_S2_Differential_AS.csv |
| Table S2 rebuild scripts | scripts/rebuild_table_s2.py, scripts/make_table_s2.py |

Table S2: JP / GSE196009, 13 tumor vs 6 normal; 7,434 tested events; 48 pass abs(dPSI) >= 0.2 and P <= 0.01. This DEAS table does **not** replace Figure 2a landscape counts (23,514 events).

## Keywords

PDAC; long-read sequencing; novel isoforms; non-canonical neoantigens; immunopeptidomics; mRNA-LNP vaccine; alternative splicing

## Layout

```
tables/    supplementary tables and Table S2
scripts/   rebuild Table S2 from local RDS / CSV
```

## Scripts

```
pip install -r requirements.txt
python scripts/rebuild_table_s2.py   # needs local RDS via PDAC_RDS_PATH
python scripts/make_table_s2.py      # CSV (+ optional PSI) -> Table_S2.xlsx
```

The long-read / DEAS / immunopeptidome analysis pipeline is not in this repository.

## License

All rights reserved.
