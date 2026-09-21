# Companion code and supplementary tables

**Paper:** Long-Read Sequencing-Enhanced Immunopeptidomics Unveils Non-Canonical Neoepitopes for mRNA-LNP Vaccination against Pancreatic Cancer

**Authors:** Xiangeng Wang, Hualiang Yao, Ziwei Wang, Ming Zhong, Tianli Luo, Yi Shuai, Tao Jiang, Hang Jiang, Xin Wang

This repository holds the **scripts and supplementary tables** that accompany the article.

本仓库放的是论文配套的脚本与补充表。

Corresponding author: Xin Wang (xinwang@cuhk.edu.hk)

## What maps to the paper

| Paper item | File / sheet |
| --- | --- |
| Table S1. SQANTI3 isoform classification (22,784 isoforms in the master transcriptome; 30,510 records before QC) | tables/Supplementary_Tables.xlsx, sheet S1 Isoform Master |
| Table S2a. Isoform-level PSI matrix | sheet S2a Isoform PSI matrix |
| Table S2b. Differential AS events, JP/GSE196009, 13 tumor vs 6 normal; 7,434 tested, 48 significant (|dPSI|>=0.2 and P<=0.01) | sheet S2b; also tables/Table_S2.xlsx and tables/Table_S2_Differential_AS.csv |
| Table S2c. Per-sample SUPPA2 PSI for AS events | sheet S2c AS Event PSI Matrix |
| Table S3. ORF predictions | sheet S3 ORF Predictions |
| Table S4a. HLA class I binding predictions | sheet S4a HLA Binding Predictions |
| Table S4b. DeepLC retention-time validation | sheet S4b DeepLC RT Validation |
| Table S5. Vaccine constructs SQ1-SQ4 | sheet S5a Vaccine Construct Comp |
| Table S5c. Known HLA-A*24:02 reference epitopes in SQ4 | sheet S5b Known Reference Epitopes |
| Table S6a. Integrated ranking of candidate neoepitopes | sheet S6a Candidate Ranking |
| Table S6b/c. IFN-gamma ELISpot (3 HLA-A*24:02 donors; 64 peptides; 16 positive at >200 SFU) | sheets S6b and S6c |

Figure 2a landscape counts (10,871 novel + 12,643 known events) are **not** replaced by Table S2b. S2b is the 7,434-event DEAS subset actually tested.

## Keywords

PDAC; long-read sequencing; novel isoforms; non-canonical neoantigens; immunopeptidomics; mRNA-LNP vaccine; alternative splicing

## Layout

```
tables/    supplementary workbook + Table S2 extract
scripts/   rebuild Table S2b from local RDS / CSV
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
