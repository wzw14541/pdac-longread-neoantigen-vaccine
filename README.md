# PDAC V17 稿件工作区

未发表稿件的私有工作副本（alternative splicing neoantigen / mRNA-LNP vaccine，V17），以及在不破坏 Paperpile 域和 Word 批注的前提下做修订的脚本。

**未发表，保留所有权利。请保持仓库私有，不要公开。**

## 目录

```
manuscript/     V17.docx (修订稿)、V17_base.docx (只读原稿)、CHANGELOG
figures/        主图 PowerPoint (v3 / v4)
tables/         补充表格与 Table S2
scripts/        修订、校验、重建 Table S2
```

当前稿件：`manuscript/V17.docx`  
**不要**点击 Paperpile Update。  
**不要** Accept All 或删除批注。

## 表格

| 文件 | 作用 |
| --- | --- |
| `tables/Supplementary_Tables_V17.xlsx` | 当前补充表格工作簿 |
| `tables/Table_S2.xlsx` | 独立 Table S2（DEAS + JP PSI） |
| `tables/Table_S2_Differential_AS.csv` | 7,434 条 DEAS；48 条通过 abs(dPSI)>=0.2 且 P<=0.01 |

Table S2 来自 JP / GSE196009（13 tumor vs 6 normal）。它**不能**替代 Figure 2a 的全景计数（23,514 条）。

重建 S2 所需的 RDS / PSI 仍留在分析机上，不进本仓库。本机可设置：

```
set PDAC_RDS_PATH=C:\ai\pdac\nag\dt\jp_evt_rst.rds
set PDAC_PSI_PATH=C:\ai\pdac\nag\dt\as\jp_evt.psi
```

## 脚本

在仓库根目录执行 `pip install -r requirements.txt` 后：

```
python scripts/integrity_check.py      # 批注 / Paperpile / 修订标记 vs V17_base
python scripts/verify_v17.py           # V17.docx 关键短语核对
python scripts/ooxml_revise.py         # OOXML 修订（已应用到当前稿）
python scripts/ooxml_fixup.py          # IFN / 枚举 / 双句点收尾
python scripts/apply_v17_edits.py      # Word COM 替代方案（Windows + Word）
python scripts/rebuild_table_s2.py     # 从本机 RDS 重建 S2 并写入 V17 工作簿
python scripts/make_table_s2.py        # 由 CSV（及可选 PSI）生成独立 Table_S2.xlsx
```

`ooxml_revise.py` 只拆分/包裹 run，写入 `w:del` / `w:ins`，不用 python-docx 重建段落。

## 图

`figures/2026-08-20_PDAC_mainfigure_v3.pptx` 与当前正文一致（Figure 2d scheme A）。v4 留作对照；图注尚未锁死。

## 不包含的内容

长读长 / DEAS / 免疫肽组分析管线不在本仓库（原机器上的 `C:\ai\pdac`）。这里只放 V17 稿件包。
