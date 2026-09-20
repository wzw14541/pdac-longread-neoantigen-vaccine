# -*- coding: utf-8 -*-
"""Repository paths. All other scripts should import from here."""
from __future__ import print_function

import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MANUSCRIPT = ROOT / "manuscript"
FIGURES = ROOT / "figures"
TABLES = ROOT / "tables"
LOGS = ROOT / "logs"

V17 = MANUSCRIPT / "V17.docx"
V17_BASE = MANUSCRIPT / "V17_base.docx"

SUPP_V17 = TABLES / "Supplementary_Tables_V17.xlsx"
SUPP_ORIGINAL = TABLES / "Supplementary_Tables_RealPDAC_revised.xlsx"
TABLE_S2 = TABLES / "Table_S2.xlsx"
TABLE_S2_CSV = TABLES / "Table_S2_Differential_AS.csv"
TABLE_S2_NOTE = TABLES / "Table_S2_source.txt"

# Optional local analysis files — not in this repo.
RDS = Path(os.environ.get("PDAC_RDS_PATH", r"C:\ai\pdac\nag\dt\jp_evt_rst.rds"))
PSI = Path(os.environ.get("PDAC_PSI_PATH", r"C:\ai\pdac\nag\dt\as\jp_evt.psi"))


def require(path, label):
    path = Path(path)
    if not path.is_file():
        raise SystemExit("Missing {}: {}".format(label, path))
    return path
