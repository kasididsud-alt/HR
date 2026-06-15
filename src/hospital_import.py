# src/hospital_import.py
from __future__ import annotations
import pandas as pd

REQUIRED = ["โรงพยาบาล", "Type", "Rate"]


def list_sheet_names(path: str) -> list[str]:
    xls = pd.ExcelFile(path)
    return list(xls.sheet_names)


def load_hospital_fee_map(
    path: str, sheet_name: str | int | None = None
) -> dict[tuple[str, int], float]:
    if sheet_name is None:
        df = pd.read_excel(path, sheet_name=0)
    else:
        df = pd.read_excel(path, sheet_name=sheet_name)

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(
            f"ไฟล์โรงพยาบาลขาดคอลัมน์: {', '.join(missing)} (ต้องมี โรงพยาบาล, Type, Rate)"
        )

    # normalize
    df["โรงพยาบาล"] = df["โรงพยาบาล"].astype(str).str.strip()
    df["Type"] = pd.to_numeric(df["Type"], errors="coerce").fillna(-1).astype(int)
    df["Rate"] = pd.to_numeric(df["Rate"], errors="coerce").fillna(0.0).astype(float)

    # build map: (hospital, type) -> rate
    fee_map: dict[tuple[str, int], float] = {}
    for _, r in df.iterrows():
        hosp = r["โรงพยาบาล"]
        t = int(r["Type"])
        rate = float(r["Rate"])
        if hosp and t >= 0:
            fee_map[(hosp, t)] = rate

    return fee_map
