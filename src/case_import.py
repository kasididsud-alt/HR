# src/case_import.py
from __future__ import annotations

import pandas as pd


# ✅ Columns required in the Case Excel
# (ตัด Scoring Fee2 ออกแล้ว)
REQUIRED_COLS = [
    "ลำดับ",
    "จำนวนเคส",
    "วันที่ตรวจ",
    "ชื่อคนไข้",
    "HN",
    "โรงพยาบาล",
    "Type",
    "Program",
    "สิทธิ์",
    "หมอส่ง",
    "Monitor Tech",
    "Monitor Fee",
    "Monitor Tech (2)",
    "Monitor Fee2",
    "Scoring Fee",
    "คำนำหน้า",
    "Scoring Tech",
    # "Scoring Fee2",  # ❌ ไม่ require แล้ว
    "Interpret MD",
    "Physician Fee",
    "IV",
    "ราคารวมIV",
    "15 เคสแรก",
    "ส่วนลด",
]

# คน/ชื่อที่ต้อง normalize
PERSON_COLS = ["Monitor Tech", "Monitor Tech (2)", "Scoring Tech", "Interpret MD"]

# คอลัมน์ตัวเลข (ส่วนลดรวมอยู่ด้วย)
NUM_COLS = [
    "Monitor Fee",
    "Monitor Fee2",
    "Scoring Fee",
    "Physician Fee",
    "ราคารวมIV",
    "15 เคสแรก",
    "ส่วนลด",
]


def is_empty_person_value(value) -> bool:
    s = "" if value is None else str(value).strip()
    if not s:
        return True
    compact = "".join(ch for ch in s if ch not in " -–—_").strip().lower()
    return compact in {"", "nan", "none", "ไม่มี", "ไมมี"}


def list_sheet_names(path: str) -> list[str]:
    xls = pd.ExcelFile(path)
    return list(xls.sheet_names)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _parse_percent_or_number(x) -> float:
    """
    ✅ รองรับการกรอกส่วนลดหลายแบบ:
      - 0.5      -> 0.5
      - "0.5"    -> 0.5
      - "50%"    -> 0.5
      - "50 %"   -> 0.5
      - 50       -> 50.0  (บาท)
      - ""/None  -> 0.0
    """
    s = "" if x is None else str(x).strip()
    if not s:
        return 0.0

    s = s.replace(",", "").strip()

    if "%" in s:
        s2 = s.replace("%", "").strip()
        v = pd.to_numeric(s2, errors="coerce")
        return 0.0 if pd.isna(v) else float(v) / 100.0

    v = pd.to_numeric(s, errors="coerce")
    return 0.0 if pd.isna(v) else float(v)


def _parse_exam_date(series: pd.Series) -> pd.Series:
    """
    ✅ Robust date parsing:
    - รองรับ Excel serial date (ตัวเลข)
    - รองรับ ISO: YYYY-MM-DD (ต้อง parse ก่อน เพื่อกันสลับวัน/เดือน)
    - รองรับ dd/mm/yyyy (fallback ด้วย dayfirst=True)
    """
    s = series.astype(str).str.strip()

    # 1) Excel serial date (ตัวเลขล้วน)
    num = pd.to_numeric(s, errors="coerce")
    mask_num = num.notna()

    out = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

    if mask_num.any():
        out.loc[mask_num] = pd.to_datetime(
            num.loc[mask_num], unit="D", origin="1899-12-30", errors="coerce"
        )

    # 2) ISO first: YYYY-MM-DD
    mask_str = ~mask_num
    if mask_str.any():
        iso = pd.to_datetime(s.loc[mask_str], errors="coerce", format="%Y-%m-%d")
        out.loc[mask_str] = iso

        # 3) fallback: dayfirst=True (dd/mm/yyyy)
        mask_left = mask_str & out.loc[mask_str].isna()
        if mask_left.any():
            out.loc[mask_left] = pd.to_datetime(
                s.loc[mask_left], errors="coerce", dayfirst=True
            )

    return out


def load_cases_xlsx(path: str, sheet_name: str | int | None = None) -> pd.DataFrame:
    # default first sheet
    df = pd.read_excel(path, sheet_name=0 if sheet_name is None else sheet_name)
    df = normalize_columns(df)

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"ชีทนี้ขาดคอลัมน์: {', '.join(missing)}")

    # ✅ 날짜 parse (กัน day/month กลับ)
    df["วันที่ตรวจ"] = _parse_exam_date(df["วันที่ตรวจ"])

    # ✅ normalize person columns
    for c in PERSON_COLS:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].fillna("").astype(str).str.strip()
        df.loc[df[c].str.lower() == "nan", c] = ""
        df.loc[df[c].map(is_empty_person_value), c] = ""

    # ✅ รองรับ "50%" ในคอลัมน์ส่วนลด (แปลงก่อนเข้า to_numeric)
    if "ส่วนลด" in df.columns:
        df["ส่วนลด"] = df["ส่วนลด"].map(_parse_percent_or_number)

    # ✅ numeric cols
    for c in NUM_COLS:
        if c not in df.columns:
            df[c] = 0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    return df
