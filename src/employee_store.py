# src/employee_store.py
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Literal, Optional

import pandas as pd

RatePlan = Literal["5.1", "5.2", "5.3", "5.4", "5.5", "5.6"]
ScoringMode = Literal["FREE", "FIX"]

EMP_SHEET = "employees"

# ✅ ไม่มี rate_set แล้ว
EMP_COLUMNS = [
    "emp_code",
    "first_name",
    "last_name",
    "full_name",
    "display_name",
    "start_date",  # yyyy-mm-dd
    # rate
    "rate_plan",  # 5.1..5.6
    "rate_mode",  # SET / FIX / FREE / COND15
    "fix_rate",  # ตัวเลข (ถ้า FIX)
    # เงื่อนไข plan 5.6
    "cond_free_first_n",  # 15
    "cond_after_fix_rate",  # 750 หรือ 1500
    "cond_pay_type",  # ประกันสังคม / เงินสด
    # scoring
    "scoring_mode",  # FIX / FREE
    "scoring_fix",  # 200/500 (ถ้า FIX)
    # misc
    "active",
    "note",
]


@dataclass
class EmployeeInput:
    first_name: str
    last_name: str
    start_date: pd.Timestamp

    rate_plan: RatePlan
    fix_rate: Optional[float] = None

    # plan 5.6
    cond_free_first_n: int = 15
    cond_after_fix_rate: Optional[int] = None  # 750/1500
    cond_pay_type: str = ""  # ประกันสังคม/เงินสด

    scoring_mode: ScoringMode = "FREE"
    scoring_fix: Optional[int] = None

    display_name: str = ""
    note: str = ""


def _ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in EMP_COLUMNS:
        if c not in df.columns:
            df[c] = ""
    return df


def load_employees_xlsx(master_dir: str) -> pd.DataFrame:
    path = os.path.join(master_dir, "employees.xlsx")
    if not os.path.exists(path):
        df = pd.DataFrame(columns=EMP_COLUMNS)
        df["active"] = True
        return df

    df = pd.read_excel(path, sheet_name=EMP_SHEET)
    df = _ensure_columns(df)

    # normalize
    df["emp_code"] = df["emp_code"].astype(str).str.strip()
    df["first_name"] = df["first_name"].astype(str).str.strip()
    df["last_name"] = df["last_name"].astype(str).str.strip()
    df["full_name"] = df["full_name"].astype(str).str.strip()
    df["display_name"] = df["display_name"].astype(str).str.strip()

    df["rate_plan"] = df["rate_plan"].astype(str).str.strip()
    df["rate_mode"] = df["rate_mode"].astype(str).str.strip().str.upper()
    df["scoring_mode"] = df["scoring_mode"].astype(str).str.strip().str.upper()

    df["active"] = df["active"].fillna(True).astype(bool)
    df["start_date"] = pd.to_datetime(df["start_date"], errors="coerce")

    return df


def save_employees_xlsx(master_dir: str, df: pd.DataFrame) -> None:
    os.makedirs(master_dir, exist_ok=True)
    path = os.path.join(master_dir, "employees.xlsx")

    df = _ensure_columns(df)

    out = df.copy()
    out["start_date"] = pd.to_datetime(out["start_date"], errors="coerce")

    with pd.ExcelWriter(path, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name=EMP_SHEET)


def generate_next_emp_code(existing_codes: list[str], prefix="E", width=3) -> str:
    """
    สร้างรหัส E001, E002,... ไม่ซ้ำ โดยดู max index จากที่มีอยู่
    """
    max_n = 0
    pat = re.compile(rf"^{re.escape(prefix)}(\d+)$", re.IGNORECASE)
    for code in existing_codes:
        m = pat.match(str(code).strip())
        if not m:
            continue
        try:
            n = int(m.group(1))
            max_n = max(max_n, n)
        except ValueError:
            continue
    return f"{prefix}{str(max_n + 1).zfill(width)}"


def _rate_mode_from_plan(inp: EmployeeInput) -> str:
    if inp.rate_plan in ("5.1", "5.2", "5.3", "5.4"):
        return "SET"
    if inp.rate_plan == "5.5":
        return "FREE" if inp.fix_rate is None else "FIX"
    if inp.rate_plan == "5.6":
        return "COND15"
    return "SET"


def validate_employee_input(inp: EmployeeInput) -> None:
    if not inp.first_name.strip():
        raise ValueError("กรุณากรอกชื่อ")
    if not inp.last_name.strip():
        raise ValueError("กรุณากรอกนามสกุล")
    if pd.isna(inp.start_date):
        raise ValueError("กรุณาเลือกวันเริ่มทำงาน")

    if inp.rate_plan == "5.5":
        if inp.fix_rate is not None and float(inp.fix_rate) < 0:
            raise ValueError("Fix Rate ต้องไม่ติดลบ")

    if inp.rate_plan == "5.6":
        if inp.cond_free_first_n <= 0:
            raise ValueError("จำนวนเคสฟรี (15 เคสแรก) ต้องมากกว่า 0")
        if inp.cond_after_fix_rate not in (750, 1500):
            raise ValueError("แผน 5.6 ต้องเลือกอัตราหลัง 15 เคสแรกเป็น 750 หรือ 1500")
        # ให้โปรแกรม set pay type ให้อัตโนมัติได้ แต่ถ้าส่งมาก็ validate
        if inp.cond_pay_type and inp.cond_pay_type.strip() not in (
            "ประกันสังคม",
            "เงินสด",
        ):
            raise ValueError("แผน 5.6 ประเภทการจ่ายต้องเป็น: ประกันสังคม หรือ เงินสด")

    # scoring
    if inp.scoring_mode == "FIX":
        if inp.scoring_fix not in (200, 500):
            raise ValueError("Scoring FIX ต้องเป็น 200 หรือ 500")


def add_employee(master_dir: str, inp: EmployeeInput) -> str:
    """
    เพิ่มพนักงานใหม่ลง employees.xlsx และคืนค่า emp_code ที่สร้างได้
    """
    validate_employee_input(inp)

    df = load_employees_xlsx(master_dir)
    existing_codes = df["emp_code"].dropna().astype(str).tolist()
    new_code = generate_next_emp_code(existing_codes)

    full_name = f"{inp.first_name.strip()} {inp.last_name.strip()}".strip()
    display_name = inp.display_name.strip() or full_name

    rate_mode = _rate_mode_from_plan(inp)
    fix_rate = (
        float(inp.fix_rate) if (rate_mode == "FIX" and inp.fix_rate is not None) else ""
    )

    cond_free_first_n = inp.cond_free_first_n if rate_mode == "COND15" else ""
    cond_after_fix_rate = inp.cond_after_fix_rate if rate_mode == "COND15" else ""
    # auto pay type ตามอัตรา (กันกรอกผิด)
    if rate_mode == "COND15":
        pay_type = "ประกันสังคม" if inp.cond_after_fix_rate == 750 else "เงินสด"
    else:
        pay_type = ""

    scoring_mode = inp.scoring_mode
    scoring_fix = inp.scoring_fix if scoring_mode == "FIX" else ""

    row = {
        "emp_code": new_code,
        "first_name": inp.first_name.strip(),
        "last_name": inp.last_name.strip(),
        "full_name": full_name,
        "display_name": display_name,
        "start_date": pd.to_datetime(inp.start_date),
        "rate_plan": inp.rate_plan,
        "rate_mode": rate_mode,
        "fix_rate": fix_rate,
        "cond_free_first_n": cond_free_first_n,
        "cond_after_fix_rate": cond_after_fix_rate,
        "cond_pay_type": pay_type,
        "scoring_mode": scoring_mode,
        "scoring_fix": scoring_fix,
        "active": True,
        "note": inp.note.strip(),
    }

    df = _ensure_columns(df)
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    save_employees_xlsx(master_dir, df)
    return new_code


def plan_to_set(plan: str) -> str:
    """
    ใช้ตอนคำนวณ: map rate_plan -> set ตาราง A/B/C/D
    """
    return {
        "5.1": "A",
        "5.2": "B",
        "5.3": "C",
        "5.4": "D",
    }.get(str(plan).strip(), "")
