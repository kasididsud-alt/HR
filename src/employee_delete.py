# src/employee_delete.py
from __future__ import annotations

import pandas as pd

from .employee_store import load_employees_xlsx, save_employees_xlsx, EMP_COLUMNS


def soft_delete_employee(master_dir: str, emp_code: str) -> None:
    """
    ✅ แนะนำ: Soft delete = active=False
    """
    emp_code = str(emp_code).strip()
    if not emp_code:
        raise ValueError("emp_code ว่าง")

    df = load_employees_xlsx(master_dir)

    # กันกรณีไฟล์เก่าขาดคอลัมน์ active
    if "active" not in df.columns:
        df["active"] = True
    # กันคอลัมน์ขาด
    for c in EMP_COLUMNS:
        if c not in df.columns:
            df[c] = ""

    mask = df["emp_code"].astype(str).str.strip() == emp_code
    if not mask.any():
        raise ValueError(f"ไม่พบพนักงานรหัส {emp_code}")

    df.loc[mask, "active"] = False
    save_employees_xlsx(master_dir, df)


def hard_delete_employee(master_dir: str, emp_code: str) -> None:
    """
    🔴 Hard delete = ลบแถวทิ้งจริง ๆ
    """
    emp_code = str(emp_code).strip()
    if not emp_code:
        raise ValueError("emp_code ว่าง")

    df = load_employees_xlsx(master_dir)
    mask = df["emp_code"].astype(str).str.strip() == emp_code
    if not mask.any():
        raise ValueError(f"ไม่พบพนักงานรหัส {emp_code}")

    df2 = df.loc[~mask].copy()
    save_employees_xlsx(master_dir, df2)
