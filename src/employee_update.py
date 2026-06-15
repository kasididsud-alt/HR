# src/employee_update.py
from __future__ import annotations

import pandas as pd

from .employee_store import (
    EmployeeInput,
    EMP_COLUMNS,
    load_employees_xlsx,
    save_employees_xlsx,
    validate_employee_input,
)


def update_employee_by_code(master_dir: str, emp_code: str, inp: EmployeeInput) -> None:
    """
    อัปเดตพนักงานใน master/employees.xlsx โดยหาแถวจาก emp_code (ห้ามเปลี่ยนรหัส)
    - ถ้าไม่พบ emp_code => raise ValueError
    - validate ตามกติกาเดียวกับ add_employee
    """
    emp_code = str(emp_code).strip()
    if not emp_code:
        raise ValueError("emp_code ว่าง")

    validate_employee_input(inp)

    df = load_employees_xlsx(master_dir)

    # หาแถวที่ต้องแก้
    mask = df["emp_code"].astype(str).str.strip() == emp_code
    if not mask.any():
        raise ValueError(f"ไม่พบพนักงานรหัส {emp_code}")

    # สร้างค่าที่จะเขียนกลับ (คุม format ให้เหมือน add_employee)
    full_name = f"{inp.first_name.strip()} {inp.last_name.strip()}".strip()
    display_name = inp.display_name.strip() or full_name

    rate_plan = inp.rate_plan
    rate_mode = _rate_mode_from_plan(rate_plan, inp.fix_rate)

    fix_rate = (
        float(inp.fix_rate) if (rate_mode == "FIX" and inp.fix_rate is not None) else ""
    )

    # plan 5.6
    cond_free_first_n = inp.cond_free_first_n if rate_mode == "COND15" else ""
    cond_after_fix_rate = inp.cond_after_fix_rate if rate_mode == "COND15" else ""
    if rate_mode == "COND15":
        pay_type = "ประกันสังคม" if inp.cond_after_fix_rate == 750 else "เงินสด"
    else:
        pay_type = ""

    scoring_mode = inp.scoring_mode
    scoring_fix = inp.scoring_fix if scoring_mode == "FIX" else ""

    # อัปเดตลง df (อัปเดตเฉพาะคอลัมน์ที่เราคุม)
    update_map = {
        "first_name": inp.first_name.strip(),
        "last_name": inp.last_name.strip(),
        "full_name": full_name,
        "display_name": display_name,
        "start_date": pd.to_datetime(inp.start_date),
        "rate_plan": rate_plan,
        "rate_mode": rate_mode,
        "fix_rate": fix_rate,
        "cond_free_first_n": cond_free_first_n,
        "cond_after_fix_rate": cond_after_fix_rate,
        "cond_pay_type": pay_type,
        "scoring_mode": scoring_mode,
        "scoring_fix": scoring_fix,
        "note": inp.note.strip(),
        # active ไม่แตะใน update นี้ (ถ้าจะทำปุ่มปิดใช้งานค่อยแยกอีกฟังก์ชัน)
    }

    # กันกรณีไฟล์เก่าขาดคอลัมน์
    for col in EMP_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    # อัปเดตแถวทั้งหมดที่ match (ปกติควรมี 1 แถว)
    for col, val in update_map.items():
        df.loc[mask, col] = val

    save_employees_xlsx(master_dir, df)


def _rate_mode_from_plan(rate_plan: str, fix_rate) -> str:
    rate_plan = str(rate_plan).strip()
    if rate_plan in ("5.1", "5.2", "5.3", "5.4"):
        return "SET"
    if rate_plan == "5.5":
        return "FREE" if fix_rate is None else "FIX"
    if rate_plan == "5.6":
        return "COND15"
    return "SET"
