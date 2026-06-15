# src/payroll_calc.py
import pandas as pd
from datetime import date

from src.employee_store import plan_to_set

# ✅ Rate table อยู่ในโปรแกรม (base rate ก่อน /2)
RATE_TABLES = {
    "A": [(-1, 3, 300), (3, 6, 500), (6, 9999, 800)],
    "B": [(-1, 4, 300), (4, 8, 500), (8, 9999, 800)],
    "C": [(-1, 4, 500), (4, 8, 800), (8, 9999, 1000)],  # >=9 => >8
    "D": [(-1, 4, 500), (4, 8, 800), (8, 9999, 1200)],  # >=9 => >8
}


def full_months_worked(start: pd.Timestamp, end: pd.Timestamp) -> int:
    """แบบ A: เดือนเต็มตามปฏิทิน"""
    if pd.isna(start) or pd.isna(end):
        return 0
    start = pd.to_datetime(start).normalize()
    end = pd.to_datetime(end).normalize()
    if end < start:
        return 0
    months = (end.year - start.year) * 12 + (end.month - start.month)
    if end.day < start.day:
        months -= 1
    return max(months, 0)


def is_active_on(emp_row: pd.Series, on_date: pd.Timestamp) -> bool:
    start = pd.to_datetime(emp_row.get("start_date"), errors="coerce")
    on_date = pd.to_datetime(on_date, errors="coerce")
    if pd.isna(start) or pd.isna(on_date):
        return False
    return on_date.normalize() >= start.normalize()


def pick_rate_from_set(set_name: str, months_worked: int) -> float:
    set_name = str(set_name).strip().upper()
    tiers = RATE_TABLES.get(set_name)
    if not tiers:
        return 0.0
    for mn_ex, mx_in, rate in tiers:
        if months_worked > mn_ex and months_worked <= mx_in:
            return float(rate)
    return 0.0


def head_rate_for_employee(
    emp_row: pd.Series, as_of: pd.Timestamp, divide_by: float = 2.0
) -> float:
    plan = str(emp_row.get("rate_plan", "")).strip()
    mode = str(emp_row.get("rate_mode", "")).strip().upper()

    if mode == "FREE":
        return 0.0

    if mode == "FIX":
        v = emp_row.get("fix_rate", 0)
        try:
            return float(v) if v not in ("", None) else 0.0
        except Exception:
            return 0.0

    if mode == "COND15":
        # แผน 5.6 ต้องนับ 15 เคสแรก/เดือน => ยังไม่คิดในฟังก์ชันนี้
        return 0.0

    # SET (5.1-5.4)
    if not is_active_on(emp_row, as_of):
        return 0.0

    set_name = plan_to_set(plan)  # 5.1->A, 5.2->B, 5.3->C, 5.4->D
    months = full_months_worked(emp_row.get("start_date"), as_of)
    base = pick_rate_from_set(set_name, months)
    return base / divide_by


def scoring_fee_for_employee(emp_row: pd.Series) -> float:
    mode = str(emp_row.get("scoring_mode", "FREE")).strip().upper()
    if mode == "FREE":
        return 0.0
    if mode == "FIX":
        v = emp_row.get("scoring_fix", 0)
        try:
            return float(v) if v not in ("", None) else 0.0
        except Exception:
            return 0.0
    return 0.0


def compute_from_cases(
    cases_df: pd.DataFrame,
    employees_df: pd.DataFrame,
    emp_map: dict,
    alias_map: dict,
    role_columns: dict,
    as_of: pd.Timestamp | None = None,
):
    if as_of is None:
        as_of = pd.Timestamp(date.today())

    case_result = cases_df.copy()
    pay_rows = []

    def add_pay(emp_row: pd.Series, role: str, amount: float):
        pay_rows.append(
            {
                "emp_code": emp_row.get("emp_code", ""),
                "full_name": emp_row.get("full_name", ""),
                "display_name": emp_row.get(
                    "display_name", emp_row.get("full_name", "")
                ),
                "role": role,
                "amount": float(amount),
            }
        )

    def resolve_emp(name_in_excel: str, source_col: str):
        from src.master_loader import resolve_employee

        return resolve_employee(emp_map, alias_map, name_in_excel, source_col)

    def case_date_for_row(row: pd.Series) -> pd.Timestamp:
        for col in row.index:
            if "ตรวจ" in str(col):
                dt = pd.to_datetime(row.get(col), errors="coerce")
                if pd.notna(dt):
                    return dt
        return as_of

    for _, row in case_result.iterrows():
        case_date = case_date_for_row(row)
        for source_col, role in role_columns.items():
            person_name = row.get(source_col, None)
            if person_name is None or str(person_name).strip() == "":
                continue

            emp_row = resolve_emp(str(person_name), source_col)
            if emp_row is None:
                continue

            if role in ("MonitorTech", "MonitorTech2", "InterpretMD"):
                amt = head_rate_for_employee(emp_row, case_date)
                if amt != 0:
                    add_pay(emp_row, role, amt)

            elif role == "ScoringTech":
                amt = scoring_fee_for_employee(emp_row)
                if amt != 0:
                    add_pay(emp_row, role, amt)

    payroll_long = pd.DataFrame(pay_rows)
    if payroll_long.empty:
        payroll_summary = pd.DataFrame(
            columns=["role", "emp_code", "display_name", "amount"]
        )
    else:
        payroll_summary = (
            payroll_long.groupby(["role", "emp_code", "display_name"], as_index=False)[
                "amount"
            ]
            .sum()
            .sort_values(["role", "amount"], ascending=[True, False])
        )

    return case_result, payroll_summary
