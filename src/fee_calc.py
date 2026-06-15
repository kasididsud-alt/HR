# src/fee_calc.py
from __future__ import annotations

from datetime import date
from typing import Dict, Tuple
import re
import os

import pandas as pd

from src.employee_store import plan_to_set
from src.case_import import is_empty_person_value
from src.payroll_calc import (
    RATE_TABLES,
)  # A–D tiers (min_exclusive, max_inclusive, base)


# =========================
# ✅ Helpers
# =========================
def _num(x) -> float:
    """Safe numeric conversion; non-numeric -> 0.0"""
    try:
        v = pd.to_numeric(x, errors="coerce")
        if pd.isna(v):
            return 0.0
        return float(v)
    except Exception:
        return 0.0


def _positive_int_or_default(x, default: int) -> int:
    v = _num(x)
    if v <= 0:
        return int(default)
    return int(v)


def _clean_key(value) -> str:
    try:
        if value is None or pd.isna(value):
            return ""
    except Exception:
        if value is None:
            return ""
    text = str(value).strip()
    return "" if text.lower() in ("nan", "none") else text


def parse_discount_value(x) -> float:
    """
    ✅ รองรับรูปแบบส่วนลดหลายแบบ:
      - 0.5           -> 0.5 (50%)
      - "0.5"         -> 0.5
      - "50%" / "50 %"-> 0.5
      - 50 / "50"     -> 50.0 (บาท)
      - ""/None       -> 0.0
    """
    if x is None:
        return 0.0
    s = str(x).strip()
    if not s:
        return 0.0

    s = s.replace(",", "").strip()

    if "%" in s:
        s2 = s.replace("%", "").strip()
        v = pd.to_numeric(s2, errors="coerce")
        if pd.isna(v):
            return 0.0
        return float(v) / 100.0

    v = pd.to_numeric(s, errors="coerce")
    if pd.isna(v):
        return 0.0
    return float(v)


def full_months_worked(start: pd.Timestamp, end: pd.Timestamp) -> int:
    """Count full months worked between start and end."""
    if pd.isna(start) or pd.isna(end):
        return 0
    start = pd.to_datetime(start, errors="coerce").normalize()
    end = pd.to_datetime(end, errors="coerce").normalize()
    if pd.isna(start) or pd.isna(end) or end < start:
        return 0
    months = (end.year - start.year) * 12 + (end.month - start.month)
    if end.day < start.day:
        months -= 1
    return max(months, 0)


def is_active_on(emp_row: pd.Series, on_date: pd.Timestamp) -> bool:
    """Return True when the employee has already started by on_date."""
    start = pd.to_datetime(emp_row.get("start_date"), errors="coerce")
    on_date = pd.to_datetime(on_date, errors="coerce")
    if pd.isna(start) or pd.isna(on_date):
        return False
    return on_date.normalize() >= start.normalize()


def pick_rate_from_set(set_name: str, months_worked: int) -> float:
    """Pick base rate from RATE_TABLES by set (A/B/C/D) and months worked."""
    tiers = RATE_TABLES.get(str(set_name).strip().upper())
    if not tiers:
        return 0.0
    for mn_ex, mx_in, base in tiers:
        if months_worked > mn_ex and months_worked <= mx_in:
            return float(base)
    return 0.0


# =========================
# ✅ Employee rates
# =========================
def current_head_rate(
    emp_row: pd.Series, as_of: pd.Timestamp, divide_by: float = 2.0
) -> float:
    """
    Monitor fee base (ก่อนหักส่วนลด):
    - FREE => 0
    - FIX => fix_rate
    - COND15 => คิดใน compute_case_fees() (ต้องอิงลำดับเคสรายเดือน)
    - SET => จาก plan_to_set() + months_worked + RATE_TABLES
    """
    mode = str(emp_row.get("rate_mode", "")).strip().upper()
    plan = str(emp_row.get("rate_plan", "")).strip()

    if mode == "FREE":
        return 0.0
    if mode == "FIX":
        v = emp_row.get("fix_rate", 0)
        try:
            return float(v) if v not in ("", None) else 0.0
        except Exception:
            return 0.0

    if mode == "COND15":
        return 0.0  # computed later

    # SET (5.1-5.4)
    if not is_active_on(emp_row, as_of):
        return 0.0

    set_name = plan_to_set(plan)
    months = full_months_worked(emp_row.get("start_date"), as_of)
    base = pick_rate_from_set(set_name, months)
    return base / divide_by


def scoring_rate(emp_row: pd.Series | None) -> float:
    """
    Scoring fee from employee master:
    - scoring_mode = FREE => 0
    - scoring_mode = FIX => scoring_fix
    """
    if emp_row is None:
        return 0.0
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


# =========================
# ✅ Scoring prefix rule
# =========================
def _is_md_prefix_400(prefix: str) -> bool:
    """
    ถ้าคำนำหน้าขึ้นต้นด้วย นพ / พญ => 400
    """
    p = str(prefix or "").strip().replace(" ", "")
    return p.startswith("นพ") or p.startswith("พญ")


def has_scoring_doctor_prefix(prefix: str) -> bool:
    """Scoring Tech is treated as a doctor only when the prefix cell is filled."""
    return not is_empty_person_value(prefix)


# =========================
# ✅ Doctor normalize
# =========================
DOCTOR_PREFIXES = [
    "นพ.",
    "นพ",
    "พญ.",
    "พญ",
    "น.พ.",
    "น.พ",
    "พ.ญ.",
    "พ.ญ",
]


def normalize_doctor_name(name: str) -> str:
    s = str(name or "").strip()
    if not s:
        return ""
    s = s.replace("\u00a0", " ").strip()
    s = re.sub(r"\s+", " ", s)

    compact = s.replace(" ", "")
    for p in DOCTOR_PREFIXES:
        p2 = p.replace(" ", "")
        if compact.startswith(p2):
            compact = compact[len(p2) :]
            break
    compact = compact.strip(" .")
    return compact


def _doctor_name_candidates(name: str) -> set[str]:
    """Build multiple comparable forms for robust name matching."""
    raw = str(name or "").strip()
    if not raw:
        return set()

    out = set()

    low = re.sub(r"\s+", " ", raw).strip().lower()
    if low:
        out.add(low)
        out.add(low.replace(" ", ""))

    norm = normalize_doctor_name(raw)
    if norm:
        norm_low = re.sub(r"\s+", " ", norm).strip().lower()
        if norm_low:
            out.add(norm_low)
            out.add(norm_low.replace(" ", ""))

    return out


def _date_key_from_value(v, fallback: pd.Timestamp) -> str:
    dt = pd.to_datetime(v, errors="coerce")
    if pd.notna(dt):
        return dt.strftime("%Y-%m-%d")
    return pd.to_datetime(fallback).strftime("%Y-%m-%d")


def _load_names_by_date(master_dir: str, file_prefix: str) -> dict[str, set[str]]:
    """Read files: {prefix}YYYY-MM-DD.txt and return mapping date -> normalized names."""
    by_date: dict[str, set[str]] = {}
    if not master_dir:
        return by_date

    try:
        for fn in os.listdir(master_dir):
            if not (fn.startswith(file_prefix) and fn.endswith(".txt")):
                continue
            date_str = fn[len(file_prefix) : -4].strip()
            if not date_str:
                continue

            path = os.path.join(master_dir, fn)
            if not os.path.isfile(path):
                continue

            names: set[str] = set()
            with open(path, "r", encoding="utf-8") as f:
                for line in f:
                    s = str(line).strip()
                    if not s or s.startswith("#"):
                        continue
                    names.update(_doctor_name_candidates(s))

            by_date[date_str] = names
    except Exception:
        pass

    return by_date


def _load_free_doctors_by_date(master_dir: str) -> dict[str, set[str]]:
    return _load_names_by_date(master_dir, "free_doctors_")


def _load_free_physicians_by_date(master_dir: str) -> dict[str, set[str]]:
    return _load_names_by_date(master_dir, "free_physicians_")


# =========================
# ✅ Main
# =========================
def compute_case_fees(
    cases_df: pd.DataFrame,
    employees_df: pd.DataFrame,
    emp_map: Dict[str, pd.Series],
    alias_map: Dict[Tuple[str, str], str],
    hospital_fee_map: Dict[tuple[str, int], float],
    master_dir: str,
    as_of: pd.Timestamp | None = None,
) -> Tuple[
    pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame
]:
    """
    Return:
      - case_out
      - payroll_summary
      - pay_details
      - physician_by_hospital
      - doctor_summary
      - doctor_details
    """
    if as_of is None:
        as_of = pd.Timestamp(date.today())

    case_out = cases_df.copy()

    # Local import to avoid circular
    from src.master_loader import resolve_employee

    def resolve_emp(name: str, source_col: str) -> pd.Series | None:
        return resolve_employee(emp_map, alias_map, name, source_col)

    # Ensure output columns exist
    has_scoring_fee2_col = "Scoring Fee2" in case_out.columns
    for col in ["Monitor Fee", "Monitor Fee2", "Scoring Fee", "Physician Fee"]:
        if col not in case_out.columns:
            case_out[col] = 0.0

    free_doctors_by_date = _load_free_doctors_by_date(master_dir)
    free_physicians_by_date = _load_free_physicians_by_date(master_dir)

    def is_doctor_free(doctor_name, case_date):
        if not doctor_name:
            return False
        date_key = _date_key_from_value(case_date, as_of)
        free_names = free_doctors_by_date.get(date_key, set())
        if not free_names:
            return False
        return len(_doctor_name_candidates(doctor_name).intersection(free_names)) > 0

    def is_physician_free(doctor_name, case_date):
        if not doctor_name:
            return False
        date_key = _date_key_from_value(case_date, as_of)
        free_names = free_physicians_by_date.get(date_key, set())
        if not free_names:
            return False
        return len(_doctor_name_candidates(doctor_name).intersection(free_names)) > 0

    # -----------------------------
    # PREPASS: build COND15 sequence per emp per month
    # ✅ stable order: (วันที่ตรวจ, original row index)
    # ✅ count BOTH slots separately (if same person appears in both slots, count twice)
    # -----------------------------
    if "วันที่ตรวจ" not in case_out.columns:
        case_out["วันที่ตรวจ"] = as_of

    case_out["_idx"] = case_out.index
    case_out["_ตรวจ_dt"] = pd.to_datetime(case_out["วันที่ตรวจ"], errors="coerce").fillna(
        as_of
    )
    case_out["_month_key"] = case_out["_ตรวจ_dt"].dt.to_period("M").astype(str)

    # resolve emp_code for monitor slots
    case_out["_emp1"] = ""
    case_out["_emp2"] = ""

    for i, r in case_out.iterrows():
        mt1 = str(r.get("Monitor Tech", "") or "").strip()
        mt2 = str(r.get("Monitor Tech (2)", "") or "").strip()

        e1 = resolve_emp(mt1, "Monitor Tech") if mt1 else None
        e2 = resolve_emp(mt2, "Monitor Tech (2)") if mt2 else None

        case_out.at[i, "_emp1"] = e1["emp_code"] if e1 is not None else ""
        case_out.at[i, "_emp2"] = e2["emp_code"] if e2 is not None else ""

    # stable sort: date + original row order (do NOT use IV)
    tmp_cols = ["_idx", "_ตรวจ_dt", "_month_key", "_emp1", "_emp2"]
    tmp = case_out[tmp_cols].copy().sort_values(["_ตรวจ_dt", "_idx"], kind="mergesort")

    long_rows: list[dict] = []
    for _, r in tmp.iterrows():
        idx0 = r["_idx"]
        mk = r["_month_key"]
        case_date_for_seq = r["_ตรวจ_dt"]
        e1 = str(r.get("_emp1", "") or "")
        e2 = str(r.get("_emp2", "") or "")

        emp1_for_seq = emp_map.get(e1) if e1 else None
        emp2_for_seq = emp_map.get(e2) if e2 else None

        if e1 and emp1_for_seq is not None and is_active_on(emp1_for_seq, case_date_for_seq):
            long_rows.append({"_idx": idx0, "_month_key": mk, "_emp": e1, "_slot": 1})
        if e2 and emp2_for_seq is not None and is_active_on(emp2_for_seq, case_date_for_seq):
            long_rows.append({"_idx": idx0, "_month_key": mk, "_emp": e2, "_slot": 2})

    long_df = pd.DataFrame(long_rows)

    case_out["_seq_emp1"] = None
    case_out["_seq_emp2"] = None

    if not long_df.empty:
        long_df["_seq"] = long_df.groupby(["_emp", "_month_key"]).cumcount() + 1
        seq_map = {
            (r["_idx"], r["_emp"], int(r["_slot"])): int(r["_seq"])
            for _, r in long_df.iterrows()
        }

        for i, r in case_out.iterrows():
            idx0 = r["_idx"]
            e1 = str(r.get("_emp1", "") or "")
            e2 = str(r.get("_emp2", "") or "")
            case_out.at[i, "_seq_emp1"] = (
                seq_map.get((idx0, e1, 1), None) if e1 else None
            )
            case_out.at[i, "_seq_emp2"] = (
                seq_map.get((idx0, e2, 2), None) if e2 else None
            )

    # -----------------------------
    # Main loop
    # -----------------------------
    pay_rows: list[dict] = []
    phy_rows: list[dict] = []
    doctor_rows: list[dict] = []

    for idx, row in case_out.iterrows():
        case_key = _clean_key(row.get("IV", ""))
        if not case_key:
            case_key = str(idx)
        join_key = str(idx)
        case_out.at[idx, "_join_key"] = join_key

        mt1 = str(row.get("Monitor Tech", "") or "").strip()
        mt2 = str(row.get("Monitor Tech (2)", "") or "").strip()
        case_date_for_rate = row.get("_ตรวจ_dt", as_of)

        hosp = str(row.get("โรงพยาบาล", "") or "").strip()
        t_raw = row.get("Type", None)
        t_num = pd.to_numeric(t_raw, errors="coerce")
        t = int(t_num) if pd.notna(t_num) else -1

        # -----------------------------
        # ✅ Discount logic
        # - 15 เคสแรก: amount (บาท)
        # - ส่วนลด: supports 0.5 and "50%"
        #     * 0 < x < 1 => percent
        #     * x >= 1    => amount (บาท)
        # -----------------------------
        disc_first15_amt = max(_num(row.get("15 เคสแรก", 0)), 0.0)

        disc_raw = parse_discount_value(row.get("ส่วนลด", 0))
        disc_percent = 0.0
        disc_amount = 0.0
        if 0.0 < disc_raw < 1.0:
            disc_percent = disc_raw
        elif disc_raw >= 1.0:
            disc_amount = disc_raw

        discount_amount_total = disc_first15_amt + disc_amount

        has_mt1 = bool(mt1)
        has_mt2 = bool(mt2)

        # split "amount discount" between two monitors (percent does not split)
        if has_mt1 and has_mt2:
            disc_amt_share_1 = discount_amount_total / 2.0
            disc_amt_share_2 = discount_amount_total / 2.0
        else:
            disc_amt_share_1 = discount_amount_total if has_mt1 else 0.0
            disc_amt_share_2 = discount_amount_total if has_mt2 else 0.0

        # resolve employees
        emp1 = resolve_emp(mt1, "Monitor Tech") if mt1 else None
        emp2 = resolve_emp(mt2, "Monitor Tech (2)") if mt2 else None

        # -----------------------------
        # Monitor Fee (slot 1)
        # -----------------------------
        fee1 = 0.0
        fee1_free_note = ""
        if emp1 is not None:
            mode = str(emp1.get("rate_mode", "")).strip().upper()
            if not is_active_on(emp1, case_date_for_rate):
                fee1 = 0.0
            elif mode == "COND15":
                n_free = _positive_int_or_default(emp1.get("cond_free_first_n", 15), 15)
                after_rate = _num(emp1.get("cond_after_fix_rate", 0))
                seq = case_out.at[idx, "_seq_emp1"]
                is_free_seq = seq is not None and int(seq) <= n_free
                fee1 = 0.0 if is_free_seq else after_rate
                if is_free_seq:
                    fee1_free_note = f"15 เคสแรกฟรี ({int(seq)}/{n_free})"
            else:
                fee1 = current_head_rate(emp1, case_date_for_rate)

        # apply discount: percent first, then amount-share
        fee1 = float(fee1) * (1.0 - float(disc_percent))
        fee1 = max(0.0, float(fee1) - float(disc_amt_share_1))
        case_out.at[idx, "Monitor Fee"] = float(fee1)

        if emp1 is not None and (fee1 or fee1_free_note):
            pay_rows.append(
                {
                    "role": "MonitorTech",
                    "emp_code": emp1["emp_code"],
                    "display_name": emp1.get("display_name", emp1.get("full_name", "")),
                    "amount": float(fee1),
                    "case_key": case_key,
                    "join_key": join_key,
                    "source_col": "Monitor Tech",
                    "source_name": mt1,
                    "note": fee1_free_note,
                    "scoring_tech_in_case": str(
                        row.get("Scoring Tech", "") or ""
                    ).strip(),
                    "interpret_md_in_case": str(
                        row.get("Interpret MD", "") or ""
                    ).strip(),
                }
            )

        # -----------------------------
        # Monitor Fee2 (slot 2)  ✅ ALWAYS PAY IF APPLICABLE (even if same person)
        # -----------------------------
        fee2 = 0.0
        fee2_free_note = ""
        if emp2 is not None:
            mode2 = str(emp2.get("rate_mode", "")).strip().upper()
            if not is_active_on(emp2, case_date_for_rate):
                fee2 = 0.0
            elif mode2 == "COND15":
                n_free2 = _positive_int_or_default(emp2.get("cond_free_first_n", 15), 15)
                after_rate2 = _num(emp2.get("cond_after_fix_rate", 0))
                seq2 = case_out.at[idx, "_seq_emp2"]
                is_free_seq2 = seq2 is not None and int(seq2) <= n_free2
                fee2 = 0.0 if is_free_seq2 else after_rate2
                if is_free_seq2:
                    fee2_free_note = f"15 เคสแรกฟรี ({int(seq2)}/{n_free2})"
            else:
                fee2 = current_head_rate(emp2, case_date_for_rate)

            fee2 = float(fee2) * (1.0 - float(disc_percent))
            fee2 = max(0.0, float(fee2) - float(disc_amt_share_2))

        case_out.at[idx, "Monitor Fee2"] = float(fee2)

        if emp2 is not None and (fee2 or fee2_free_note):
            pay_rows.append(
                {
                    "role": "MonitorTech2",
                    "emp_code": emp2["emp_code"],
                    "display_name": emp2.get("display_name", emp2.get("full_name", "")),
                    "amount": float(fee2),
                    "case_key": case_key,
                    "join_key": join_key,
                    "source_col": "Monitor Tech (2)",
                    "source_name": mt2,
                    "note": fee2_free_note,
                    "scoring_tech_in_case": str(
                        row.get("Scoring Tech", "") or ""
                    ).strip(),
                    "interpret_md_in_case": str(
                        row.get("Interpret MD", "") or ""
                    ).strip(),
                }
            )

        # -----------------------------
        # Scoring Fee (RULE 1–4)
        # -----------------------------
        scoring_tech_name = str(row.get("Scoring Tech", "") or "").strip()
        if is_empty_person_value(scoring_tech_name):
            scoring_tech_name = ""
            case_out.at[idx, "Scoring Tech"] = ""
        prefix_raw = str(row.get("คำนำหน้า", "") or "").strip()

        sfee = 0.0
        emp_s: pd.Series | None = None

        case_date = row.get("วันที่ตรวจ", as_of)

        if scoring_tech_name:
            if _is_md_prefix_400(prefix_raw):
                sfee = 400.0
            else:
                emp_s = resolve_emp(scoring_tech_name, "Scoring Tech")
                if emp_s is not None:
                    sfee = float(scoring_rate(emp_s))
                else:
                    sfee = 500.0

            if is_doctor_free(scoring_tech_name, case_date):
                sfee = 0.0

        case_out.at[idx, "Scoring Fee"] = float(sfee)
        if has_scoring_fee2_col:
            case_out.at[idx, "Scoring Fee2"] = float(sfee)

        if scoring_tech_name:
            if emp_s is None:
                emp_s = resolve_emp(scoring_tech_name, "Scoring Tech")
            if emp_s is not None and sfee:
                pay_rows.append(
                    {
                        "role": "ScoringTech",
                        "emp_code": emp_s["emp_code"],
                        "display_name": emp_s.get(
                            "display_name", emp_s.get("full_name", "")
                        ),
                        "amount": float(sfee),
                        "case_key": case_key,
                        "join_key": join_key,
                        "source_col": "Scoring Tech",
                        "source_name": scoring_tech_name,
                        "note": "",
                        "scoring_tech_in_case": scoring_tech_name,
                        "interpret_md_in_case": str(
                            row.get("Interpret MD", "") or ""
                        ).strip(),
                    }
                )

        # -----------------------------
        # Physician Fee
        # -----------------------------
        pfee = float(hospital_fee_map.get((hosp, t), 0.0))
        
        interpret = str(row.get("Interpret MD", "") or "").strip()
        scoring_tech_name = str(row.get("Scoring Tech", "") or "").strip()
        if is_empty_person_value(scoring_tech_name):
            scoring_tech_name = ""

        # Check if interpret doctor is free
        if interpret and (is_doctor_free(interpret, case_date) or is_physician_free(interpret, case_date)):
            pfee = 0.0
        
        case_out.at[idx, "Physician Fee"] = float(pfee)

        if interpret and pfee:
            phy_rows.append(
                {
                    "Interpret MD": interpret,
                    "โรงพยาบาล": hosp,
                    "Type": t,
                    "Physician Fee": float(pfee),
                }
            )

        # -----------------------------
        # Doctor tracking
        # -----------------------------
        scoring_is_doctor = scoring_tech_name and has_scoring_doctor_prefix(prefix_raw)

        if scoring_is_doctor:
            # Check if scoring doctor is free
            sfee_for_doctor = sfee
            if is_doctor_free(scoring_tech_name, case_date):
                sfee_for_doctor = 0.0
                
            key = normalize_doctor_name(scoring_tech_name)
            if key:
                doctor_rows.append(
                    {
                        "doctor_key": key,
                        "doctor_name_raw": scoring_tech_name,
                        "source_col": "Scoring Tech",
                        "case_key": case_key,
                        "join_key": join_key,
                        "โรงพยาบาล": hosp,
                        "Type": t,
                        "amount": float(sfee_for_doctor),
                    }
                )

        if interpret:
            # Check if interpret doctor is free
            pfee_for_doctor = pfee
            if is_doctor_free(interpret, case_date) or is_physician_free(interpret, case_date):
                pfee_for_doctor = 0.0
                
            key = normalize_doctor_name(interpret)
            if key:
                doctor_rows.append(
                    {
                        "doctor_key": key,
                        "doctor_name_raw": interpret,
                        "source_col": "Interpret MD",
                        "case_key": case_key,
                        "join_key": join_key,
                        "โรงพยาบาล": hosp,
                        "Type": t,
                        "amount": float(pfee_for_doctor),
                    }
                )

    # -----------------------------
    # Build outputs
    # -----------------------------
    pay_details = pd.DataFrame(pay_rows)

    if not pay_details.empty:
        payroll_summary = pay_details.groupby(
            ["emp_code", "display_name"], as_index=False
        ).agg(
            found_count=("amount", "size"),
            total_amount=("amount", "sum"),
        )
    else:
        payroll_summary = pd.DataFrame(
            columns=["emp_code", "display_name", "found_count", "total_amount"]
        )

    phy_df = pd.DataFrame(phy_rows)
    if not phy_df.empty:
        physician_by_hospital = (
            phy_df.groupby(["Interpret MD", "โรงพยาบาล", "Type"], as_index=False)[
                "Physician Fee"
            ]
            .agg(["count", "sum"])
            .reset_index()
        )
    else:
        physician_by_hospital = pd.DataFrame(
            columns=["Interpret MD", "โรงพยาบาล", "Type", "count", "sum"]
        )

    doctor_details = pd.DataFrame(doctor_rows)
    if not doctor_details.empty:
        doctor_count_col = "join_key" if "join_key" in doctor_details.columns else "case_key"
        doctor_summary = doctor_details.groupby(["doctor_key"], as_index=False).agg(
            found_count=(doctor_count_col, "nunique"),
            total_amount=("amount", "sum"),
        )
    else:
        doctor_summary = pd.DataFrame(
            columns=["doctor_key", "found_count", "total_amount"]
        )

    # drop helper cols (don’t export)
    for c in [
        "_idx",
        "_ตรวจ_dt",
        "_month_key",
        "_emp1",
        "_emp2",
        "_seq_emp1",
        "_seq_emp2",
    ]:
        if c in case_out.columns:
            case_out.drop(columns=[c], inplace=True)

    return (
        case_out,
        payroll_summary,
        pay_details,
        physician_by_hospital,
        doctor_summary,
        doctor_details,
    )
