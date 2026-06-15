import pandas as pd

ROLE_COLUMNS = {
    "Monitor Tech": "MonitorTech",
    "Monitor Tech (2)": "MonitorTech2",
    "Scoring Tech": "ScoringTech",
    "Interpret MD": "InterpretMD",
}


def load_employees(master_dir: str = "master") -> pd.DataFrame:
    emp = pd.read_excel(f"{master_dir}/employees.xlsx", sheet_name="employees")
    emp["start_date"] = pd.to_datetime(emp["start_date"], errors="coerce")
    emp["emp_code"] = emp["emp_code"].astype(str).str.strip()
    emp["active"] = emp["active"].fillna(True).astype(bool)

    # normalize
    for c in ["rate_mode", "rate_set", "scoring_mode"]:
        if c in emp.columns:
            emp[c] = emp[c].astype(str).str.strip().str.upper()

    return emp[emp["active"]].copy()


def load_aliases(master_dir: str = "master") -> pd.DataFrame:
    al = pd.read_excel(f"{master_dir}/aliases.xlsx", sheet_name="aliases")
    al["alias_name"] = al["alias_name"].astype(str).str.strip()
    al["emp_code"] = al["emp_code"].astype(str).str.strip()
    al["source_col"] = al.get("source_col", "").fillna("").astype(str).str.strip()
    al["active"] = al["active"].fillna(True).astype(bool)
    return al[al["active"]].copy()


def load_rate_tables(master_dir: str = "master") -> pd.DataFrame:
    rt = pd.read_excel(f"{master_dir}/rate_tables.xlsx", sheet_name="rate_sets")
    rt["set"] = rt["set"].astype(str).str.strip().str.upper()
    return rt


def build_lookups(employees: pd.DataFrame, aliases: pd.DataFrame):
    emp_map = {r["emp_code"]: r for _, r in employees.iterrows()}

    alias_map = {}
    for _, r in aliases.iterrows():
        key = (r["alias_name"].lower(), r["source_col"])
        alias_map[key] = r["emp_code"]
        # fallback: no source col
        key2 = (r["alias_name"].lower(), "")
        alias_map.setdefault(key2, r["emp_code"])

    return emp_map, alias_map


def resolve_employee(
    emp_map: dict, alias_map: dict, name_in_excel: str, source_col: str = ""
):
    if name_in_excel is None:
        return None
    name = str(name_in_excel).strip()
    if not name:
        return None

    key = (name.lower(), source_col.strip())
    key2 = (name.lower(), "")

    emp_code = alias_map.get(key) or alias_map.get(key2)
    if emp_code:
        return emp_map.get(emp_code)

    # ✅ FALLBACK: match against employee names directly
    name_l = name.lower()
    for emp in emp_map.values():
        disp = str(emp.get("display_name", "") or "").strip().lower()
        full = str(emp.get("full_name", "") or "").strip().lower()
        first = str(emp.get("first_name", "") or "").strip().lower()
        last = str(emp.get("last_name", "") or "").strip().lower()
        full2 = (first + " " + last).strip()

        if name_l and (name_l == disp or name_l == full or name_l == full2):
            return emp

    return None
