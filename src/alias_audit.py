import pandas as pd
from .master_loader import resolve_employee

STRICT_EMPLOYEE_COLS = ("Monitor Tech", "Monitor Tech (2)")


def find_unknown_names(
    cases_df: pd.DataFrame, emp_map: dict, alias_map: dict
) -> pd.DataFrame:
    unknown = []
    for col in STRICT_EMPLOYEE_COLS:
        if col not in cases_df.columns:
            continue
        vals = cases_df[col].dropna().astype(str).str.strip().unique()
        for v in vals:
            if not v:
                continue
            emp = resolve_employee(emp_map, alias_map, v, source_col=col)
            if emp is None:
                unknown.append({"source_col": col, "alias_name": v})
    if not unknown:
        return pd.DataFrame(columns=["source_col", "alias_name"])
    return (
        pd.DataFrame(unknown)
        .drop_duplicates()
        .sort_values(["source_col", "alias_name"])
    )
