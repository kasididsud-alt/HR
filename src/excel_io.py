import io
import pandas as pd


def read_excel_bytes(uploaded_file) -> pd.DataFrame:
    return pd.read_excel(uploaded_file)


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, index=False, sheet_name=name[:31])
    return buf.getvalue()
