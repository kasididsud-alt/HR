from __future__ import annotations

from typing import Optional

import pandas as pd
from PySide6 import QtCore, QtWidgets

from .excel_io import to_excel_bytes
from .fee_calc import normalize_doctor_name


DATE_COL = "\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48\u0e15\u0e23\u0e27\u0e08"
HOSPITAL_COL = "\u0e42\u0e23\u0e07\u0e1e\u0e22\u0e32\u0e1a\u0e32\u0e25"
REFER_COL = "\u0e2b\u0e21\u0e2d\u0e2a\u0e48\u0e07"


class DoctorDetailDialog(QtWidgets.QDialog):
    def __init__(
        self,
        doctor_key: str,
        doctor_details_df: pd.DataFrame,
        case_out_df: pd.DataFrame,
        parent: Optional[QtWidgets.QWidget] = None,
    ):
        super().__init__(parent)
        self.doctor_key = str(doctor_key)

        self.setWindowTitle(
            f"\u0e23\u0e32\u0e22\u0e25\u0e30\u0e40\u0e2d\u0e35\u0e22\u0e14\u0e41\u0e1e\u0e17\u0e22\u0e4c: {self.doctor_key}"
        )
        self.resize(1280, 760)

        self.df_all = self._build_df_all(doctor_details_df, case_out_df)
        self.df_view = self.df_all.copy()

        self.lbl_summary = QtWidgets.QLabel()
        self.lbl_summary.setStyleSheet("font-weight: bold; font-size: 14px;")

        self.search = QtWidgets.QLineEdit()
        self.search.setPlaceholderText(
            "\u0e04\u0e49\u0e19\u0e2b\u0e32 "
            f"({DATE_COL}/IV/{HOSPITAL_COL}/{REFER_COL}/Scoring Tech/Interpret MD)..."
        )
        self.search.textChanged.connect(self._apply_filter)

        self.table = QtWidgets.QTableWidget()
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)

        headers = [
            DATE_COL,
            HOSPITAL_COL,
            REFER_COL,
            "Scoring Tech",
            "Interpret MD",
            "Scoring Fee",
            "Physician Fee",
            "Total",
            "IV / case_key",
        ]
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)

        hh = self.table.horizontalHeader()
        hh.setStretchLastSection(True)
        hh.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(4, QtWidgets.QHeaderView.ResizeToContents)
        for col in [5, 6, 7]:
            hh.setSectionResizeMode(col, QtWidgets.QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(8, QtWidgets.QHeaderView.ResizeToContents)

        self.btn_back = QtWidgets.QPushButton(
            "\u0e01\u0e25\u0e31\u0e1a"
        )
        self.btn_back.clicked.connect(self.close)

        self.btn_export = QtWidgets.QPushButton("Export Excel")
        self.btn_export.clicked.connect(self._export_excel)

        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addWidget(self.btn_back)
        btn_row.addStretch(1)
        btn_row.addWidget(self.btn_export)

        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(self.lbl_summary)
        layout.addWidget(self.search)
        layout.addWidget(self.table, stretch=1)
        layout.addLayout(btn_row)

        self._refresh_summary()
        self._render_table()

    def _build_df_all(
        self, doctor_details_df: pd.DataFrame, case_out_df: pd.DataFrame
    ) -> pd.DataFrame:
        _ = doctor_details_df
        columns = [
            DATE_COL,
            HOSPITAL_COL,
            REFER_COL,
            "Scoring Tech",
            "Interpret MD",
            "Scoring Fee",
            "Physician Fee",
            "total_amount",
            "case_key",
        ]

        dfc = case_out_df.copy() if case_out_df is not None else pd.DataFrame()
        if dfc is None or dfc.empty:
            return pd.DataFrame(columns=columns)

        for col in [DATE_COL, HOSPITAL_COL, REFER_COL, "Scoring Tech", "Interpret MD", "IV"]:
            if col not in dfc.columns:
                dfc[col] = ""
        for col in ["Scoring Fee", "Physician Fee"]:
            if col not in dfc.columns:
                dfc[col] = 0.0

        dfc["_scoring_key"] = dfc["Scoring Tech"].map(normalize_doctor_name)
        dfc["_interpret_key"] = dfc["Interpret MD"].map(normalize_doctor_name)

        dfc = dfc[
            (dfc["_scoring_key"].astype(str) == self.doctor_key)
            | (dfc["_interpret_key"].astype(str) == self.doctor_key)
        ].copy()
        if dfc.empty:
            return pd.DataFrame(columns=columns)

        dfc["Scoring Fee"] = pd.to_numeric(dfc["Scoring Fee"], errors="coerce").fillna(
            0.0
        )
        dfc["Physician Fee"] = pd.to_numeric(
            dfc["Physician Fee"], errors="coerce"
        ).fillna(0.0)

        dfc["Scoring Fee"] = dfc["Scoring Fee"].where(
            dfc["_scoring_key"].astype(str) == self.doctor_key, 0.0
        )
        dfc["Physician Fee"] = dfc["Physician Fee"].where(
            dfc["_interpret_key"].astype(str) == self.doctor_key, 0.0
        )
        dfc["total_amount"] = dfc["Scoring Fee"] + dfc["Physician Fee"]

        case_key = dfc["IV"].astype(str).str.strip()
        if "case_key" in dfc.columns:
            fallback_case_key = dfc["case_key"].astype(str).str.strip()
            case_key = case_key.where(case_key != "", fallback_case_key)
        if "_join_key" in dfc.columns:
            join_fallback = dfc["_join_key"].astype(str).str.strip()
            case_key = case_key.where(case_key != "", join_fallback)
        case_key = case_key.where(case_key != "", dfc.index.astype(str))
        dfc["case_key"] = case_key

        try:
            dt = pd.to_datetime(dfc[DATE_COL], errors="coerce")
            mask = dt.notna()
            dfc.loc[mask, DATE_COL] = dt.loc[mask].dt.strftime("%Y-%m-%d")
        except Exception:
            pass

        return dfc[columns].copy()

    def _refresh_summary(self) -> None:
        total_scoring = (
            float(self.df_all["Scoring Fee"].sum()) if not self.df_all.empty else 0.0
        )
        total_phys = (
            float(self.df_all["Physician Fee"].sum()) if not self.df_all.empty else 0.0
        )
        total = float(self.df_all["total_amount"].sum()) if not self.df_all.empty else 0.0
        count = int(len(self.df_all))

        self.lbl_summary.setText(
            f"\u0e41\u0e1e\u0e17\u0e22\u0e4c: {self.doctor_key} | "
            f"\u0e08\u0e33\u0e19\u0e27\u0e19\u0e40\u0e04\u0e2a: {count:,d} | "
            f"\u0e23\u0e27\u0e21 Scoring Fee: {total_scoring:,.2f} | "
            f"\u0e23\u0e27\u0e21 Physician Fee: {total_phys:,.2f} | "
            f"\u0e23\u0e27\u0e21\u0e17\u0e31\u0e49\u0e07\u0e2b\u0e21\u0e14: {total:,.2f}"
        )

    def _render_table(self) -> None:
        df = self.df_view
        self.table.setRowCount(0)
        if df is None or df.empty:
            return

        df = df.copy()
        try:
            df.sort_values([DATE_COL, "case_key"], inplace=True, na_position="last")
        except Exception:
            pass

        self.table.setRowCount(len(df))
        for row_idx, (_, row) in enumerate(df.iterrows()):
            self.table.setItem(
                row_idx, 0, QtWidgets.QTableWidgetItem(str(row.get(DATE_COL, "")))
            )
            self.table.setItem(
                row_idx, 1, QtWidgets.QTableWidgetItem(str(row.get(HOSPITAL_COL, "")))
            )
            self.table.setItem(
                row_idx, 2, QtWidgets.QTableWidgetItem(str(row.get(REFER_COL, "")))
            )
            self.table.setItem(
                row_idx, 3, QtWidgets.QTableWidgetItem(str(row.get("Scoring Tech", "")))
            )
            self.table.setItem(
                row_idx, 4, QtWidgets.QTableWidgetItem(str(row.get("Interpret MD", "")))
            )

            def set_amount(col_idx: int, value) -> None:
                try:
                    num = float(value)
                except Exception:
                    num = 0.0
                item = QtWidgets.QTableWidgetItem(f"{num:,.2f}")
                item.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                self.table.setItem(row_idx, col_idx, item)

            set_amount(5, row.get("Scoring Fee", 0.0))
            set_amount(6, row.get("Physician Fee", 0.0))
            set_amount(7, row.get("total_amount", 0.0))

            self.table.setItem(
                row_idx, 8, QtWidgets.QTableWidgetItem(str(row.get("case_key", "")))
            )

        self.table.resizeRowsToContents()

    def _apply_filter(self, text: str) -> None:
        query = (text or "").strip().lower()
        if not query:
            self.df_view = self.df_all.copy()
            self._render_table()
            return

        df = self.df_all.copy()
        cols = [DATE_COL, HOSPITAL_COL, REFER_COL, "Scoring Tech", "Interpret MD", "case_key"]
        for col in cols:
            if col not in df.columns:
                df[col] = ""

        mask = None
        for col in cols:
            current = df[col].astype(str).str.lower().str.contains(query, na=False)
            mask = current if mask is None else (mask | current)

        self.df_view = df[mask].copy() if mask is not None else df.iloc[0:0].copy()
        self._render_table()

    def _export_excel(self) -> None:
        if self.df_all is None or self.df_all.empty:
            QtWidgets.QMessageBox.information(
                self,
                "Export",
                "\u0e44\u0e21\u0e48\u0e21\u0e35\u0e02\u0e49\u0e2d\u0e21\u0e39\u0e25\u0e43\u0e2b\u0e49 export",
            )
            return

        default_name = f"doctor_cases_{self.doctor_key}.xlsx"
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Export \u0e23\u0e32\u0e22\u0e25\u0e30\u0e40\u0e2d\u0e35\u0e22\u0e14\u0e41\u0e1e\u0e17\u0e22\u0e4c\u0e40\u0e1b\u0e47\u0e19 Excel",
            default_name,
            "Excel Files (*.xlsx)",
        )
        if not path:
            return

        try:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            out_bytes = to_excel_bytes({"DoctorDetails": self.df_all})
            with open(path, "wb") as file_obj:
                file_obj.write(out_bytes)
            QtWidgets.QMessageBox.information(
                self,
                "Export",
                f"\u0e1a\u0e31\u0e19\u0e17\u0e36\u0e01\u0e2a\u0e33\u0e40\u0e23\u0e47\u0e08:\n{path}",
            )
        except Exception as exc:
            QtWidgets.QMessageBox.critical(
                self,
                "Export",
                f"\u0e1a\u0e31\u0e19\u0e17\u0e36\u0e01\u0e25\u0e49\u0e21\u0e40\u0e2b\u0e25\u0e27:\n{exc}",
            )
