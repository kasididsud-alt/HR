# src/employee_list_ui.py
from __future__ import annotations

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QTableWidget,
    QTableWidgetItem,
    QPushButton,
    QMessageBox,
    QTextEdit,
)

from .employee_store import load_employees_xlsx, EMP_COLUMNS, plan_to_set
from .payroll_calc import RATE_TABLES  # Rate A–D
from .employee_edit_ui import EditEmployeeDialog
from .employee_delete import soft_delete_employee


# -------------------------
# Helpers
# -------------------------
def _safe_str(x) -> str:
    try:
        if x is None or pd.isna(x):
            return ""
    except Exception:
        if x is None:
            return ""
    return str(x)


def _fmt_date(x) -> str:
    try:
        if pd.isna(x):
            return ""
        ts = pd.to_datetime(x, errors="coerce")
        if pd.isna(ts):
            return _safe_str(x)
        return ts.strftime("%Y-%m-%d")
    except Exception:
        return _safe_str(x)


def _fmt_money(x) -> str:
    """
    ✅ format ที่ถูกต้อง: :,.0f (ไม่ใช่ :,0f)
    """
    try:
        if x is None or x == "" or (isinstance(x, float) and pd.isna(x)):
            return "0"
        v = float(x)
        return f"{v:,.0f}"
    except Exception:
        return _safe_str(x)


# ✅ แบบ A: เดือนเต็มปฏิทิน + วันคงเหลือ + วันรวม
def work_duration_calendar_A(
    start: pd.Timestamp, end: pd.Timestamp
) -> tuple[int, int, int]:
    """
    return (full_months, remaining_days, total_days)
    - full_months: เดือนเต็มตามปฏิทิน (แบบ A)
    - remaining_days: วันคงเหลือหลังหักเดือนเต็ม
    - total_days: วันรวมจริง
    """
    if pd.isna(start) or pd.isna(end):
        return 0, 0, 0

    start = pd.to_datetime(start, errors="coerce")
    end = pd.to_datetime(end, errors="coerce")
    if pd.isna(start) or pd.isna(end):
        return 0, 0, 0

    start = start.normalize()
    end = end.normalize()
    if end < start:
        return 0, 0, 0

    total_days = int((end - start).days)

    months = (end.year - start.year) * 12 + (end.month - start.month)
    if end.day < start.day:
        months -= 1
    months = max(months, 0)

    anchor = start + pd.DateOffset(months=months)
    remaining_days = int((end - anchor).days)
    remaining_days = max(remaining_days, 0)

    return months, remaining_days, total_days


def is_active_on(emp: dict, on_date: pd.Timestamp) -> bool:
    start = pd.to_datetime(emp.get("start_date", None), errors="coerce")
    on_date = pd.to_datetime(on_date, errors="coerce")
    if pd.isna(start) or pd.isna(on_date):
        return False
    return on_date.normalize() >= start.normalize()


def pick_rate_from_set(set_name: str, months_worked: int) -> float:
    """
    RATE_TABLES: list[(min_exclusive, max_inclusive, base)]
    """
    tiers = RATE_TABLES.get(str(set_name).strip().upper())
    if not tiers:
        return 0.0
    for mn_ex, mx_in, base in tiers:
        if months_worked > mn_ex and months_worked <= mx_in:
            return float(base)
    return 0.0


def compute_current_head_rate(emp: dict, as_of: pd.Timestamp) -> tuple[float, str]:
    """
    คำนวณรายหัว/ตอนนี้ ตาม master:
      - rate_mode=SET: ใช้ rate_plan -> plan_to_set -> RATE_TABLES -> base/2
      - rate_mode=FIX: ใช้ fix_rate
      - rate_mode=FREE: 0
      - rate_mode=COND15: (5.6) ยังต้องนับเคสรายเดือน => แสดง 0 พร้อม note
    """
    mode = str(emp.get("rate_mode", "") or "").strip().upper()
    plan = str(emp.get("rate_plan", "") or "").strip()

    if mode == "FREE":
        return 0.0, "FREE=0"

    if mode == "FIX":
        v = emp.get("fix_rate", 0)
        try:
            return float(v) if v not in ("", None) else 0.0, f"FIX={_safe_str(v)}"
        except Exception:
            return 0.0, "FIX invalid"

    if mode == "COND15":
        # ยังไม่ทำการนับ 15 เคสแรก/เดือนในส่วน employee view
        after_rate = emp.get("cond_after_fix_rate", "")
        pay_type = emp.get("cond_pay_type", "")
        return 0.0, f"5.6 ต้องนับเคส/เดือน (หลัง 15 = {after_rate} {pay_type})"

    # SET (5.1-5.4)
    if not is_active_on(emp, as_of):
        return 0.0, "not started"

    set_name = plan_to_set(plan)
    start = pd.to_datetime(emp.get("start_date", None), errors="coerce")
    m_full, _, _ = work_duration_calendar_A(start, as_of)
    base = pick_rate_from_set(set_name, m_full)
    head = base / 2.0

    return float(head), f"SET {set_name} base={base} /2 (เดือนเต็ม={m_full})"


def compute_scoring_fee(emp: dict) -> tuple[float, str]:
    """
    scoring_mode:
      - FIX -> scoring_fix (200/500)
      - FREE -> 0
    """
    mode = str(emp.get("scoring_mode", "FREE") or "FREE").strip().upper()
    if mode == "FREE":
        return 0.0, "FREE=0"

    if mode == "FIX":
        v = emp.get("scoring_fix", 0)
        try:
            return float(v) if v not in ("", None) else 0.0, f"FIX={_safe_str(v)}"
        except Exception:
            return 0.0, "FIX invalid"

    return 0.0, f"unknown scoring_mode={mode}"


def plan_display_text(emp: dict, months_full: int) -> str:
    """
    อธิบาย plan แบบอ่านง่าย (สำหรับ dialog รายละเอียด)
    """

    def fmt(x):  # แสดงตัวเลขแบบ comma
        try:
            return f"{float(x):,.0f}"
        except Exception:
            return str(x)

    plan = str(emp.get("rate_plan", "") or "").strip()

    if plan == "5.1":
        tier = (
            "ช่วง <=3 เดือน"
            if months_full <= 3
            else ("ช่วง >3<=6 เดือน" if months_full <= 6 else "ช่วง >=7 เดือน")
        )
        return (
            "แผน 5.1 (อิงเดือนเต็มแบบปฏิทิน)\n"
            f"• <= 3 เดือน  : {fmt(300)} /2\n"
            f"• > 3 <= 6    : {fmt(500)} /2\n"
            f"• >= 7        : {fmt(800)} /2\n"
            f"➡️ ตอนนี้: {tier} (เดือนเต็ม={months_full})"
        )

    if plan == "5.2":
        tier = (
            "ช่วง <=4 เดือน"
            if months_full <= 4
            else ("ช่วง >4<=8 เดือน" if months_full <= 8 else "ช่วง >=9 เดือน")
        )
        return (
            "แผน 5.2 (อิงเดือนเต็มแบบปฏิทิน)\n"
            f"• <= 4 เดือน  : {fmt(300)} /2\n"
            f"• > 4 <= 8    : {fmt(500)} /2\n"
            f"• >= 9        : {fmt(800)} /2\n"
            f"➡️ ตอนนี้: {tier} (เดือนเต็ม={months_full})"
        )

    if plan == "5.3":
        tier = (
            "ช่วง <=4 เดือน"
            if months_full <= 4
            else ("ช่วง >4<=8 เดือน" if months_full <= 8 else "ช่วง >=9 เดือน")
        )
        return (
            "แผน 5.3 (อิงเดือนเต็มแบบปฏิทิน)\n"
            f"• <= 4 เดือน  : {fmt(500)} /2\n"
            f"• > 4 <= 8    : {fmt(800)} /2\n"
            f"• >= 9        : {fmt(1000)} /2\n"
            f"➡️ ตอนนี้: {tier} (เดือนเต็ม={months_full})"
        )

    if plan == "5.4":
        tier = (
            "ช่วง <=4 เดือน"
            if months_full <= 4
            else ("ช่วง >4<=8 เดือน" if months_full <= 8 else "ช่วง >=9 เดือน")
        )
        return (
            "แผน 5.4 (อิงเดือนเต็มแบบปฏิทิน)\n"
            f"• <= 4 เดือน  : {fmt(500)} /2\n"
            f"• > 4 <= 8    : {fmt(800)} /2\n"
            f"• >= 9        : {fmt(1200)} /2\n"
            f"➡️ ตอนนี้: {tier} (เดือนเต็ม={months_full})"
        )

    if plan == "5.5":
        return (
            "แผน 5.5 (Fix Rate / Free)\n"
            "• เลือก Fix rate: 400, 500, 800, 1000, 1200, 1500, 2000, 750\n"
            "• หรือเลือก Free = 0\n"
            "➡️ หมายเหตุ: ค่า “รายหัว/ตอนนี้” = fix_rate ที่ตั้งไว้"
        )

    if plan == "5.6":
        return (
            "แผน 5.6 (มีเงื่อนไขรายเดือน)\n"
            "• 15 เคสแรก / เดือน = ฟรี\n"
            "• หลัง 15 เคสแรก = fix rate 750 (ประกันสังคม) หรือ fix rate 1500 (เงินสด)\n"
            "➡️ หมายเหตุ: จะรู้ “รายหัว/ตอนนี้” จริง ๆ ต้องนับจำนวนเคสของเดือนนั้นก่อน"
        )

    return "ไม่พบรายละเอียดแผน (rate_plan ว่างหรือไม่ถูกต้อง)"


# -------------------------
# Dialogs
# -------------------------
class EmployeeDetailDialog(QDialog):
    """Popup แสดงรายละเอียดพนักงาน 1 คน + อายุงาน + เรทปัจจุบัน + รายละเอียดแผน"""

    def __init__(self, row: dict, as_of: pd.Timestamp, parent=None):
        super().__init__(parent)
        self.setWindowTitle("รายละเอียดพนักงาน")
        self.resize(600, 680)

        layout = QVBoxLayout(self)

        title = QLabel(f"👤 {row.get('emp_code','')} - {row.get('full_name','')}")
        title.setStyleSheet("font-weight: 600; font-size: 14px;")
        layout.addWidget(title)

        start = pd.to_datetime(row.get("start_date", None), errors="coerce")
        m, d, _td = work_duration_calendar_A(start, as_of)

        head, head_note = compute_current_head_rate(row, as_of)
        scoring, scoring_note = compute_scoring_fee(row)

        rate_line = QLabel(
            f"📅 อายุงาน: {m} เดือน {d} วัน | 💰 รายหัว/ตอนนี้: {_fmt_money(head)} | 🧾 Scoring/ตอนนี้: {_fmt_money(scoring)}"
        )
        rate_line.setStyleSheet("color: #111; font-weight: 600;")
        layout.addWidget(rate_line)

        note_line = QLabel(f"ℹ️ รายละเอียดเรท: {head_note} | Scoring: {scoring_note}")
        note_line.setStyleSheet("color: #444;")
        note_line.setWordWrap(True)
        layout.addWidget(note_line)

        plan_box = QTextEdit()
        plan_box.setReadOnly(True)
        plan_box.setFixedHeight(170)
        plan_box.setText(plan_display_text(row, m))
        layout.addWidget(plan_box)

        box = QTextEdit()
        box.setReadOnly(True)
        lines = []
        for c in EMP_COLUMNS:
            v = row.get(c, "")
            if c == "start_date":
                v = _fmt_date(v)
            lines.append(f"{c}: {_safe_str(v)}")
        box.setText("\n".join(lines))
        layout.addWidget(box)

        btn = QPushButton("ปิด")
        btn.clicked.connect(self.accept)
        layout.addWidget(btn)


class EmployeeListDialog(QDialog):
    """หน้าดูพนักงานทั้งหมด + ค้นหา + แสดงอายุงาน + เรทปัจจุบัน + แก้ไข + ลบ"""

    def __init__(self, master_dir: str, parent=None, as_of: pd.Timestamp | None = None):
        super().__init__(parent)
        self.master_dir = master_dir
        self.as_of = pd.Timestamp.today().normalize()

        self.setWindowTitle("👥 รายการพนักงาน")
        self.resize(1100, 700)

        self.df_all = pd.DataFrame()
        self.df_view = pd.DataFrame()

        layout = QVBoxLayout(self)

        # --- Search bar
        top = QHBoxLayout()
        layout.addLayout(top)

        top.addWidget(QLabel("ค้นหา:"))
        self.search = QLineEdit()
        self.search.setPlaceholderText("พิมพ์ รหัส/ชื่อ/นามสกุล/ชื่อเต็ม/ชื่อแสดง/แผน แล้วค้นหา.")
        self.search.textChanged.connect(self.apply_filter)
        top.addWidget(self.search, 1)

        self.btn_refresh = QPushButton("รีเฟรช")
        self.btn_refresh.clicked.connect(self.reload)
        top.addWidget(self.btn_refresh)

        self.lbl_asof = QLabel(f"อายุงาน ณ วันนี้: {self.as_of.strftime('%Y-%m-%d')}")
        self.lbl_asof.setStyleSheet("color: #555;")
        top.addWidget(self.lbl_asof)

        self.lbl_count = QLabel("")
        self.lbl_count.setStyleSheet("color: #555;")
        top.addWidget(self.lbl_count)

        # --- Table
        self.table = QTableWidget(0, 8)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.cellDoubleClicked.connect(self.open_detail)
        layout.addWidget(self.table, 1)

        # --- Bottom buttons
        bottom = QHBoxLayout()
        layout.addLayout(bottom)

        self.btn_detail = QPushButton("ดูรายละเอียด (ดับเบิลคลิกแถวก็ได้)")
        self.btn_detail.clicked.connect(self.open_detail)
        bottom.addWidget(self.btn_detail)

        self.btn_edit = QPushButton("✏️ แก้ไข")
        self.btn_edit.clicked.connect(self.open_edit)
        bottom.addWidget(self.btn_edit)

        self.btn_delete = QPushButton("🗑️ ลบ (ปิดใช้งาน)")
        self.btn_delete.clicked.connect(self.delete_employee)
        bottom.addWidget(self.btn_delete)

        self.btn_close = QPushButton("ปิด")
        self.btn_close.clicked.connect(self.accept)
        bottom.addWidget(self.btn_close)

        self.reload()

    def reload(self):
        try:
            self.as_of = pd.Timestamp.today().normalize()
            self.lbl_asof.setText(f"อายุงาน ณ วันนี้: {self.as_of.strftime('%Y-%m-%d')}")
            self.df_all = load_employees_xlsx(self.master_dir)

            # only active
            if "active" in self.df_all.columns:
                self.df_all = self.df_all[self.df_all["active"] == True].copy()

            self._add_computed_columns(self.df_all)
            self.apply_filter()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def _add_computed_columns(self, df: pd.DataFrame):
        if df is None or df.empty:
            return

        tenure_text = []
        head_rates = []
        scoring_fees = []
        plan_texts = []

        for _, r in df.iterrows():
            start = pd.to_datetime(r.get("start_date", None), errors="coerce")
            m, d, _ = work_duration_calendar_A(start, self.as_of)
            tenure_text.append(f"{m} เดือน {d} วัน")

            head, _ = compute_current_head_rate(r.to_dict(), self.as_of)
            scoring, _ = compute_scoring_fee(r.to_dict())

            head_rates.append(head)
            scoring_fees.append(scoring)
            plan_texts.append(plan_display_text(r.to_dict(), m))

        df["tenure_text"] = tenure_text
        df["current_head_rate"] = head_rates
        df["current_scoring_fee"] = scoring_fees
        df["plan_text"] = plan_texts

    def apply_filter(self):
        if self.df_all is None or self.df_all.empty:
            self.df_view = pd.DataFrame()
            self.render_table()
            return

        q = (self.search.text() or "").strip().lower()
        if not q:
            self.df_view = self.df_all.copy()
        else:
            cols = [
                "emp_code",
                "first_name",
                "last_name",
                "full_name",
                "display_name",
                "rate_plan",
                "rate_mode",
                "tenure_text",
                "plan_text",
            ]
            cols = [c for c in cols if c in self.df_all.columns]

            mask = None
            for c in cols:
                s = self.df_all[c].astype(str).str.lower().str.contains(q, na=False)
                mask = s if mask is None else (mask | s)

            self.df_view = (
                self.df_all[mask].copy() if mask is not None else self.df_all.copy()
            )

        self.render_table()

    def render_table(self):
        show_cols = [
            ("emp_code", "รหัส"),
            ("full_name", "ชื่อ-นามสกุล"),
            ("start_date", "เริ่มงาน"),
            ("tenure_text", "อายุงาน"),
            ("current_head_rate", "รายหัว/ตอนนี้"),
            ("current_scoring_fee", "Scoring/ตอนนี้"),
            ("plan_text", "แผน"),
            ("rate_mode", "โหมด"),
        ]
        cols = [
            (k, t)
            for k, t in show_cols
            if self.df_view is not None and k in self.df_view.columns
        ]

        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels([t for _, t in cols])

        df = self.df_view if self.df_view is not None else pd.DataFrame()
        self.table.setRowCount(len(df))

        for r in range(len(df)):
            for c, (key, _) in enumerate(cols):
                v = df.iloc[r][key]
                if key == "start_date":
                    v = _fmt_date(v)
                if key in ("current_head_rate", "current_scoring_fee"):
                    v = _fmt_money(v)

                item = QTableWidgetItem(_safe_str(v))
                if key in ("current_head_rate", "current_scoring_fee"):
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

                self.table.setItem(r, c, item)

        self.table.resizeColumnsToContents()
        self.lbl_count.setText(f"แสดง {len(df)} / ทั้งหมด {len(self.df_all)} คน")

    def _selected_row_dict(self) -> dict | None:
        r = self.table.currentRow()
        if r < 0 or self.df_view is None or self.df_view.empty:
            return None
        return self.df_view.iloc[r].to_dict()

    def open_detail(self, *_):
        row_dict = self._selected_row_dict()
        if not row_dict:
            return
        dlg = EmployeeDetailDialog(row_dict, self.as_of, self)
        dlg.exec()

    def open_edit(self, *_):
        row_dict = self._selected_row_dict()
        if not row_dict:
            return
        dlg = EditEmployeeDialog(self.master_dir, row_dict, self)
        if dlg.exec() == QDialog.Accepted:
            self.reload()

    def delete_employee(self, *_):
        row_dict = self._selected_row_dict()
        if not row_dict:
            return

        emp_code = str(row_dict.get("emp_code", "")).strip()
        name = str(row_dict.get("full_name", "")).strip()

        ret = QMessageBox.question(
            self,
            "ยืนยันการลบ",
            f"ต้องการปิดใช้งานพนักงานนี้ใช่ไหม?\n\n{emp_code} - {name}\n\n(ลบแบบปิดใช้งาน active=False)",
            QMessageBox.Yes | QMessageBox.No,
        )
        if ret != QMessageBox.Yes:
            return

        try:
            soft_delete_employee(self.master_dir, emp_code)
            QMessageBox.information(self, "Done", f"✅ ปิดใช้งาน {emp_code} แล้ว")
            self.reload()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
