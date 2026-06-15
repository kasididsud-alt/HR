# src/employee_edit_ui.py
from __future__ import annotations

import pandas as pd
from PySide6.QtCore import QDate
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QFormLayout,
    QLineEdit,
    QComboBox,
    QDateEdit,
    QPushButton,
    QHBoxLayout,
    QMessageBox,
    QTextEdit,
    QLabel,
)

from .employee_store import EmployeeInput
from .employee_update import update_employee_by_code


def _msg_error(parent, text: str):
    QMessageBox.critical(parent, "Error", text)


def _msg_info(parent, text: str):
    QMessageBox.information(parent, "Done", text)


class EditEmployeeDialog(QDialog):
    """
    ฟอร์มแก้ไขพนักงาน (อัปเดตลง master/employees.xlsx)
    - emp_code ล็อค (ห้ามแก้)
    """

    def __init__(self, master_dir: str, emp_row: dict, parent=None):
        super().__init__(parent)
        self.master_dir = master_dir
        self.emp_row = emp_row
        self.emp_code = str(emp_row.get("emp_code", "")).strip()

        self.setWindowTitle(f"✏️ แก้ไขพนักงาน ({self.emp_code})")
        self.resize(540, 480)

        layout = QVBoxLayout(self)
        form = QFormLayout()
        layout.addLayout(form)

        # --- read-only emp_code
        self.emp_code_label = QLabel(self.emp_code)
        self.emp_code_label.setStyleSheet("font-weight:600;")
        form.addRow("รหัสพนักงาน:", self.emp_code_label)

        # --- fields
        self.first_name = QLineEdit()
        self.last_name = QLineEdit()

        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDisplayFormat("yyyy-MM-dd")

        self.display_name = QLineEdit()

        self.note = QTextEdit()
        self.note.setFixedHeight(80)

        # Rate Plan
        self.rate_plan = QComboBox()
        self.rate_plan.addItem("5.1 (<=3=300/2, >3<=6=500/2, >6=800/2)", "5.1")
        self.rate_plan.addItem("5.2 (<=4=300/2, >4<=8=500/2, >8=800/2)", "5.2")
        self.rate_plan.addItem("5.3 (<=4=500/2, >4<=8=800/2, >=9=1000/2)", "5.3")
        self.rate_plan.addItem("5.4 (<=4=500/2, >4<=8=800/2, >=9=1200/2)", "5.4")
        self.rate_plan.addItem("5.5 (Fix Rate / Free)", "5.5")
        self.rate_plan.addItem("5.6 (15 เคสแรกฟรี แล้วเป็น Fix 750/1500)", "5.6")

        # Fix rate selector - ใช้กับ 5.5
        self.fix_rate = QComboBox()
        self.fix_rate.addItem("Free", "")
        for v in [400, 500, 800, 1000, 1200, 1500, 2000, 750]:
            self.fix_rate.addItem(f"fix rate {v}", str(v))
        self.fix_rate.addItem("อื่น... (กรอกเอง)", "__OTHER__")

        # ✅ ช่องกรอกเอง
        self.fix_rate_other = QLineEdit()
        self.fix_rate_other.setPlaceholderText("กรอก Fix Rate เอง เช่น 650")
        self.fix_rate_other.setEnabled(False)

        # Cond15 - ใช้กับ 5.6
        self.cond_after_fix = QComboBox()
        self.cond_after_fix.addItem("fix rate 750 (ประกันสังคม)", "750")
        self.cond_after_fix.addItem("fix rate 1500 (เงินสด)", "1500")

        # Scoring
        self.scoring = QComboBox()
        self.scoring.addItem("free", "FREE")
        self.scoring.addItem("fix rate 200", "200")
        self.scoring.addItem("fix rate 500", "500")

        # layout
        form.addRow("ชื่อ:", self.first_name)
        form.addRow("นามสกุล:", self.last_name)
        form.addRow("วันเริ่มทำงาน:", self.start_date)
        form.addRow("ชื่อแสดง (ไม่กรอกจะใช้ ชื่อ+นามสกุล):", self.display_name)
        form.addRow("แผนเงินรายหัว:", self.rate_plan)
        form.addRow("Fix Rate (ใช้กับแผน 5.5):", self.fix_rate)
        form.addRow("Fix Rate (อื่น):", self.fix_rate_other)  # ✅ เพิ่ม
        form.addRow("แผน 5.6 หลัง 15 เคสแรก:", self.cond_after_fix)
        form.addRow("Scoring Fee:", self.scoring)
        form.addRow("หมายเหตุ:", self.note)

        # buttons
        btn_row = QHBoxLayout()
        layout.addLayout(btn_row)
        self.btn_cancel = QPushButton("ยกเลิก")
        self.btn_save = QPushButton("บันทึกการแก้ไข")
        btn_row.addWidget(self.btn_cancel)
        btn_row.addWidget(self.btn_save)

        self.btn_cancel.clicked.connect(self.reject)
        self.btn_save.clicked.connect(self.on_save)

        self.rate_plan.currentIndexChanged.connect(self._update_visibility)
        self.fix_rate.currentIndexChanged.connect(self._update_visibility)  # ✅

        # เติมค่าจาก emp_row
        self._prefill_from_row()
        self._update_visibility()

    def _set_combo_by_data(self, combo: QComboBox, data_value: str):
        data_value = str(data_value).strip()
        for i in range(combo.count()):
            if str(combo.itemData(i)).strip() == data_value:
                combo.setCurrentIndex(i)
                return

    def _prefill_from_row(self):
        self.first_name.setText(str(self.emp_row.get("first_name", "") or ""))
        self.last_name.setText(str(self.emp_row.get("last_name", "") or ""))
        self.display_name.setText(str(self.emp_row.get("display_name", "") or ""))
        self.note.setText(str(self.emp_row.get("note", "") or ""))

        # start_date
        sd = pd.to_datetime(self.emp_row.get("start_date", None), errors="coerce")
        if pd.isna(sd):
            self.start_date.setDate(QDate.currentDate())
        else:
            self.start_date.setDate(QDate(sd.year, sd.month, sd.day))

        # rate_plan
        self._set_combo_by_data(
            self.rate_plan, str(self.emp_row.get("rate_plan", "") or "")
        )

        # fix_rate (plan 5.5) - รองรับค่าอื่น
        fr = self.emp_row.get("fix_rate", "")
        fr_val = None
        try:
            if fr not in ("", None) and not (isinstance(fr, float) and pd.isna(fr)):
                fr_val = float(fr)
        except Exception:
            fr_val = None

        known = {400, 500, 800, 1000, 1200, 1500, 2000, 750}

        if fr_val is None:
            self._set_combo_by_data(self.fix_rate, "")
            self.fix_rate_other.setText("")
        else:
            if float(fr_val).is_integer() and int(fr_val) in known:
                self._set_combo_by_data(self.fix_rate, str(int(fr_val)))
                self.fix_rate_other.setText("")
            else:
                self._set_combo_by_data(self.fix_rate, "__OTHER__")
                self.fix_rate_other.setText(str(fr_val))

        # cond_after_fix_rate (plan 5.6)
        caf = self.emp_row.get("cond_after_fix_rate", "")
        caf_data = ""
        try:
            if caf not in ("", None) and not (isinstance(caf, float) and pd.isna(caf)):
                caf_data = str(int(float(caf)))
        except Exception:
            caf_data = ""
        self._set_combo_by_data(self.cond_after_fix, caf_data or "750")

        # scoring
        sm = str(self.emp_row.get("scoring_mode", "FREE") or "FREE").strip().upper()
        if sm == "FREE":
            self._set_combo_by_data(self.scoring, "FREE")
        else:
            sf = self.emp_row.get("scoring_fix", "")
            sf_data = "200"
            try:
                if sf not in ("", None) and not (isinstance(sf, float) and pd.isna(sf)):
                    sf_data = str(int(float(sf)))
            except Exception:
                sf_data = "200"
            self._set_combo_by_data(self.scoring, sf_data)

    def _update_visibility(self):
        plan = self.rate_plan.currentData()

        is_fix = plan == "5.5"
        self.fix_rate.setEnabled(is_fix)

        use_other = is_fix and (self.fix_rate.currentData() == "__OTHER__")
        self.fix_rate_other.setEnabled(use_other)

        self.cond_after_fix.setEnabled(plan == "5.6")

    def on_save(self):
        try:
            plan = self.rate_plan.currentData()

            d = self.start_date.date().toPython()
            start_ts = pd.Timestamp(d)

            # scoring
            scoring_val = self.scoring.currentData()
            if scoring_val == "FREE":
                scoring_mode = "FREE"
                scoring_fix = None
            else:
                scoring_mode = "FIX"
                scoring_fix = int(scoring_val)

            # fix rate (plan 5.5) - รองรับอื่น
            fix_rate_val = None
            if plan == "5.5":
                raw = self.fix_rate.currentData()
                if raw == "":
                    fix_rate_val = None
                elif raw == "__OTHER__":
                    txt = (self.fix_rate_other.text() or "").strip()
                    if not txt:
                        raise ValueError("กรุณากรอก Fix Rate (อื่น)")
                    try:
                        fix_rate_val = float(txt)
                    except Exception:
                        raise ValueError("Fix Rate (อื่น) ต้องเป็นตัวเลข เช่น 650")
                else:
                    fix_rate_val = float(raw)

            # cond after fix (plan 5.6)
            cond_after_fix = None
            if plan == "5.6":
                cond_after_fix = int(self.cond_after_fix.currentData())

            inp = EmployeeInput(
                first_name=self.first_name.text(),
                last_name=self.last_name.text(),
                start_date=start_ts,
                rate_plan=plan,
                fix_rate=fix_rate_val,
                cond_free_first_n=15,
                cond_after_fix_rate=cond_after_fix,
                cond_pay_type="",  # update จะ set ตาม 750/1500 ให้อัตโนมัติ
                scoring_mode=scoring_mode,
                scoring_fix=scoring_fix,
                display_name=self.display_name.text(),
                note=self.note.toPlainText(),
            )

            update_employee_by_code(self.master_dir, self.emp_code, inp)
            _msg_info(self, f"✅ แก้ไขพนักงาน {self.emp_code} สำเร็จ")
            self.accept()

        except Exception as e:
            _msg_error(self, str(e))
