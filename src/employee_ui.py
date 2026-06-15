# src/employee_ui.py
from __future__ import annotations

import pandas as pd
from PySide6.QtCore import QDate, Qt
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
    QGroupBox,
)

from .employee_store import EmployeeInput, add_employee
from .styles import get_style_sheet


def _msg_error(parent, text: str):
    QMessageBox.critical(parent, "Error", text)


def _msg_info(parent, text: str):
    QMessageBox.information(parent, "Done", text)


class AddEmployeeDialog(QDialog):
    """
    ฟอร์มเพิ่มพนักงาน (บันทึกลง master/employees.xlsx)
    - ไม่มี Rate Set แล้ว (ใช้ rate_plan อย่างเดียว)
    """

    def __init__(self, master_dir: str, parent=None):
        super().__init__(parent)
        self.master_dir = master_dir
        self.setWindowTitle("➕ เพิ่มพนักงาน")
        self.resize(520, 480)
        self.setStyleSheet(get_style_sheet())

        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)
        
        # Header
        header = QLabel("👤 เพิ่มพนักงานใหม่")
        header.setProperty("class", "header")
        header.setAlignment(Qt.AlignCenter)
        layout.addWidget(header)
        
        # Form container
        form_container = QGroupBox("📝 ข้อมูลพนักงาน")
        form_layout = QFormLayout(form_container)
        form_layout.setSpacing(10)
        layout.addWidget(form_container)

        # fields
        self.first_name = QLineEdit()
        self.first_name.setPlaceholderText("กรอกชื่อจริง")
        self.first_name.setMinimumHeight(32)
        
        self.last_name = QLineEdit()
        self.last_name.setPlaceholderText("กรอกนามสกุล")
        self.last_name.setMinimumHeight(32)

        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDisplayFormat("yyyy-MM-dd")
        self.start_date.setDate(QDate.currentDate())
        self.start_date.setMinimumHeight(32)

        self.display_name = QLineEdit()
        self.display_name.setPlaceholderText("ชื่อที่แสดง (ไม่กรอกจะใช้ ชื่อ+นามสกุล)")
        self.display_name.setMinimumHeight(32)

        self.note = QTextEdit()
        self.note.setPlaceholderText("หมายเหตุเพิ่มเติม (ถ้ามี)")
        self.note.setFixedHeight(80)

        # Rate Plan
        self.rate_plan = QComboBox()
        self.rate_plan.addItem("📋 เลือกแผนเงินรายหัว...", "")
        self.rate_plan.addItem("5.1 (<=3=300/2, >3<=6=500/2, >6=800/2)", "5.1")
        self.rate_plan.addItem("5.2 (<=4=300/2, >4<=8=500/2, >8=800/2)", "5.2")
        self.rate_plan.addItem("5.3 (<=4=500/2, >4<=8=800/2, >=9=1000/2)", "5.3")
        self.rate_plan.addItem("5.4 (<=4=500/2, >4<=8=800/2, >=9=1200/2)", "5.4")
        self.rate_plan.addItem("5.5 (Fix Rate / Free)", "5.5")
        self.rate_plan.addItem("5.6 (15 เคสแรกฟรี แล้วเป็น Fix 750/1500)", "5.6")
        self.rate_plan.setMinimumHeight(32)

        # Fix rate selector - ใช้กับ 5.5
        self.fix_rate = QComboBox()
        self.fix_rate.addItem("🔄 เลือก Fix Rate...", "__SELECT__")
        self.fix_rate.addItem("Free", "")
        for v in [400, 500, 800, 1000, 1200, 1500, 2000, 750]:
            self.fix_rate.addItem(f"💰 fix rate {v}", str(v))
        self.fix_rate.addItem("✏️ อื่น... (กรอกเอง)", "__OTHER__")
        self.fix_rate.setMinimumHeight(32)

        # ✅ ช่องกรอกเองสำหรับ Fix rate (อื่น)
        self.fix_rate_other = QLineEdit()
        self.fix_rate_other.setPlaceholderText("กรอก Fix Rate เอง เช่น 650")
        self.fix_rate_other.setEnabled(False)
        self.fix_rate_other.setMinimumHeight(32)

        # Cond15 - ใช้กับ 5.6
        self.cond_after_fix = QComboBox()
        self.cond_after_fix.addItem("🔄 เลือกอัตราหลัง 15 เคส...", "")
        self.cond_after_fix.addItem("💰 fix rate 750 (ประกันสังคม)", "750")
        self.cond_after_fix.addItem("💰 fix rate 1500 (เงินสด)", "1500")
        self.cond_after_fix.setMinimumHeight(32)

        # Scoring
        self.scoring = QComboBox()
        self.scoring.addItem("🔄 เลือก Scoring Rate...", "")
        self.scoring.addItem("🆓 free", "FREE")
        self.scoring.addItem("💰 fix rate 200", "200")
        self.scoring.addItem("💰 fix rate 500", "500")
        self.scoring.setMinimumHeight(32)

        # layout form
        form_layout.addRow("👤 ชื่อ:", self.first_name)
        form_layout.addRow("👤 นามสกุล:", self.last_name)
        form_layout.addRow("📅 วันเริ่มทำงาน:", self.start_date)
        form_layout.addRow("📝 ชื่อแสดง (ไม่กรอกจะใช้ ชื่อ+นามสกุล):", self.display_name)
        form_layout.addRow("💰 แผนเงินรายหัว:", self.rate_plan)
        form_layout.addRow("💵 Fix Rate (ใช้กับแผน 5.5):", self.fix_rate)
        form_layout.addRow("✏️ Fix Rate (อื่น):", self.fix_rate_other)  # ✅ เพิ่ม
        form_layout.addRow("📈 แผน 5.6 หลัง 15 เคสแรก:", self.cond_after_fix)
        form_layout.addRow("🎯 Scoring:", self.scoring)
        form_layout.addRow("📝 หมายเหตุ:", self.note)

        # buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        layout.addLayout(btn_row)
        
        self.btn_cancel = QPushButton("❌ ยกเลิก")
        self.btn_cancel.setProperty("class", "secondary")
        self.btn_cancel.setMinimumHeight(36)
        
        self.btn_save = QPushButton("✅ บันทึกพนักงาน")
        self.btn_save.setProperty("class", "success")
        self.btn_save.setMinimumHeight(36)
        
        btn_row.addStretch()
        btn_row.addWidget(self.btn_cancel)
        btn_row.addWidget(self.btn_save)

        self.btn_cancel.clicked.connect(self.reject)
        self.btn_save.clicked.connect(self.on_save)

        # toggle
        self.rate_plan.currentIndexChanged.connect(self._update_visibility)
        self.fix_rate.currentIndexChanged.connect(
            self._update_visibility
        )  # ✅ ให้สลับอื่นได้
        self._update_visibility()

    def _update_visibility(self):
        plan = self.rate_plan.currentData()

        is_fix = plan == "5.5"
        self.fix_rate.setEnabled(is_fix)

        # ✅ เปิดช่องกรอกเองเมื่อเลือก "อื่น..."
        use_other = is_fix and (self.fix_rate.currentData() == "__OTHER__")
        self.fix_rate_other.setEnabled(use_other)

        is_cond = plan == "5.6"
        self.cond_after_fix.setEnabled(is_cond)

    def on_save(self):
        try:
            plan = self.rate_plan.currentData()
            if not plan:
                raise ValueError("กรุณาเลือกแผนเงินรายหัว")

            d = self.start_date.date().toPython()  # datetime.date
            start_ts = pd.Timestamp(d)

            # scoring
            scoring_val = self.scoring.currentData()
            if not scoring_val:
                raise ValueError("กรุณาเลือก Scoring")
            if scoring_val == "FREE":
                scoring_mode = "FREE"
                scoring_fix = None
            else:
                scoring_mode = "FIX"
                scoring_fix = int(scoring_val)

            # fix rate (plan 5.5)
            fix_rate_val = None
            if plan == "5.5":
                raw = self.fix_rate.currentData()
                if raw == "__SELECT__":
                    raise ValueError("กรุณาเลือก Fix Rate หรือ Free")
                elif raw == "":
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
                raw_cond = self.cond_after_fix.currentData()
                if not raw_cond:
                    raise ValueError("กรุณาเลือกอัตราหลัง 15 เคสแรก")
                cond_after_fix = int(raw_cond)

            inp = EmployeeInput(
                first_name=self.first_name.text(),
                last_name=self.last_name.text(),
                start_date=start_ts,
                rate_plan=plan,
                fix_rate=fix_rate_val,
                cond_free_first_n=15,
                cond_after_fix_rate=cond_after_fix,
                cond_pay_type="",  # store จะ set ให้เองตาม 750/1500
                scoring_mode=scoring_mode,
                scoring_fix=scoring_fix,
                display_name=self.display_name.text(),
                note=self.note.toPlainText(),
            )

            emp_code = add_employee(self.master_dir, inp)
            _msg_info(self, f"✅ บันทึกพนักงานสำเร็จ\nรหัสพนักงาน: {emp_code}")
            self.accept()

        except Exception as e:
            _msg_error(self, str(e))
