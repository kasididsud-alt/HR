# app.py
import sys
import os
import importlib.util
import ctypes


def _maybe_relaunch_from_project_venv() -> None:
    if getattr(sys, "frozen", False):
        return

    base_dir = os.path.dirname(os.path.abspath(__file__))
    current_python = os.path.realpath(sys.executable)
    candidates = [
        os.path.join(base_dir, ".venv311", "bin", "python"),
        os.path.join(base_dir, ".venv", "bin", "python"),
        os.path.join(base_dir, "venv", "bin", "python"),
        os.path.join(base_dir, ".venv", "Scripts", "python.exe"),
        os.path.join(base_dir, "venv", "Scripts", "python.exe"),
    ]

    required_modules = ("pandas", "PySide6", "openpyxl")
    missing_dependency = any(
        importlib.util.find_spec(module) is None for module in required_modules
    )
    old_python = sys.version_info < (3, 10)

    if not old_python and not missing_dependency:
        return

    for candidate in candidates:
        if os.path.exists(candidate) and os.path.realpath(candidate) != current_python:
            os.execv(candidate, [candidate, *sys.argv])

    if old_python:
        raise RuntimeError(
            "This app requires Python 3.10 or newer. "
            "Run it with .venv311/bin/python app.py."
        )


_maybe_relaunch_from_project_venv()

import pandas as pd

from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QBoxLayout,
    QLabel,
    QPushButton,
    QLineEdit,
    QDateEdit,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QGroupBox,
    QFormLayout,
    QComboBox,
    QScrollArea,
    QFrame,
    QStatusBar,
    QProgressBar,
    QTabWidget,
    QListWidget,
    QListWidgetItem,
    QAbstractItemView,
    QHeaderView,
    QSizePolicy,
)

from src.excel_io import to_excel_bytes

from src.employee_ui import AddEmployeeDialog
from src.employee_list_ui import EmployeeListDialog

from src.master_loader import load_employees, load_aliases, build_lookups

from src.case_import import list_sheet_names as list_case_sheets, load_cases_xlsx
from src.hospital_import import (
    list_sheet_names as list_hosp_sheets,
    load_hospital_fee_map,
)

from src.fee_calc import compute_case_fees, has_scoring_doctor_prefix, normalize_doctor_name
from src.alias_audit import find_unknown_names

from src.payroll_detail_ui import PayrollDetailDialog
from src.doctor_detail_ui import DoctorDetailDialog
from src.styles import get_style_sheet, apply_light_app_palette, apply_modern_style
from src.app_runtime import asset_path, default_master_dir

import traceback


def _user_friendly_error_message(msg: str) -> str:
    text = str(msg or "").strip()
    low = text.lower()

    if not text:
        return "เกิดข้อผิดพลาดที่ระบบยังอธิบายไม่ได้ กรุณาลองใหม่อีกครั้ง"
    if "traceback" in low:
        return "ระบบทำงานไม่สำเร็จ กรุณาลองใหม่อีกครั้ง หรือแจ้งผู้ดูแลระบบ"
    if "permission denied" in low or "winerror 32" in low:
        return "ไม่สามารถบันทึกไฟล์ได้ เพราะไฟล์ปลายทางกำลังถูกเปิดใช้งานอยู่ใน Excel หรือโปรแกรมอื่น"
    if "no such file" in low or "filenotfounderror" in low:
        return "ไม่พบไฟล์ที่ต้องใช้ กรุณาตรวจสอบว่าเลือกไฟล์และโฟลเดอร์ถูกต้อง"
    if "ชีทนี้ขาดคอลัมน์" in text:
        return text + "\n\nกรุณาตรวจสอบหัวคอลัมน์ในไฟล์ Excel ให้ตรงกับรูปแบบที่โปรแกรมต้องใช้"
    if "ไม่พบข้อมูลโรงพยาบาล" in text:
        return text + "\n\nกรุณาเพิ่มข้อมูลโรงพยาบาลและอัตราค่าตอบแทนก่อนคำนวณ"
    if "ไม่สามารถอ่านข้อมูลโรงพยาบาล" in text:
        return text + "\n\nกรุณาตรวจสอบไฟล์ hospitals.json ในโฟลเดอร์ master"
    return text


def show_error(parent, title: str, msg: str, details: str | None = None):
    box = QMessageBox(parent)
    box.setIcon(QMessageBox.Critical)
    box.setWindowTitle(title)
    box.setText(_user_friendly_error_message(msg))
    if details:
        box.setDetailedText(details)
    box.exec()


def show_info(parent, title: str, msg: str):
    QMessageBox.information(parent, title, msg)


def _load_app_icon() -> QIcon:
    for icon_name in ("salary_calc.ico", "salary_calc.png"):
        path = asset_path(icon_name)
        if os.path.exists(path):
            icon = QIcon(path)
            if not icon.isNull():
                return icon
    return QIcon()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("💼 Salary & Fee Calculator")
        self.resize(1180, 900)
        self.setMinimumSize(640, 480)
        icon = _load_app_icon()
        if not icon.isNull():
            self.setWindowIcon(icon)
        
        # Apply modern styling
        self.setStyleSheet(get_style_sheet())

        self.case_file_path = ""

        # ✅ keep case result for joining in detail dialog
        self.case_out = pd.DataFrame()

        self.pay_details = pd.DataFrame()
        self.emp_summary_all = pd.DataFrame()
        self.emp_summary_view = pd.DataFrame()

        self.doctor_details = pd.DataFrame()
        self.doctor_summary_all = pd.DataFrame()
        self.doctor_summary_view = pd.DataFrame()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        self.setCentralWidget(scroll)

        root = QWidget()
        root.setContentsMargins(15, 15, 15, 15)
        scroll.setWidget(root)

        main_layout = QVBoxLayout(root)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # -------------------------
        # Header
        # -------------------------
        header_wrap = QWidget()
        header_layout = QVBoxLayout(header_wrap)
        header_layout.setContentsMargins(0, 0, 0, 6)
        header_layout.setSpacing(2)

        header_label = QLabel("Salary & Fee Calculator")
        header_label.setProperty("class", "header")
        header_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        header_layout.addWidget(header_label)

        subtitle_label = QLabel("จัดการไฟล์เคส ตรวจสอบข้อมูลหลัก และสรุปผลค่าตอบแทนได้ในหน้าจอเดียว")
        subtitle_label.setProperty("class", "helper")
        subtitle_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        header_layout.addWidget(subtitle_label)
        main_layout.addWidget(header_wrap)
        
        # -------------------------
        # Top Workspace
        # -------------------------
        top_workspace = QWidget()
        top_workspace_layout = QHBoxLayout(top_workspace)
        top_workspace_layout.setContentsMargins(0, 0, 0, 0)
        top_workspace_layout.setSpacing(14)
        self.top_workspace = top_workspace
        self.top_workspace_layout = top_workspace_layout

        case_box = QGroupBox("ไฟล์เคส")
        case_layout = QVBoxLayout(case_box)
        case_layout.setSpacing(8)

        self.case_edit = QLineEdit()
        self.case_edit.setPlaceholderText("เลือกไฟล์เคส .xlsx")
        self.case_edit.setReadOnly(True)
        self.case_edit.setMinimumHeight(36)

        btn_case = QPushButton("เลือกไฟล์เคส")
        btn_case.clicked.connect(self.browse_case_file)
        btn_case.setMinimumHeight(36)

        sheet_label = QLabel("ชีท")
        self.case_sheet_combo = QComboBox()
        self.case_sheet_combo.addItem("(ยังไม่ได้เลือกไฟล์)", "")
        self.case_sheet_combo.setEnabled(False)
        self.case_sheet_combo.setMinimumHeight(32)

        case_layout.addWidget(self.case_edit)
        case_layout.addWidget(btn_case)
        case_layout.addWidget(sheet_label)
        case_layout.addWidget(self.case_sheet_combo)
        config_box = QGroupBox("การตั้งค่า")
        config_layout = QFormLayout(config_box)
        config_layout.setSpacing(10)

        date_label = QLabel("วันที่อ้างอิง")
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setMinimumHeight(32)
        config_layout.addRow(date_label, self.date_edit)

        master_label = QLabel("โฟลเดอร์ master")
        self.master_dir_edit = QLineEdit()
        self.master_dir_edit.setText(default_master_dir())
        self.master_dir_edit.setMinimumHeight(32)
        config_layout.addRow(master_label, self.master_dir_edit)

        master_btn_row = QWidget()
        master_btn_layout = QHBoxLayout(master_btn_row)
        master_btn_layout.setContentsMargins(0, 0, 0, 0)
        master_btn_layout.setSpacing(8)

        btn_browse_master = QPushButton("เลือกโฟลเดอร์")
        btn_browse_master.clicked.connect(self.browse_master_folder)
        btn_browse_master.setProperty("class", "secondary")
        btn_browse_master.setMinimumHeight(32)

        btn_open_master = QPushButton("เปิดโฟลเดอร์")
        btn_open_master.clicked.connect(self.open_master_folder)
        btn_open_master.setProperty("class", "secondary")
        btn_open_master.setMinimumHeight(32)

        master_btn_layout.addWidget(btn_browse_master)
        master_btn_layout.addWidget(btn_open_master)
        config_layout.addRow("", master_btn_row)

        top_workspace_layout.addWidget(case_box, 3)
        top_workspace_layout.addWidget(config_box, 2)
        main_layout.addWidget(top_workspace)

        # -------------------------
        # Action Buttons
        # -------------------------
        actions_box = QGroupBox("เมนูลัด")
        actions_layout = QGridLayout(actions_box)
        actions_layout.setHorizontalSpacing(10)
        actions_layout.setVerticalSpacing(10)
        self.actions_layout = actions_layout
        
        self.btn_add_emp = QPushButton("เพิ่มพนักงาน")
        self.btn_add_emp.clicked.connect(self.on_add_employee)
        self.btn_add_emp.setProperty("class", "success")
        self.btn_add_emp.setMinimumHeight(40)

        self.btn_view_emp = QPushButton("จัดการพนักงาน")
        self.btn_view_emp.clicked.connect(self.on_view_employees)
        self.btn_view_emp.setProperty("class", "secondary")
        self.btn_view_emp.setMinimumHeight(40)

        self.btn_summary_emp = QPushButton("สรุปพนักงาน")
        self.btn_summary_emp.clicked.connect(self.on_summary_employees)
        self.btn_summary_emp.setProperty("class", "secondary")
        self.btn_summary_emp.setMinimumHeight(40)

        self.btn_free_emp = QPushButton("ไม่รับเงิน")
        self.btn_free_emp.clicked.connect(self.on_free_employees)
        self.btn_free_emp.setProperty("class", "secondary")
        self.btn_free_emp.setMinimumHeight(40)

        self.btn_free_physician = QPushButton("ไม่คิด Physician")
        self.btn_free_physician.clicked.connect(self.on_free_physicians)
        self.btn_free_physician.setProperty("class", "secondary")
        self.btn_free_physician.setMinimumHeight(40)

        self.btn_hospital = QPushButton("จัดการโรงพยาบาล")
        self.btn_hospital.clicked.connect(self.on_manage_hospitals)
        self.btn_hospital.setProperty("class", "secondary")
        self.btn_hospital.setMinimumHeight(40)

        self.btn_calc = QPushButton("คำนวณและส่งออกไฟล์")
        self.btn_calc.clicked.connect(self.on_calculate)
        self.btn_calc.setProperty("class", "success")
        self.btn_calc.setMinimumHeight(48)

        for btn in (
            self.btn_add_emp,
            self.btn_view_emp,
            self.btn_summary_emp,
            self.btn_free_emp,
            self.btn_free_physician,
            self.btn_hospital,
        ):
            btn.setMinimumHeight(42)
            btn.setMinimumWidth(110)
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.btn_calc.setMinimumWidth(150)
        self.btn_calc.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.action_buttons = [
            self.btn_add_emp,
            self.btn_view_emp,
            self.btn_summary_emp,
            self.btn_free_emp,
            self.btn_free_physician,
            self.btn_hospital,
        ]

        self._relayout_action_buttons(compact=False)
        main_layout.addWidget(actions_box)

        # -------------------------
        # Results
        # -------------------------
        results_header = QLabel("สรุปผลการคำนวณ")
        results_header.setProperty("class", "header")
        main_layout.addWidget(results_header)

        results_hint = QLabel("ดับเบิลคลิกที่แถวเพื่อเปิดรายละเอียดของข้อมูลที่เลือก")
        results_hint.setProperty("class", "helper")
        main_layout.addWidget(results_hint)

        stats_container = QWidget()
        stats_layout = QGridLayout(stats_container)
        stats_layout.setContentsMargins(0, 0, 0, 0)
        stats_layout.setHorizontalSpacing(10)
        stats_layout.setVerticalSpacing(10)
        self.stats_layout = stats_layout

        (
            self.metric_mode_card,
            self.metric_mode_value,
            self.metric_mode_label,
        ) = self._create_metric_card("มุมมอง")
        (
            self.metric_rows_card,
            self.metric_rows_value,
            self.metric_rows_label,
        ) = self._create_metric_card("รายการ")
        (
            self.metric_monitor_card,
            self.metric_monitor_value,
            self.metric_monitor_label,
        ) = self._create_metric_card("Monitor")
        (
            self.metric_scoring_card,
            self.metric_scoring_value,
            self.metric_scoring_label,
        ) = self._create_metric_card("Scoring")
        (
            self.metric_total_card,
            self.metric_total_value,
            self.metric_total_label,
        ) = self._create_metric_card("ยอดรวม")

        self.metric_cards = [
            self.metric_mode_card,
            self.metric_rows_card,
            self.metric_monitor_card,
            self.metric_scoring_card,
            self.metric_total_card,
        ]
        self._relayout_metric_cards(compact=False)
        main_layout.addWidget(stats_container)

        # Filter Controls
        filter_container = QWidget()
        filter_layout = QHBoxLayout(filter_container)
        filter_layout.setSpacing(10)
        
        mode_label = QLabel("มุมมอง")
        filter_layout.addWidget(mode_label)
        
        self.view_mode = QComboBox()
        self.view_mode.addItem("พนักงาน", "EMP")
        self.view_mode.addItem("แพทย์", "DOC")
        self.view_mode.currentIndexChanged.connect(self.apply_summary_filter)
        self.view_mode.setMinimumHeight(32)
        self.view_mode.setMinimumWidth(120)
        filter_layout.addWidget(self.view_mode)
        
        filter_layout.addSpacing(20)
        
        search_label = QLabel("ค้นหา")
        filter_layout.addWidget(search_label)
        
        self.summary_search = QLineEdit()
        self.summary_search.setPlaceholderText("ค้นหาชื่อ/รหัส...")
        self.summary_search.textChanged.connect(self.apply_summary_filter)
        self.summary_search.setMinimumHeight(32)
        filter_layout.addWidget(self.summary_search, 1)
        
        self.btn_clear_search = QPushButton("ล้าง")
        self.btn_clear_search.clicked.connect(lambda: self.summary_search.setText(""))
        self.btn_clear_search.setProperty("class", "secondary")
        self.btn_clear_search.setMinimumHeight(32)
        self.btn_clear_search.setMinimumWidth(96)
        self.btn_clear_search.setMaximumWidth(96)
        filter_layout.addWidget(self.btn_clear_search)
        
        main_layout.addWidget(filter_container)

        # Results Table
        self.table = QTableWidget(0, 4)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.cellDoubleClicked.connect(self.open_detail_by_mode)
        self.table.setMinimumHeight(400)
        self.table.setAlternatingRowColors(True)
        main_layout.addWidget(self.table, 1)

        # Status
        status_container = QWidget()
        status_layout = QHBoxLayout(status_container)
        status_layout.setContentsMargins(0, 8, 0, 0)
        
        self.status = QLabel("พร้อมใช้งาน")
        self.status.setProperty("class", "status")
        status_layout.addWidget(self.status)
        
        status_layout.addStretch()
        
        # Progress bar
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setMaximumHeight(6)
        self.progress.setTextVisible(False)
        self.progress.setMaximumWidth(220)
        self.progress.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        status_layout.addWidget(self.progress)
        
        main_layout.addWidget(status_container)

        self.loading_overlay = QFrame(self.centralWidget())
        self.loading_overlay.setVisible(False)
        self.loading_overlay.setStyleSheet("background-color: rgba(15, 23, 42, 0.20);")

        loading_layout = QVBoxLayout(self.loading_overlay)
        loading_layout.setContentsMargins(24, 24, 24, 24)
        loading_layout.addStretch(1)

        self.loading_card = QFrame()
        self.loading_card.setStyleSheet(
            "background-color: #ffffff;"
            "border: 1px solid #dbe4f0;"
            "border-radius: 8px;"
        )
        self.loading_card.setMinimumWidth(320)
        self.loading_card.setMaximumWidth(420)

        loading_card_layout = QVBoxLayout(self.loading_card)
        loading_card_layout.setContentsMargins(22, 20, 22, 20)
        loading_card_layout.setSpacing(10)

        self.loading_title = QLabel("กำลังคำนวณ...")
        self.loading_title.setStyleSheet(
            "font-size: 18px; font-weight: 700; color: #0f172a;"
        )
        loading_card_layout.addWidget(self.loading_title)

        self.loading_detail = QLabel("กำลังเตรียมข้อมูล")
        self.loading_detail.setWordWrap(True)
        self.loading_detail.setStyleSheet(
            "font-size: 13px; color: #475569; font-weight: 500;"
        )
        loading_card_layout.addWidget(self.loading_detail)

        self.loading_bar = QProgressBar()
        self.loading_bar.setRange(0, 0)
        self.loading_bar.setTextVisible(False)
        self.loading_bar.setFixedHeight(8)
        loading_card_layout.addWidget(self.loading_bar)

        loading_layout.addWidget(self.loading_card, 0, Qt.AlignCenter)
        loading_layout.addStretch(1)

        self._sync_loading_overlay_geometry()

        self._setup_table_headers_for_mode("EMP")
        self._update_summary_metrics(pd.DataFrame(), mode="EMP")
        apply_modern_style(self)
        self._apply_responsive_layout()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._apply_responsive_layout()
        self._sync_loading_overlay_geometry()

    def _sync_loading_overlay_geometry(self):
        if not hasattr(self, "loading_overlay"):
            return
        if self.loading_overlay is None or self.centralWidget() is None:
            return
        self.loading_overlay.setGeometry(self.centralWidget().rect())

    def _set_loading_visible(self, visible: bool, detail: str = ""):
        if visible:
            self._sync_loading_overlay_geometry()
            self.loading_title.setText("กำลังคำนวณ...")
            self.loading_detail.setText(detail or "กำลังเตรียมข้อมูล")
            self.loading_overlay.raise_()
            self.loading_overlay.show()
            self.btn_calc.setEnabled(False)
            self.progress.setVisible(True)
            self.progress.setRange(0, 0)
            QApplication.processEvents()
            return

        self.loading_overlay.hide()
        self.btn_calc.setEnabled(True)
        self.progress.setVisible(False)
        self.progress.setRange(0, 100)

    def _update_loading_detail(self, detail: str):
        if not self.loading_overlay.isVisible():
            return
        self.loading_detail.setText(detail)
        QApplication.processEvents()

    # -------------------------
    # File Pickers
    # -------------------------
    def browse_case_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "เลือกไฟล์เคส", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return

        self.case_file_path = path
        self.case_edit.setText(path)

        try:
            sheets = list_case_sheets(path)
            self.case_sheet_combo.clear()
            for s in sheets:
                self.case_sheet_combo.addItem(str(s), s)
            self.case_sheet_combo.setEnabled(True)
            self.status.setText("✅ เลือกไฟล์เคสแล้ว")
            self.status.setProperty("class", "status-success")
        except Exception as e:
            show_error(self, "Error", f"อ่านรายชื่อชีทไม่สำเร็จ: {e}")
            self.case_sheet_combo.clear()
            self.case_sheet_combo.addItem("(อ่านชีทไม่สำเร็จ)", "")
            self.case_sheet_combo.setEnabled(False)

    def browse_master_folder(self):
        current_dir = self.master_dir_edit.text().strip() or default_master_dir()
        path = QFileDialog.getExistingDirectory(
            self, "เลือกโฟลเดอร์ master", current_dir
        )
        if not path:
            return

        self.master_dir_edit.setText(path)
        self.status.setText("✅ เปลี่ยนโฟลเดอร์ master แล้ว")
        self.status.setProperty("class", "status-success")

    def open_master_folder(self):
        master_dir = self.master_dir_edit.text().strip()
        if not master_dir:
            show_error(self, "Error", "กรุณาระบุโฟลเดอร์ master ก่อน")
            return
        os.makedirs(master_dir, exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(master_dir)
            elif sys.platform == "darwin":
                os.system(f'open "{master_dir}"')
            else:
                os.system(f'xdg-open "{master_dir}"')
        except Exception as e:
            show_error(self, "Error", f"เปิดโฟลเดอร์ไม่สำเร็จ: {e}")

    # -------------------------
    # Master loading
    # -------------------------
    def _load_master(self):
        master_dir = self.master_dir_edit.text().strip()
        if not master_dir:
            raise RuntimeError("กรุณาระบุโฟลเดอร์ master ก่อน")
        os.makedirs(master_dir, exist_ok=True)

        # ✅ load_employees() คืนค่าเป็น DataFrame ตัวเดียว
        employees = load_employees(master_dir)

        try:
            aliases = load_aliases(master_dir)
        except Exception:
            aliases = pd.DataFrame(
                columns=["alias_name", "emp_code", "source_col", "active"]
            )

        emp_map, alias_map = build_lookups(employees, aliases)

        # ✅ คืน 3 ค่า
        return employees, emp_map, alias_map

    def _export_unknown_names_report(self, unknown_df: pd.DataFrame) -> str:
        master_dir = self.master_dir_edit.text().strip() or default_master_dir()
        os.makedirs(master_dir, exist_ok=True)

        report_path = os.path.join(master_dir, "aliases_to_review.xlsx")
        report_df = unknown_df.copy()
        report_df["emp_code"] = ""
        report_df["หมายเหตุ"] = "กรอก emp_code ให้ตรงกับพนักงาน แล้วนำข้อมูลนี้ไปเพิ่มใน aliases.xlsx"

        out_bytes = to_excel_bytes(
            {"UnknownNames": report_df},
            required_blank_columns={"UnknownNames": ["emp_code"]},
        )
        with open(report_path, "wb") as f:
            f.write(out_bytes)
        return report_path

    def _build_case_result_highlights(
        self, case_df: pd.DataFrame, unknown_df: pd.DataFrame | None
    ) -> list[dict]:
        if case_df is None or case_df.empty:
            return []

        highlights: list[dict] = []
        unknown_by_col: dict[str, set[str]] = {}

        if unknown_df is not None and not unknown_df.empty:
            for _, row in unknown_df.iterrows():
                source_col = str(row.get("source_col", "") or "").strip()
                alias_name = str(row.get("alias_name", "") or "").strip()
                if source_col and alias_name:
                    unknown_by_col.setdefault(source_col, set()).add(alias_name)

        row_warned: set[int] = set()
        monitor_cols = [
            ("Monitor Tech", "Monitor Fee"),
            ("Monitor Tech (2)", "Monitor Fee2"),
        ]

        for row_pos, (_, row) in enumerate(case_df.iterrows()):
            for source_col, fee_col in monitor_cols:
                value = str(row.get(source_col, "") or "").strip()
                if not value or value not in unknown_by_col.get(source_col, set()):
                    continue

                if row_pos not in row_warned:
                    highlights.append(
                        {
                            "row": row_pos,
                            "entire_row": True,
                            "fill": "warn",
                        }
                    )
                    row_warned.add(row_pos)

                highlights.append(
                    {
                        "row": row_pos,
                        "columns": [source_col, fee_col],
                        "fill": "error",
                        "comment": (
                            "ชื่อนี้ยังจับคู่กับพนักงานไม่ได้ "
                            "จึงยังไม่คำนวณ Monitor Fee สำหรับช่องนี้"
                        ),
                    }
                )

            if "วันที่ตรวจ" in case_df.columns and pd.isna(row.get("วันที่ตรวจ")):
                highlights.append(
                    {
                        "row": row_pos,
                        "columns": ["วันที่ตรวจ"],
                        "fill": "error",
                        "comment": "วันที่ตรวจว่างหรืออ่านวันที่ไม่ได้",
                    }
                )

        return highlights

    def _show_unknown_names_error(self, unknown_df: pd.DataFrame):
        preview = []
        for _, row in unknown_df.head(12).iterrows():
            preview.append(f"- {row.get('source_col', '')}: {row.get('alias_name', '')}")

        more_count = max(len(unknown_df) - len(preview), 0)
        report_path = self._export_unknown_names_report(unknown_df)

        msg = (
            "พบชื่อในไฟล์เคสที่ยังจับคู่กับพนักงานไม่ได้\n\n"
            "กรุณาตรวจสอบและเพิ่ม alias ก่อนคำนวณ:\n"
            + "\n".join(preview)
        )
        if more_count:
            msg += f"\n- ... และอีก {more_count} รายการ"
        msg += f"\n\nระบบสร้างไฟล์รายการไว้ให้แล้ว:\n{report_path}"

        show_error(self, "ชื่อในไฟล์เคสไม่ตรง", msg)

    def _read_cases_df(self):
        if not self.case_file_path or not os.path.exists(self.case_file_path):
            raise RuntimeError("กรุณาเลือกไฟล์เคส (.xlsx) ก่อน")

        sheet = (
            self.case_sheet_combo.currentData()
            if self.case_sheet_combo.isEnabled()
            else None
        )
        if not sheet:
            sheet = 0
        return load_cases_xlsx(self.case_file_path, sheet_name=sheet)

    def _read_hospital_fee_map(self):
        """อ่านข้อมูลโรงพยาบาลจากไฟล์ JSON ที่เพิ่มไว้"""
        master_dir = self.master_dir_edit.text().strip()
        if not master_dir:
            raise RuntimeError("กรุณาระบุโฟลเดอร์ master ก่อน")
        
        hospitals_file = os.path.join(master_dir, "hospitals.json")
        if not os.path.exists(hospitals_file):
            raise RuntimeError("ไม่พบข้อมูลโรงพยาบาล กรุณาเพิ่มโรงพยาบาลก่อน")
        
        try:
            import json
            with open(hospitals_file, 'r', encoding='utf-8') as f:
                hospitals = json.load(f)
            
            # สร้าง hospital fee map จากข้อมูล JSON
            hospital_fee_map = {}
            for hospital in hospitals:
                hospital_name = hospital['name']
                types = hospital['types']
                
                # สร้าง map สำหรับทุก type
                for type_info in types:
                    type_name = type_info['type']
                    rate = type_info['rate']
                    
                    # สนับสนุนทั้งชื่อ type และตัวเลข type
                    # 1. (hospital, type_name) - เช่น ("โรงพยาบาล A", "OPD")
                    hospital_fee_map[(hospital_name, type_name)] = rate
                    
                    # 2. (hospital, type_number) - เช่น ("โรงพยาบาล A", "1")
                    if type_name in ["1", "2", "3", "ไม่จ่าย"]:
                        hospital_fee_map[(hospital_name, type_name)] = rate
                    
                    # 3. (hospital, int_type) - เช่น ("โรงพยาบาล A", 1)
                    try:
                        type_int = int(type_name)
                        hospital_fee_map[(hospital_name, type_int)] = rate
                    except ValueError:
                        pass  # ไม่ใช่ตัวเลข ข้าม
            
            return hospital_fee_map
            
        except Exception as e:
            raise RuntimeError(f"ไม่สามารถอ่านข้อมูลโรงพยาบาล: {e}")

    # -------------------------
    # UI Actions
    # -------------------------
    def on_add_employee(self):
        try:
            master_dir = self.master_dir_edit.text().strip()
            if not master_dir:
                show_error(self, "Error", "กรุณาระบุโฟลเดอร์ master ก่อน")
                return
            os.makedirs(master_dir, exist_ok=True)
            dlg = AddEmployeeDialog(master_dir, self)
            dlg.exec()
            self.status.setText("✅ เพิ่มพนักงานแล้ว")
            self.status.setProperty("class", "status-success")
        except Exception as e:
            show_error(self, "Error", str(e), details=traceback.format_exc())
            self.status.setText("❌ เกิดข้อผิดพลาด")
            self.status.setProperty("class", "status-error")

    def on_view_employees(self):
        try:
            master_dir = self.master_dir_edit.text().strip()
            if not master_dir:
                show_error(self, "Error", "กรุณาระบุโฟลเดอร์ master ก่อน")
                return
            os.makedirs(master_dir, exist_ok=True)

            dlg = EmployeeListDialog(master_dir, self)
            dlg.exec()
            self.status.setText("✅ เปิดรายการพนักงานแล้ว")
            self.status.setProperty("class", "status-success")
        except Exception as e:
            show_error(self, "Error", str(e), details=traceback.format_exc())
            self.status.setText("❌ เกิดข้อผิดพลาด")
            self.status.setProperty("class", "status-error")

    def on_summary_employees(self):
        try:
            master_dir = self.master_dir_edit.text().strip()
            if not master_dir:
                show_error(self, "Error", "กรุณาระบุโฟลเดอร์ master ก่อน")
                return
            os.makedirs(master_dir, exist_ok=True)

            # Load employees data
            employees, emp_map, alias_map = self._load_master()
            
            if employees.empty:
                show_info(self, "ข้อมูลพนักงาน", "ไม่มีข้อมูลพนักงานในระบบ")
                return

            # Create summary dialog
            from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QTextEdit, QPushButton
            
            summary_dlg = QDialog(self)
            summary_dlg.setWindowTitle("📊 สรุปข้อมูลพนักงาน")
            summary_dlg.resize(600, 500)
            summary_dlg.setStyleSheet(get_style_sheet())
            
            layout = QVBoxLayout(summary_dlg)
            
            # Header
            header = QLabel("📊 สรุปข้อมูลพนักงานทั้งหมด")
            header.setProperty("class", "header")
            header.setAlignment(Qt.AlignCenter)
            layout.addWidget(header)
            
            # Summary text
            summary_text = QTextEdit()
            summary_text.setReadOnly(True)
            
            # Generate summary
            summary_content = f"📋 จำนวนพนักงานทั้งหมด: {len(employees)} คน\n\n"
            
            # Group by rate plan
            if 'rate_plan' in employees.columns:
                rate_summary = employees['rate_plan'].value_counts().to_dict()
                summary_content += "📊 สรุปตามแผนเงินเดือน:\n"
                for plan, count in rate_summary.items():
                    summary_content += f"  • แผน {plan}: {count} คน\n"
                summary_content += "\n"
            
            # Active employees
            if 'start_date' in employees.columns:
                as_of = pd.Timestamp.today().normalize()
                active_emp = employees[employees['start_date'] <= as_of]
                summary_content += f"✅ พนักงานที่ทำงานอยู่ (ณ วันนี้ {as_of.strftime('%Y-%m-%d')}): {len(active_emp)} คน\n"
                summary_content += f"🔴 พนักงานที่ยังไม่เริ่มทำงาน: {len(employees) - len(active_emp)} คน\n\n"
            
            # List all employees
            summary_content += "👥 รายชื่อพนักงาน:\n"
            for idx, emp in employees.iterrows():
                emp_code = emp.get('emp_code', 'N/A')
                first_name = emp.get('first_name', '')
                last_name = emp.get('last_name', '')
                display_name = emp.get('display_name', f"{first_name} {last_name}".strip())
                rate_plan = emp.get('rate_plan', 'N/A')
                start_date = emp.get('start_date', '')
                
                if pd.notna(start_date):
                    start_str = pd.to_datetime(start_date).strftime('%Y-%m-%d')
                else:
                    start_str = 'N/A'
                
                summary_content += f"\n  📝 {emp_code}: {display_name}"
                summary_content += f"\n     📅 เริ่มทำงาน: {start_str}"
                summary_content += f"\n     💰 แผน: {rate_plan}"
            
            summary_text.setPlainText(summary_content)
            layout.addWidget(summary_text)
            
            # Close button
            close_btn = QPushButton("❌ ปิด")
            close_btn.setProperty("class", "secondary")
            close_btn.clicked.connect(summary_dlg.accept)
            layout.addWidget(close_btn)
            
            summary_dlg.exec()
            self.status.setText("✅ แสดงสรุปพนักงานแล้ว")
            self.status.setProperty("class", "status-success")
            
        except Exception as e:
            show_error(self, "Error", str(e), details=traceback.format_exc())
            self.status.setText("❌ เกิดข้อผิดพลาด")
            self.status.setProperty("class", "status-error")

    def on_free_employees(self):
        try:
            master_dir = self.master_dir_edit.text().strip()
            if not master_dir:
                show_error(self, "Error", "กรุณาระบุโฟลเดอร์ master ก่อน")
                return
            os.makedirs(master_dir, exist_ok=True)

            # Create main dialog
            from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout, QTabWidget
            
            main_dlg = QDialog(self)
            main_dlg.setWindowTitle("🆓 จัดการผู้ไม่รับเงิน")
            main_dlg.resize(700, 600)
            main_dlg.setStyleSheet(get_style_sheet())
            
            layout = QVBoxLayout(main_dlg)
            
            # Header
            header = QLabel("🆓 จัดการผู้ไม่รับเงิน")
            header.setProperty("class", "header")
            header.setAlignment(Qt.AlignCenter)
            layout.addWidget(header)
            
            # Create tabs
            tab_widget = QTabWidget()
            
            # Tab 1: พนักงานที่ไม่รับเงิน
            emp_tab = self.create_free_employees_tab(master_dir)
            tab_widget.addTab(emp_tab, "👥 พนักงาน")
            
            # Tab 2: แพทย์ที่ไม่รับเงิน
            doctor_tab = self.create_free_doctors_tab(master_dir)
            tab_widget.addTab(doctor_tab, "🩺 แพทย์")

            # เปิดแท็บกรอกชื่อแพทย์อัตโนมัติ (ใช้งานเร็วขึ้น)
            tab_widget.setCurrentIndex(1)
            
            layout.addWidget(tab_widget)
            
            # Close button
            close_btn = QPushButton("❌ ปิด")
            close_btn.setProperty("class", "secondary")
            close_btn.clicked.connect(main_dlg.accept)
            layout.addWidget(close_btn)
            
            main_dlg.exec()
            self.status.setText("✅ เปิดจัดการผู้ไม่รับเงินแล้ว")
            self.status.setProperty("class", "status-success")
            
        except Exception as e:
            show_error(self, "Error", str(e), details=traceback.format_exc())
            self.status.setText("❌ เกิดข้อผิดพลาด")
            self.status.setProperty("class", "status-error")

    def create_free_employees_tab(self, master_dir):
        """สร้าง Tab พนักงานที่ไม่รับเงิน"""
        from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QTextEdit, QHBoxLayout, QPushButton
        
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Load employees data
        employees, emp_map, alias_map = self._load_master()
        
        # Filter free employees
        free_employees = employees[
            (employees['rate_plan'] == '5.5') & 
            (employees['fix_rate'].isna() | (employees['fix_rate'] == ''))
        ]
        
        # Summary info
        info_label = QLabel(f"📋 พนักงานที่ไม่รับเงิน: {len(free_employees)} คน")
        info_label.setStyleSheet("color: #10b981; font-size: 16px; padding: 10px; background-color: #dcfce7; border-radius: 6px; margin: 10px;")
        layout.addWidget(info_label)
        
        # Free employees list
        free_text = QTextEdit()
        free_text.setReadOnly(True)
        
        if free_employees.empty:
            free_content = "🎉 ไม่มีพนักงานที่ทำงานแบบไม่รับเงินในระบบ\n\n"
            free_content += "📝 หมายเหตุ: พนักงานที่ไม่รับเงินคือพนักงานที่:\n"
            free_content += "   • อยู่ในแผน 5.5 (Fix Rate / Free)\n"
            free_content += "   • ไม่มีการกำหนด Fix Rate (ว่างหรือเป็นค่าว่าง)\n"
        else:
            free_content = f"👥 รายชื่อพนักงานที่ไม่รับเงิน ({len(free_employees)} คน):\n\n"
            
            for idx, emp in free_employees.iterrows():
                emp_code = emp.get('emp_code', 'N/A')
                first_name = emp.get('first_name', '')
                last_name = emp.get('last_name', '')
                display_name = emp.get('display_name', f"{first_name} {last_name}".strip())
                start_date = emp.get('start_date', '')
                
                if pd.notna(start_date):
                    start_str = pd.to_datetime(start_date).strftime('%Y-%m-%d')
                else:
                    start_str = 'N/A'
                
                free_content += f"  📝 {emp_code}: {display_name}\n"
                free_content += f"     📅 เริ่มทำงาน: {start_str}\n"
                free_content += f"     💰 แผน: 5.5 (Free)\n"
                free_content += f"     🆓 สถานะ: ไม่รับค่าตอบแทน\n\n"
        
        free_text.setPlainText(free_content)
        layout.addWidget(free_text)
        
        # Export button
        btn_layout = QHBoxLayout()
        export_btn = QPushButton("📄 Export Excel")
        export_btn.setProperty("class", "success")
        export_btn.clicked.connect(lambda: self.export_free_employees(free_employees))
        btn_layout.addWidget(export_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        return tab

    def create_free_doctors_tab(self, master_dir):
        """สร้าง Tab แพทย์ที่ไม่รับเงิน"""
        from PySide6.QtWidgets import (
            QWidget,
            QVBoxLayout,
            QLabel,
            QTextEdit,
            QHBoxLayout,
            QPushButton,
            QGroupBox,
            QFormLayout,
        )

        tab = QWidget()
        layout = QVBoxLayout(tab)

        info_label = QLabel(
            "🩺 จัดการรายชื่อแพทย์ที่ไม่รับเงินตามวันที่\n"
            "กรอกชื่อแพทย์ได้ทีละบรรทัด หรือคั่นด้วยเครื่องหมายจุลภาค (,)"
        )
        info_label.setStyleSheet(
            "color: #2563eb; font-size: 14px; padding: 10px; "
            "background-color: #dbeafe; border-radius: 6px; margin: 10px 0;"
        )
        layout.addWidget(info_label)

        form_group = QGroupBox("📋 ข้อมูลแพทย์ที่ไม่รับเงิน")
        form_layout = QFormLayout(form_group)

        self.free_date_edit = QDateEdit()
        self.free_date_edit.setCalendarPopup(True)
        self.free_date_edit.setDate(QDate.currentDate())
        self.free_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.free_date_edit.setMinimumHeight(32)
        form_layout.addRow("📅 วันที่:", self.free_date_edit)

        self.free_doctors_text = QTextEdit()
        self.free_doctors_text.setPlaceholderText(
            "ตัวอย่าง:\n"
            "นพ.สมชาย ใจดี\n"
            "นพ.กิตติ แพทย์ดี, พญ.นภา สุขใจ"
        )
        self.free_doctors_text.setMinimumHeight(220)
        form_layout.addRow("👨‍⚕️ รายชื่อแพทย์:", self.free_doctors_text)

        layout.addWidget(form_group)

        btn_layout = QHBoxLayout()

        load_btn = QPushButton("📂 โหลด")
        load_btn.setProperty("class", "secondary")
        load_btn.clicked.connect(lambda: self.load_free_doctors_simple(master_dir))
        btn_layout.addWidget(load_btn)

        save_btn = QPushButton("💾 บันทึก")
        save_btn.setProperty("class", "success")
        save_btn.clicked.connect(lambda: self.save_free_doctors_simple(master_dir))
        btn_layout.addWidget(save_btn)

        clear_btn = QPushButton("🧹 ล้าง")
        clear_btn.setProperty("class", "secondary")
        clear_btn.clicked.connect(self.clear_free_doctors_form)
        btn_layout.addWidget(clear_btn)

        manage_btn = QPushButton("✏️ ดู/ลบข้อมูลเดิม")
        manage_btn.setProperty("class", "secondary")
        manage_btn.clicked.connect(self.edit_free_doctors)
        btn_layout.addWidget(manage_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.free_doctors_status = QLabel("📄 เลือกวันที่แล้วกดโหลด หรือกรอกรายชื่อแล้วกดบันทึก")
        self.free_doctors_status.setStyleSheet(
            "padding: 8px; background-color: #f8fafc; border-radius: 6px; margin: 10px 0;"
        )
        layout.addWidget(self.free_doctors_status)

        # Load existing data for the selected date when opening the tab.
        self.load_free_doctors_simple(master_dir)
        self.free_date_edit.dateChanged.connect(lambda _date: self.load_free_doctors_simple(master_dir))

        return tab

    def edit_free_doctors(self):
        """แก้ไขข้อมูลแพทย์ที่ไม่รับเงิน"""
        try:
            # สร้างโฟลเดอร์ถ้าไม่มี
            master_dir = self.master_dir_edit.text().strip()
            if not master_dir:
                show_error(self, "Error", "กรุณาระบุโฟลเดอร์ master ก่อน")
                return
            
            os.makedirs(master_dir, exist_ok=True)
            
            # ค้นหาไฟล์แพทย์ฟรีทั้งหมด
            import glob
            free_doctor_files = glob.glob(os.path.join(master_dir, "free_doctors_*.txt"))
            
            # สร้าง Dialog แก้ไข
            from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout, QTabWidget
            
            edit_dlg = QDialog(self)
            edit_dlg.setWindowTitle("✏️ แก้ไขข้อมูลแพทย์ที่ไม่รับเงิน")
            edit_dlg.resize(700, 600)
            edit_dlg.setStyleSheet(get_style_sheet())
            
            layout = QVBoxLayout(edit_dlg)
            
            # Header
            header = QLabel("✏️ แก้ไขข้อมูลแพทย์ที่ไม่รับเงิน")
            header.setProperty("class", "header")
            header.setAlignment(Qt.AlignCenter)
            layout.addWidget(header)
            
            # Create tabs
            tab_widget = QTabWidget()
            
            # Tab 1: ดูข้อมูลเดิม
            view_tab = self.create_edit_view_tab(master_dir, free_doctor_files)
            tab_widget.addTab(view_tab, "📄 ดูข้อมูลเดิม")
            
            # Tab 2: ลบข้อมูล
            delete_tab = self.create_edit_delete_tab(master_dir, free_doctor_files)
            tab_widget.addTab(delete_tab, "🗑️ ลบข้อมูล")
            
            layout.addWidget(tab_widget)
            
            # Close button
            close_btn = QPushButton("❌ ปิด")
            close_btn.setProperty("class", "secondary")
            close_btn.clicked.connect(edit_dlg.accept)
            layout.addWidget(close_btn)
            
            edit_dlg.exec()
            
        except Exception as e:
            show_error(self, "Error", str(e), details=traceback.format_exc())

    def create_edit_view_tab(self, master_dir, files):
        """สร้าง Tab ดูข้อมูลเดิม"""
        from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QTextEdit, QScrollArea
        
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        if not files:
            no_data = QLabel("📄 ไม่มีข้อมูลแพทย์ที่ไม่รับเงิน")
            no_data.setStyleSheet("color: #64748b; font-size: 16px; padding: 20px; text-align: center;")
            layout.addWidget(no_data)
            return tab
        
        # Create scrollable text area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        
        for file_path in sorted(files):
            file_name = os.path.basename(file_path)
            
            # File header
            file_label = QLabel(f"📁 {file_name}")
            file_label.setStyleSheet("color: #2563eb; font-size: 14px; font-weight: 600; padding: 8px; background-color: #dbeafe; border-radius: 4px; margin: 5px 0;")
            content_layout.addWidget(file_label)
            
            # File content
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                content_text = QTextEdit()
                content_text.setPlainText(content)
                content_text.setReadOnly(True)
                content_text.setMaximumHeight(150)
                content_text.setStyleSheet("font-family: monospace; font-size: 12px;")
                content_layout.addWidget(content_text)
                
            except Exception as e:
                error_label = QLabel(f"❌ ไม่สามารถอ่านไฟล์: {e}")
                error_label.setStyleSheet("color: #ef4444; padding: 5px;")
                content_layout.addWidget(error_label)
            
            # Separator
            separator = QLabel("")
            separator.setStyleSheet("border-bottom: 1px solid #e2e8f0; margin: 10px 0;")
            content_layout.addWidget(separator)
        
        scroll.setWidget(content_widget)
        layout.addWidget(scroll)
        
        return tab

    def create_edit_delete_tab(self, master_dir, files):
        """สร้าง Tab ลบข้อมูล"""
        from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QListWidget, QListWidgetItem, QPushButton, QHBoxLayout
        
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        if not files:
            no_data = QLabel("📄 ไม่มีข้อมูลให้ลบ")
            no_data.setStyleSheet("color: #64748b; font-size: 16px; padding: 20px; text-align: center;")
            layout.addWidget(no_data)
            return tab
        
        # Instructions
        info_label = QLabel("🗑️ เลือกไฟล์ที่ต้องการลบ:")
        info_label.setStyleSheet("color: #dc2626; font-size: 16px; padding: 10px; background-color: #fee2e2; border-radius: 6px; margin: 10px;")
        layout.addWidget(info_label)
        
        # File list
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.MultiSelection)
        
        for file_path in sorted(files):
            file_name = os.path.basename(file_path)
            item = QListWidgetItem(file_name)
            item.setData(1, file_path)  # Store full path
            self.file_list.addItem(item)
        
        layout.addWidget(self.file_list)
        
        # Buttons
        btn_layout = QHBoxLayout()
        
        select_all_btn = QPushButton("✅ เลือกทั้งหมด")
        select_all_btn.setProperty("class", "secondary")
        select_all_btn.clicked.connect(self.select_all_files)
        
        clear_btn = QPushButton("🔄 ล้างการเลือก")
        clear_btn.setProperty("class", "secondary")
        clear_btn.clicked.connect(self.file_list.clearSelection)
        
        delete_btn = QPushButton("🗑️ ลบไฟล์")
        delete_btn.setProperty("class", "danger")
        delete_btn.clicked.connect(self.delete_selected_files)
        
        btn_layout.addWidget(select_all_btn)
        btn_layout.addWidget(clear_btn)
        btn_layout.addWidget(delete_btn)
        layout.addLayout(btn_layout)
        
        return tab

    def select_all_files(self):
        """เลือกไฟล์ทั้งหมด"""
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            item.setSelected(True)

    def delete_selected_files(self):
        """ลบไฟล์ที่เลือก"""
        try:
            selected_items = self.file_list.selectedItems()
            
            if not selected_items:
                show_info(self, "ลบไฟล์", "กรุณาเลือกไฟล์อย่างน้อย 1 ไฟล์")
                return
            
            # สร้างรายชื่อไฟล์ที่จะลบ
            file_names = [item.text() for item in selected_items]
            files_to_delete = [item.data(1) for item in selected_items]
            
            # แสดง Dialog ยืนยันการลบ
            from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout
            
            confirm_dlg = QDialog(self)
            confirm_dlg.setWindowTitle("🗑️ ยืนยันการลบไฟล์")
            confirm_dlg.resize(500, 300)
            confirm_dlg.setStyleSheet(get_style_sheet())
            
            layout = QVBoxLayout(confirm_dlg)
            
            # Header
            header = QLabel("🗑️ ยืนยันการลบไฟล์")
            header.setProperty("class", "header")
            header.setAlignment(Qt.AlignCenter)
            layout.addWidget(header)
            
            # Warning
            warning_label = QLabel(f"⚠️ คุณกำลังจะลบ {len(selected_items)} ไฟล์:\n\n" + "\n".join([f"📄 {name}" for name in file_names]) + "\n\n❗️ ข้อมูลจะถูกลบถาววิ!")
            warning_label.setStyleSheet("color: #dc2626; font-size: 14px; padding: 15px; background-color: #fee2e2; border-radius: 6px;")
            layout.addWidget(warning_label)
            
            # Buttons
            btn_layout = QHBoxLayout()
            
            # Delete button
            delete_btn = QPushButton("🗑️ ลบ")
            delete_btn.setProperty("class", "danger")
            delete_btn.clicked.connect(lambda: self.perform_file_deletion(files_to_delete, confirm_dlg))
            
            # Cancel button
            cancel_btn = QPushButton("❌ ยกเลิก")
            cancel_btn.setProperty("class", "secondary")
            cancel_btn.clicked.connect(confirm_dlg.reject)
            
            btn_layout.addWidget(delete_btn)
            btn_layout.addWidget(cancel_btn)
            layout.addLayout(btn_layout)
            
            confirm_dlg.exec()
            
        except Exception as e:
            show_error(self, "ลบไฟล์ Error", str(e))

    def perform_file_deletion(self, files_to_delete, confirm_dlg):
        """ดำเนินการลบไฟล์จริง"""
        try:
            deleted_count = 0
            error_files = []
            
            for file_path in files_to_delete:
                try:
                    os.remove(file_path)
                    deleted_count += 1
                except Exception as e:
                    error_files.append(f"{os.path.basename(file_path)}: {e}")
            
            # ปิด dialog ยืนยัน
            confirm_dlg.accept()
            
            # แสดงผลลัพธ์
            if error_files:
                error_msg = f"⚠️ ลบไฟล์สำเร็จ {deleted_count} ไฟล์\n\n❌ ลบไม่ได้:\n" + "\n".join(error_files)
                show_info(self, "ผลการลบ", error_msg)
            else:
                show_info(self, "ลบสำเร็จ", f"✅ ลบไฟล์สำเร็จ {deleted_count} ไฟล์")
            
            self.status.setText(f"✅ ลบข้อมูลแพทย์ฟรี {deleted_count} ไฟล์แล้ว")
            self.status.setProperty("class", "status-success")
            
        except Exception as e:
            show_error(self, "ลบไฟล์ Error", str(e))

    def on_manage_hospitals(self):
        """จัดการข้อมูลโรงพยาบาล"""
        try:
            # สร้างโฟลเดอร์ถ้าไม่มี
            master_dir = self.master_dir_edit.text().strip()
            if not master_dir:
                show_error(self, "Error", "กรุณาระบุโฟลเดอร์ master ก่อน")
                return
            
            os.makedirs(master_dir, exist_ok=True)
            
            # สร้าง Dialog จัดการโรงพยาบาล
            from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout, QTabWidget, QWidget, QFormLayout, QLineEdit, QComboBox, QMessageBox
            
            hospital_dlg = QDialog(self)
            hospital_dlg.setWindowTitle("🏥 จัดการข้อมูลโรงพยาบาล")
            hospital_dlg.resize(600, 500)
            hospital_dlg.setStyleSheet(get_style_sheet())
            
            layout = QVBoxLayout(hospital_dlg)
            
            # Header
            header = QLabel("🏥 จัดการข้อมูลโรงพยาบาล")
            header.setProperty("class", "header")
            header.setAlignment(Qt.AlignCenter)
            layout.addWidget(header)
            
            # Create tabs
            tab_widget = QTabWidget()
            self.hospital_tab_widget = tab_widget
            self.hospital_edit_index = None
            
            # Tab 1: เพิ่มโรงพยาบาล
            add_tab = self.create_hospital_add_tab(master_dir)
            tab_widget.addTab(add_tab, "➕ เพิ่มโรงพยาบาล")
            
            # Tab 2: ดู/ลบโรงพยาบาล
            manage_tab = self.create_hospital_manage_tab(master_dir)
            tab_widget.addTab(manage_tab, "📋 จัดการโรงพยาบาล")
            
            layout.addWidget(tab_widget)
            
            # Close button
            close_btn = QPushButton("❌ ปิด")
            close_btn.setProperty("class", "secondary")
            close_btn.clicked.connect(hospital_dlg.accept)
            layout.addWidget(close_btn)
            
            hospital_dlg.exec()
            
        except Exception as e:
            show_error(self, "Error", str(e), details=traceback.format_exc())

    def create_hospital_add_tab(self, master_dir):
        """สร้าง Tab เพิ่มโรงพยาบาล"""
        from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QFormLayout, QLineEdit, QComboBox, QPushButton, QHBoxLayout, QScrollArea
        
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Form for adding hospital
        form_container = QGroupBox("📝 เพิ่มโรงพยาบาลใหม่")
        form_layout = QFormLayout(form_container)
        
        # Hospital name
        self.hospital_name_edit = QLineEdit()
        self.hospital_name_edit.setPlaceholderText("กรอกชื่อโรงพยาบาล...")
        self.hospital_name_edit.setMinimumHeight(32)
        form_layout.addRow("🏥 ชื่อโรงพยาบาล:", self.hospital_name_edit)
        
        # Type and Rate container
        types_container = QGroupBox("📋 ประเภทและราคา")
        types_layout = QVBoxLayout(types_container)
        
        # Type 1
        type1_layout = QHBoxLayout()
        type1_layout.addWidget(QLabel("📋 Type:"))
        self.type1_combo = QComboBox()
        self.type1_combo.setEditable(True)
        self.type1_combo.addItems(["1", "2", "3", "ไม่จ่าย"])
        self.type1_combo.setMinimumHeight(32)
        self.type1_combo.setMinimumWidth(150)
        type1_layout.addWidget(self.type1_combo)
        
        type1_layout.addWidget(QLabel("💰 ราคา:"))
        self.rate1_edit = QLineEdit()
        self.rate1_edit.setPlaceholderText("ราคา...")
        self.rate1_edit.setMinimumHeight(32)
        self.rate1_edit.setMinimumWidth(100)
        type1_layout.addWidget(self.rate1_edit)
        type1_layout.addStretch()
        types_layout.addLayout(type1_layout)
        
        # Type 2
        type2_layout = QHBoxLayout()
        type2_layout.addWidget(QLabel("📋 Type:"))
        self.type2_combo = QComboBox()
        self.type2_combo.setEditable(True)
        self.type2_combo.addItems(["1", "2", "3", "ไม่จ่าย"])
        self.type2_combo.setMinimumHeight(32)
        self.type2_combo.setMinimumWidth(150)
        type2_layout.addWidget(self.type2_combo)
        
        type2_layout.addWidget(QLabel("💰 ราคา:"))
        self.rate2_edit = QLineEdit()
        self.rate2_edit.setPlaceholderText("ราคา...")
        self.rate2_edit.setMinimumHeight(32)
        self.rate2_edit.setMinimumWidth(100)
        type2_layout.addWidget(self.rate2_edit)
        type2_layout.addStretch()
        types_layout.addLayout(type2_layout)
        
        # Type 3
        type3_layout = QHBoxLayout()
        type3_layout.addWidget(QLabel("📋 Type:"))
        self.type3_combo = QComboBox()
        self.type3_combo.setEditable(True)
        self.type3_combo.addItems(["1", "2", "3", "ไม่จ่าย"])
        self.type3_combo.setMinimumHeight(32)
        self.type3_combo.setMinimumWidth(150)
        type3_layout.addWidget(self.type3_combo)
        
        type3_layout.addWidget(QLabel("💰 ราคา:"))
        self.rate3_edit = QLineEdit()
        self.rate3_edit.setPlaceholderText("ราคา...")
        self.rate3_edit.setMinimumHeight(32)
        self.rate3_edit.setMinimumWidth(100)
        type3_layout.addWidget(self.rate3_edit)
        type3_layout.addStretch()
        types_layout.addLayout(type3_layout)
        
        layout.addWidget(form_container)
        layout.addWidget(types_container)
        
        # Buttons
        btn_layout = QHBoxLayout()
        
        self.hospital_save_btn = QPushButton("💾 บันทึกโรงพยาบาล")
        self.hospital_save_btn.setProperty("class", "success")
        self.hospital_save_btn.clicked.connect(lambda: self.save_hospital(master_dir))
        
        clear_btn = QPushButton("🔄 ล้างฟอร์ม")
        clear_btn.setProperty("class", "secondary")
        clear_btn.clicked.connect(self.clear_hospital_form)
        
        btn_layout.addWidget(self.hospital_save_btn)
        btn_layout.addWidget(clear_btn)
        layout.addLayout(btn_layout)
        
        return tab

    def add_type_field(self):
        """เพิ่มช่องกรอก Type ใหม่"""
        from PySide6.QtWidgets import QHBoxLayout, QLineEdit, QPushButton, QComboBox, QLabel
        
        type_layout = QHBoxLayout()
        
        # Type combo
        type_combo = QComboBox()
        type_combo.setEditable(True)
        type_combo.addItems(["OPD", "IPD", "ER", "ห้องปฏิบัติการ", "ห้องซักซ้อน", "อื่นๆ"])
        type_combo.setMinimumHeight(32)
        type_combo.setMinimumWidth(150)
        
        # Rate edit
        rate_edit = QLineEdit()
        rate_edit.setPlaceholderText("ราคา...")
        rate_edit.setMinimumHeight(32)
        rate_edit.setMinimumWidth(100)
        
        # Remove button
        remove_btn = QPushButton("❌")
        remove_btn.setProperty("class", "danger")
        remove_btn.setMaximumWidth(40)
        remove_btn.clicked.connect(lambda: self.remove_type_field(type_layout))
        
        type_layout.addWidget(QLabel("📋 Type:"))
        type_layout.addWidget(type_combo)
        type_layout.addWidget(QLabel("💰 ราคา:"))
        type_layout.addWidget(rate_edit)
        type_layout.addWidget(remove_btn)
        type_layout.addStretch()
        
        self.types_layout.addLayout(type_layout)
        
        # Store references
        self.type_fields.append({
            'layout': type_layout,
            'type_combo': type_combo,
            'rate_edit': rate_edit,
            'remove_btn': remove_btn
        })

    def remove_type_field(self, layout_to_remove):
        """ลบช่องกรอก Type"""
        # Remove layout
        while layout_to_remove.count():
            item = layout_to_remove.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        # Remove from types layout
        self.types_layout.removeItem(layout_to_remove)
        
        # Remove from type_fields
        self.type_fields = [tf for tf in self.type_fields if tf['layout'] != layout_to_remove]

    def clear_hospital_form(self):
        """ล้างฟอร์มโรงพยาบาล"""
        self.hospital_edit_index = None
        self.hospital_name_edit.clear()
        if hasattr(self, 'hospital_save_btn'):
            self.hospital_save_btn.setText("💾 บันทึกโรงพยาบาล")
        
        # Clear Type 1, 2, 3
        if hasattr(self, 'type1_combo'):
            self.type1_combo.setCurrentIndex(0)
            self.rate1_edit.clear()
        
        if hasattr(self, 'type2_combo'):
            self.type2_combo.setCurrentIndex(0)
            self.rate2_edit.clear()
        
        if hasattr(self, 'type3_combo'):
            self.type3_combo.setCurrentIndex(0)
            self.rate3_edit.clear()

    def save_hospital(self, master_dir):
        """บันทึกข้อมูลโรงพยาบาล"""
        try:
            hospital_name = self.hospital_name_edit.text().strip()
            if not hospital_name:
                show_info(self, "บันทึก", "กรุณากรอกชื่อโรงพยาบาล")
                return
            
            # สร้างข้อมูลโรงพยาบาล
            hospital_data = {
                'name': hospital_name,
                'types': []
            }
            
            # เก็บข้อมูล Type และ Rate (1, 2, 3)
            type_widgets = [
                (self.type1_combo, self.rate1_edit),
                (self.type2_combo, self.rate2_edit),
                (self.type3_combo, self.rate3_edit)
            ]
            
            for i, (type_combo, rate_edit) in enumerate(type_widgets, 1):
                type_value = type_combo.currentText().strip()
                rate_value = rate_edit.text().strip()
                
                if type_value:
                    # ถ้าเป็น "ไม่จ่าย" ให้ใช้อัตราค่าเป็น 0
                    if type_value == "ไม่จ่าย":
                        hospital_data['types'].append({
                            'type': type_value,
                            'rate': 0.0
                        })
                    # ถ้ามีราคา ให้ตรวจสอบว่าเป็นตัวเลข
                    elif rate_value:
                        try:
                            rate_float = float(rate_value)
                            hospital_data['types'].append({
                                'type': type_value,
                                'rate': rate_float
                            })
                        except ValueError:
                            show_info(self, "บันทึก", f"ราคา Type {i} ไม่ถูกต้อง")
                            return
            
            if not hospital_data['types']:
                show_info(self, "บันทึก", "กรุณากรอกข้อมูล Type และ Rate อย่างน้อย 1 ชุด")
                return
            
            # บันทึกลงไฟล์
            hospitals_file = os.path.join(master_dir, "hospitals.json")
            
            # อ่านข้อมูลเดิม
            hospitals = []
            if os.path.exists(hospitals_file):
                try:
                    import json
                    with open(hospitals_file, 'r', encoding='utf-8') as f:
                        hospitals = json.load(f)
                except:
                    hospitals = []
            
            edit_index = getattr(self, 'hospital_edit_index', None)
            is_editing = edit_index is not None
            if is_editing:
                if not (0 <= int(edit_index) < len(hospitals)):
                    show_info(self, "แก้ไข", "ไม่พบข้อมูลโรงพยาบาลเดิม กรุณารีเฟรชรายการแล้วลองใหม่")
                    return
                hospitals[int(edit_index)] = hospital_data
            else:
                hospitals.append(hospital_data)
            
            # บันทึกลงไฟล์
            import json
            with open(hospitals_file, 'w', encoding='utf-8') as f:
                json.dump(hospitals, f, ensure_ascii=False, indent=2)
            
            # ล้างฟอร์ม
            self.clear_hospital_form()
            
            # รีเฟรชรายการ
            self.load_hospital_list(master_dir)
            
            action_text = "แก้ไข" if is_editing else "บันทึก"
            show_info(self, f"{action_text}สำเร็จ", f"✅ {action_text}โรงพยาบาล '{hospital_name}'\n📊 จำนวนประเภท: {len(hospital_data['types'])}")
            
        except Exception as e:
            show_error(self, "บันทึก Error", str(e))

    def create_hospital_manage_tab(self, master_dir):
        """สร้าง Tab จัดการโรงพยาบาล"""
        from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QPushButton, QHBoxLayout
        
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Instructions
        info_label = QLabel("📋 รายชื่อโรงพยาบาลที่บันทึกไว้:")
        info_label.setStyleSheet("color: #2563eb; font-size: 16px; padding: 10px; background-color: #dbeafe; border-radius: 6px; margin: 10px;")
        layout.addWidget(info_label)
        
        # Hospital table
        self.hospital_list = QTableWidget(0, 3)
        self.hospital_list.setHorizontalHeaderLabels(["#", "ชื่อโรงพยาบาล", "ประเภท / ราคา"])
        self.hospital_list.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.hospital_list.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.hospital_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.hospital_list.itemDoubleClicked.connect(lambda _item: self.edit_selected_hospital(master_dir))
        header = self.hospital_list.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        self.load_hospital_list(master_dir)
        layout.addWidget(self.hospital_list)
        
        # Buttons
        btn_layout = QHBoxLayout()
        
        refresh_btn = QPushButton("🔄 รีเฟรช")
        refresh_btn.setProperty("class", "secondary")
        refresh_btn.clicked.connect(lambda: self.load_hospital_list(master_dir))

        edit_btn = QPushButton("✏️ แก้ไขที่เลือก")
        edit_btn.setProperty("class", "secondary")
        edit_btn.clicked.connect(lambda: self.edit_selected_hospital(master_dir))
        
        delete_btn = QPushButton("🗑️ ลบที่เลือก")
        delete_btn.setProperty("class", "danger")
        delete_btn.clicked.connect(lambda: self.delete_selected_hospitals(master_dir))
        
        btn_layout.addWidget(refresh_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        layout.addLayout(btn_layout)
        
        return tab

    def save_free_doctors_simple(self, master_dir):
        """บันทึกชื่อแพทย์ที่ไม่รับเงินแบบง่าย"""
        try:
            # Get date and doctor names
            selected_date = self.free_date_edit.date().toString("yyyy-MM-dd")
            doctors_text = self.free_doctors_text.toPlainText().strip()
            
            if not doctors_text:
                show_info(self, "บันทึก", "กรุณากรอกชื่อแพทย์อย่างน้อย 1 คน")
                return
            
            # Parse doctor names (split by comma or new line)
            doctor_names = []
            for line in doctors_text.split('\n'):
                line = line.strip()
                if line:
                    # Split by comma
                    for name in line.split(','):
                        name = name.strip()
                        if name:
                            doctor_names.append(name)
            
            if not doctor_names:
                show_info(self, "บันทึก", "กรุณากรอกชื่อแพทย์อย่างน้อย 1 คน")
                return
            
            # Create file path
            os.makedirs(master_dir, exist_ok=True)
            file_path = os.path.join(master_dir, f"free_doctors_{selected_date}.txt")
            
            # Save to file
            with open(file_path, 'w', encoding='utf-8') as f:
                for doctor_name in doctor_names:
                    f.write(f"{doctor_name}\n")
            
            # Update status
            self.free_doctors_status.setText(f"✅ บันทึกแล้ว {len(doctor_names)} คน วันที่ {selected_date}")
            self.free_doctors_status.setStyleSheet("padding: 8px; background-color: #dcfce7; color: #16a34a; border-radius: 6px; margin: 10px 0;")
            
            show_info(self, "บันทึกสำเร็จ", f"✅ บันทึกแพทย์ที่ไม่รับเงิน {len(doctor_names)} คน\n📅 วันที่: {selected_date}")
            
        except Exception as e:
            show_error(self, "บันทึก Error", str(e))
            self.free_doctors_status.setText(f"❌ บันทึกไม่สำเร็จ: {str(e)}")
            self.free_doctors_status.setStyleSheet("padding: 8px; background-color: #fee2e2; color: #dc2626; border-radius: 6px; margin: 10px 0;")

    def load_free_doctors_simple(self, master_dir):
        """โหลดข้อมูลแพทย์ที่ไม่รับเงินแบบง่าย"""
        try:
            selected_date = self.free_date_edit.date().toString("yyyy-MM-dd")
            file_path = os.path.join(master_dir, f"free_doctors_{selected_date}.txt")
            
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                
                if content:
                    self.free_doctors_text.setPlainText(content)
                    doctor_count = len([line for line in content.split('\n') if line.strip()])
                    self.free_doctors_status.setText(f"📂 โหลดแล้ว {doctor_count} คน วันที่ {selected_date}")
                    self.free_doctors_status.setStyleSheet("padding: 8px; background-color: #dbeafe; color: #2563eb; border-radius: 6px; margin: 10px 0;")
                else:
                    self.free_doctors_text.clear()
                    self.free_doctors_status.setText(f"📄 ไม่มีข้อมูล วันที่ {selected_date}")
                    self.free_doctors_status.setStyleSheet("padding: 8px; background-color: #f8fafc; border-radius: 6px; margin: 10px 0;")
            else:
                self.free_doctors_text.clear()
                self.free_doctors_status.setText(f"📄 ไม่มีข้อมูล วันที่ {selected_date}")
                self.free_doctors_status.setStyleSheet("padding: 8px; background-color: #f8fafc; border-radius: 6px; margin: 10px 0;")
                
        except Exception as e:
            show_error(self, "โหลด Error", str(e))
            self.free_doctors_status.setText(f"❌ โหลดไม่สำเร็จ: {str(e)}")
            self.free_doctors_status.setStyleSheet("padding: 8px; background-color: #fee2e2; color: #dc2626; border-radius: 6px; margin: 10px 0;")

    def clear_free_doctors_form(self):
        """ล้างฟอร์มแพทย์ที่ไม่รับเงิน"""
        self.free_doctors_text.clear()
        self.free_doctors_status.setText("📄 ล้างฟอร์มแล้ว")
        self.free_doctors_status.setStyleSheet("padding: 8px; background-color: #f8fafc; border-radius: 6px; margin: 10px 0;")

    def on_free_physicians(self):
        """เปิดหน้าจัดการรายชื่อแพทย์ที่ไม่คิด Physician Fee"""
        try:
            master_dir = self.master_dir_edit.text().strip()
            if not master_dir:
                show_error(self, "Error", "กรุณาระบุโฟลเดอร์ master ก่อน")
                pass
            os.makedirs(master_dir, exist_ok=True)

            from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton

            dlg = QDialog(self)
            dlg.setWindowTitle("🩺 จัดการรายชื่อไม่คิด Physician Fee")
            dlg.resize(700, 560)
            dlg.setStyleSheet(get_style_sheet())

            layout = QVBoxLayout(dlg)
            header = QLabel("🩺 จัดการรายชื่อไม่คิด Physician Fee")
            header.setProperty("class", "header")
            header.setAlignment(Qt.AlignCenter)
            layout.addWidget(header)

            tab = self.create_free_physicians_tab(master_dir)
            layout.addWidget(tab)

            close_btn = QPushButton("❌ ปิด")
            close_btn.setProperty("class", "secondary")
            close_btn.clicked.connect(dlg.accept)
            layout.addWidget(close_btn)

            dlg.exec()
            self.status.setText("✅ เปิดจัดการไม่คิด Physician Fee แล้ว")
            self.status.setProperty("class", "status-success")
        except Exception as e:
            show_error(self, "Error", str(e), details=traceback.format_exc())
            self.status.setText("❌ เกิดข้อผิดพลาด")
            self.status.setProperty("class", "status-error")

    def create_free_physicians_tab(self, master_dir):
        """สร้างหน้าใส่รายชื่อที่ไม่คิด Physician Fee ตามวันที่"""
        from PySide6.QtWidgets import (
            QWidget,
            QVBoxLayout,
            QLabel,
            QTextEdit,
            QHBoxLayout,
            QPushButton,
            QGroupBox,
            QFormLayout,
        )

        tab = QWidget()
        layout = QVBoxLayout(tab)

        info_label = QLabel(
            "🩺 รายชื่อในหน้านี้จะใช้ตัดเฉพาะ Physician Fee ตามวันที่\n"
            "กรอกชื่อได้ทีละบรรทัด หรือคั่นด้วยเครื่องหมายจุลภาค (,)"
        )
        info_label.setStyleSheet(
            "color: #2563eb; font-size: 14px; padding: 10px; "
            "background-color: #dbeafe; border-radius: 6px; margin: 10px 0;"
        )
        layout.addWidget(info_label)

        form_group = QGroupBox("📋 ข้อมูลรายชื่อไม่คิด Physician Fee")
        form_layout = QFormLayout(form_group)

        self.free_phys_date_edit = QDateEdit()
        self.free_phys_date_edit.setCalendarPopup(True)
        self.free_phys_date_edit.setDate(QDate.currentDate())
        self.free_phys_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.free_phys_date_edit.setMinimumHeight(32)
        form_layout.addRow("📅 วันที่:", self.free_phys_date_edit)

        self.free_physicians_text = QTextEdit()
        self.free_physicians_text.setPlaceholderText(
            "ตัวอย่าง:\n"
            "นพ.สมชาย ใจดี\n"
            "พญ.สุดา แพทย์ไทย, นพ.ธนา สุขใจ"
        )
        self.free_physicians_text.setMinimumHeight(220)
        form_layout.addRow("👨‍⚕️ รายชื่อแพทย์:", self.free_physicians_text)

        layout.addWidget(form_group)

        btn_layout = QHBoxLayout()

        load_btn = QPushButton("📂 โหลด")
        load_btn.setProperty("class", "secondary")
        load_btn.clicked.connect(lambda: self.load_free_physicians_simple(master_dir))
        btn_layout.addWidget(load_btn)

        save_btn = QPushButton("💾 บันทึก")
        save_btn.setProperty("class", "success")
        save_btn.clicked.connect(lambda: self.save_free_physicians_simple(master_dir))
        btn_layout.addWidget(save_btn)

        clear_btn = QPushButton("🧹 ล้าง")
        clear_btn.setProperty("class", "secondary")
        clear_btn.clicked.connect(self.clear_free_physicians_form)
        btn_layout.addWidget(clear_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.free_physicians_status = QLabel(
            "📄 เลือกวันที่แล้วกดโหลด หรือกรอกรายชื่อแล้วกดบันทึก"
        )
        self.free_physicians_status.setStyleSheet(
            "padding: 8px; background-color: #f8fafc; border-radius: 6px; margin: 10px 0;"
        )
        layout.addWidget(self.free_physicians_status)

        self.load_free_physicians_simple(master_dir)
        self.free_phys_date_edit.dateChanged.connect(
            lambda _date: self.load_free_physicians_simple(master_dir)
        )

        return tab

    def save_free_physicians_simple(self, master_dir):
        """บันทึกรายชื่อแพทย์ที่ไม่คิด Physician Fee แบบง่าย"""
        try:
            selected_date = self.free_phys_date_edit.date().toString("yyyy-MM-dd")
            names_text = self.free_physicians_text.toPlainText().strip()

            if not names_text:
                show_info(self, "บันทึก", "กรุณากรอกชื่อแพทย์อย่างน้อย 1 คน")

            names = []
            for line in names_text.split("\n"):
                line = line.strip()
                if not line:
                    continue
                for name in line.split(","):
                    name = name.strip()
                    if name:
                        names.append(name)

            if not names:
                show_info(self, "บันทึก", "กรุณากรอกชื่อแพทย์อย่างน้อย 1 คน")
                return

            os.makedirs(master_dir, exist_ok=True)
            file_path = os.path.join(master_dir, f"free_physicians_{selected_date}.txt")
            with open(file_path, "w", encoding="utf-8") as f:
                for name in names:
                    f.write(f"{name}\n")

            self.free_physicians_status.setText(
                f"✅ บันทึกแล้ว {len(names)} คน วันที่ {selected_date}"
            )
            self.free_physicians_status.setStyleSheet(
                "padding: 8px; background-color: #dcfce7; color: #16a34a; border-radius: 6px; margin: 10px 0;"
            )
            show_info(
                self,
                "บันทึกสำเร็จ",
                f"✅ บันทึกรายชื่อไม่คิด Physician Fee {len(names)} คน\n📅 วันที่: {selected_date}",
            )
        except Exception as e:
            show_error(self, "บันทึก Error", str(e))
            self.free_physicians_status.setText(f"❌ บันทึกไม่สำเร็จ: {str(e)}")
            self.free_physicians_status.setStyleSheet(
                "padding: 8px; background-color: #fee2e2; color: #dc2626; border-radius: 6px; margin: 10px 0;"
            )

    def load_free_physicians_simple(self, master_dir):
        """โหลดรายชื่อแพทย์ที่ไม่คิด Physician Fee แบบง่าย"""
        try:
            selected_date = self.free_phys_date_edit.date().toString("yyyy-MM-dd")
            file_path = os.path.join(master_dir, f"free_physicians_{selected_date}.txt")

            if os.path.exists(file_path):
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read().strip()

                if content:
                    self.free_physicians_text.setPlainText(content)
                    count = len([line for line in content.split("\n") if line.strip()])
                    self.free_physicians_status.setText(
                        f"📂 โหลดแล้ว {count} คน วันที่ {selected_date}"
                    )
                    self.free_physicians_status.setStyleSheet(
                        "padding: 8px; background-color: #dbeafe; color: #2563eb; border-radius: 6px; margin: 10px 0;"
                    )
                else:
                    self.free_physicians_text.clear()
                    self.free_physicians_status.setText(f"📄 ไม่มีข้อมูล วันที่ {selected_date}")
                    self.free_physicians_status.setStyleSheet(
                        "padding: 8px; background-color: #f8fafc; border-radius: 6px; margin: 10px 0;"
                    )
            else:
                self.free_physicians_text.clear()
                self.free_physicians_status.setText(f"📄 ไม่มีข้อมูล วันที่ {selected_date}")
                self.free_physicians_status.setStyleSheet(
                    "padding: 8px; background-color: #f8fafc; border-radius: 6px; margin: 10px 0;"
                )
        except Exception as e:
            show_error(self, "โหลด Error", str(e))
            self.free_physicians_status.setText(f"❌ โหลดไม่สำเร็จ: {str(e)}")
            self.free_physicians_status.setStyleSheet(
                "padding: 8px; background-color: #fee2e2; color: #dc2626; border-radius: 6px; margin: 10px 0;"
            )

    def clear_free_physicians_form(self):
        """ล้างฟอร์มรายชื่อที่ไม่คิด Physician Fee"""
        self.free_physicians_text.clear()
        self.free_physicians_status.setText("📄 ล้างฟอร์มแล้ว")
        self.free_physicians_status.setStyleSheet(
            "padding: 8px; background-color: #f8fafc; border-radius: 6px; margin: 10px 0;"
        )

    def load_hospital_list(self, master_dir):
        """โหลดรายชื่อโรงพยาบาล"""
        try:
            self.hospital_list.setRowCount(0)
            
            hospitals_file = os.path.join(master_dir, "hospitals.json")
            if not os.path.exists(hospitals_file):
                self.hospital_list.setRowCount(1)
                self.hospital_list.setItem(0, 1, QTableWidgetItem("ยังไม่มีข้อมูลโรงพยาบาล"))
                return
            
            import json
            with open(hospitals_file, 'r', encoding='utf-8') as f:
                hospitals = json.load(f)
            
            for idx, hospital in enumerate(hospitals):
                name = str(hospital.get('name', '') or '').strip()
                types = hospital.get('types', []) or []
                types_text = ", ".join(
                    [
                        f"Type {str(t.get('type', '') or '').strip()}: {float(t.get('rate', 0) or 0):,.0f}"
                        for t in types
                    ]
                )
                row = self.hospital_list.rowCount()
                self.hospital_list.insertRow(row)

                no_item = QTableWidgetItem(str(idx + 1))
                name_item = QTableWidgetItem(name)
                types_item = QTableWidgetItem(types_text)
                for item in (no_item, name_item, types_item):
                    item.setData(Qt.UserRole, hospital)
                    item.setData(Qt.UserRole + 1, idx)
                    item.setToolTip(f"{name}\n{types_text}")

                self.hospital_list.setItem(row, 0, no_item)
                self.hospital_list.setItem(row, 1, name_item)
                self.hospital_list.setItem(row, 2, types_item)
                
        except Exception as e:
            self.hospital_list.setRowCount(1)
            self.hospital_list.setItem(0, 1, QTableWidgetItem(f"ไม่สามารถโหลดข้อมูล: {e}"))

    def edit_selected_hospital(self, master_dir):
        """โหลดข้อมูลโรงพยาบาลที่เลือกกลับเข้าแบบฟอร์มเพื่อแก้ไข"""
        try:
            row = self.hospital_list.currentRow()
            if row < 0:
                show_info(self, "แก้ไข", "กรุณาเลือกโรงพยาบาลที่ต้องการแก้ไข")
                return

            item = self.hospital_list.item(row, 1) or self.hospital_list.item(row, 0)
            if item is None:
                show_info(self, "แก้ไข", "รายการนี้ไม่มีข้อมูลสำหรับแก้ไข")
                return

            hospital = item.data(Qt.UserRole)
            edit_index = item.data(Qt.UserRole + 1)
            if not isinstance(hospital, dict) or edit_index is None:
                show_info(self, "แก้ไข", "รายการนี้ไม่มีข้อมูลสำหรับแก้ไข")
                return

            self.hospital_edit_index = int(edit_index)
            self.hospital_name_edit.setText(str(hospital.get('name', '') or ''))

            type_widgets = [
                (self.type1_combo, self.rate1_edit),
                (self.type2_combo, self.rate2_edit),
                (self.type3_combo, self.rate3_edit)
            ]
            for type_combo, rate_edit in type_widgets:
                type_combo.setCurrentText("")
                rate_edit.clear()

            for (type_combo, rate_edit), type_info in zip(type_widgets, hospital.get('types', [])):
                type_value = str(type_info.get('type', '') or '')
                rate_value = type_info.get('rate', '')
                type_combo.setCurrentText(type_value)
                if type_value == "ไม่จ่าย":
                    rate_edit.setText("")
                else:
                    rate_edit.setText("" if rate_value in ("", None) else str(rate_value))

            if hasattr(self, 'hospital_save_btn'):
                self.hospital_save_btn.setText("💾 บันทึกการแก้ไข")
            if hasattr(self, 'hospital_tab_widget'):
                self.hospital_tab_widget.setCurrentIndex(0)

        except Exception as e:
            show_error(self, "แก้ไข Error", str(e))

    def delete_selected_hospitals(self, master_dir):
        """ลบโรงพยาบาลที่เลือก"""
        try:
            selected_rows = [
                idx.row()
                for idx in self.hospital_list.selectionModel().selectedRows()
            ]
            
            if not selected_rows:
                show_info(self, "ลบ", "กรุณาเลือกโรงพยาบาลที่ต้องการลบ")
                return
            
            # สร้าง Dialog ยืนยัน
            from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout
            
            confirm_dlg = QDialog(self)
            confirm_dlg.setWindowTitle("🗑️ ยืนยันการลบ")
            confirm_dlg.resize(400, 200)
            confirm_dlg.setStyleSheet(get_style_sheet())
            
            layout = QVBoxLayout(confirm_dlg)
            
            header = QLabel("🗑️ ยืนยันการลบโรงพยาบาล")
            header.setProperty("class", "header")
            header.setAlignment(Qt.AlignCenter)
            layout.addWidget(header)
            
            warning = QLabel(f"⚠️ คุณกำลังจะลบ {len(selected_rows)} โรงพยาบาล\n\n❗️ ข้อมูลจะถูกลบถาววิ!")
            warning.setStyleSheet("color: #dc2626; font-size: 14px; padding: 15px; background-color: #fee2e2; border-radius: 6px;")
            layout.addWidget(warning)
            
            btn_layout = QHBoxLayout()
            
            delete_btn = QPushButton("🗑️ ลบ")
            delete_btn.setProperty("class", "danger")
            delete_btn.clicked.connect(lambda: self.perform_hospital_deletion(master_dir, selected_rows, confirm_dlg))
            
            cancel_btn = QPushButton("❌ ยกเลิก")
            cancel_btn.setProperty("class", "secondary")
            cancel_btn.clicked.connect(confirm_dlg.reject)
            
            btn_layout.addWidget(delete_btn)
            btn_layout.addWidget(cancel_btn)
            layout.addLayout(btn_layout)
            
            confirm_dlg.exec()
            
        except Exception as e:
            show_error(self, "ลบ Error", str(e))

    def perform_hospital_deletion(self, master_dir, selected_rows, confirm_dlg):
        """ดำเนินการลบโรงพยาบาลจริง"""
        try:
            hospitals_file = os.path.join(master_dir, "hospitals.json")
            
            if not os.path.exists(hospitals_file):
                confirm_dlg.accept()
                return
            
            import json
            with open(hospitals_file, 'r', encoding='utf-8') as f:
                hospitals = json.load(f)
            
            selected_indices = {
                int(item.data(Qt.UserRole + 1))
                for row in selected_rows
                for item in [self.hospital_list.item(row, 1) or self.hospital_list.item(row, 0)]
                if item is not None and item.data(Qt.UserRole + 1) is not None
            }
            if not selected_indices:
                confirm_dlg.accept()
                show_info(self, "ลบ", "ไม่พบข้อมูลโรงพยาบาลที่เลือก กรุณารีเฟรชรายการแล้วลองใหม่")
                return

            # ลบตาม index เพื่อไม่ให้รายการชื่อซ้ำ/ข้อมูลซ้ำถูกลบผิดตัว
            remaining_hospitals = [
                hospital
                for idx, hospital in enumerate(hospitals)
                if idx not in selected_indices
            ]
            
            # บันทึกข้อมูลใหม่
            with open(hospitals_file, 'w', encoding='utf-8') as f:
                json.dump(remaining_hospitals, f, ensure_ascii=False, indent=2)
            
            # ปิด dialog
            confirm_dlg.accept()
            
            # รีเฟรชรายการ
            self.load_hospital_list(master_dir)
            
            show_info(self, "ลบสำเร็จ", f"✅ ลบโรงพยาบาล {len(selected_indices)} แห่งแล้ว")
            
        except Exception as e:
            show_error(self, "ลบ Error", str(e))

    def export_free_employees(self, free_employees):
        """Export free employees to Excel"""
        try:
            if free_employees.empty:
                show_info(self, "Export", "ไม่มีข้อมูลพนักงานฟรีให้ Export")
                return
            
            # Select save location
            out_path, _ = QFileDialog.getSaveFileName(
                self, "บันทึกรายชื่อพนักงานฟรี", "free_employees.xlsx", "Excel Files (*.xlsx)"
            )
            
            if not out_path:
                return
            
            # Prepare data for export
            export_data = free_employees.copy()
            export_data['สถานะ'] = 'ไม่รับเงิน (Free)'
            
            # Export to Excel
            from src.excel_io import to_excel_bytes
            out_bytes = to_excel_bytes({"พนักงานฟรี": export_data})
            
            try:
                with open(out_path, "wb") as f:
                    f.write(out_bytes)
            except PermissionError:
                show_info(
                    self,
                    "บันทึกไม่สำเร็จ",
                    "ไม่สามารถเขียนไฟล์ได้ เพราะไฟล์ปลายทางอาจกำลังถูกใช้งานอยู่\n"
                    f"ไฟล์: {out_path}\n\n"
                    "กรุณาปิดไฟล์ใน Excel (หรือโปรแกรมอื่น) แล้วลองบันทึกใหม่ "
                    "หรือเลือกชื่อไฟล์ใหม่",
                )
                self.status.setText("❌ บันทึกไฟล์ไม่ได้ (Permission denied)")
                self.status.setProperty("class", "status-error")
                return
            
            show_info(self, "Export สำเร็จ", f"✅ บันทึกรายชื่อพนักงานฟรีแล้ว:\n{out_path}")
            
        except Exception as e:
            show_error(self, "Export Error", str(e))

    def on_calculate(self):
        try:
            self.status.setText("⏳ กำลังคำนวณ...")
            self.status.setProperty("class", "status")
            self._set_loading_visible(True, "กำลังเตรียมข้อมูลสำหรับการคำนวณ")

            self._update_loading_detail("กำลังอ่านไฟล์เคส")
            cases_df = self._read_cases_df()
            self._update_loading_detail("กำลังโหลดอัตราโรงพยาบาล")
            hospital_fee_map = self._read_hospital_fee_map()
            self._update_loading_detail("กำลังโหลดข้อมูล master และ alias")
            employees, emp_map, alias_map = self._load_master()
            master_dir = self.master_dir_edit.text().strip()
            as_of = pd.Timestamp(self.date_edit.date().toPython())

            self._update_loading_detail("กำลังตรวจสอบรายชื่อที่ยังไม่รู้จัก")
            unknown_df = find_unknown_names(cases_df, emp_map, alias_map)
            unknown_report_path = ""
            unknown_count = 0
            if unknown_df is not None and not unknown_df.empty:
                unknown_count = len(unknown_df)
                unknown_report_path = self._export_unknown_names_report(unknown_df)
                self.status.setText(
                    f"⚠️ พบชื่อที่ยังไม่ได้ผูก alias {unknown_count} รายการ "
                    "กำลังคำนวณเฉพาะรายชื่อที่จับคู่ได้"
                )
                self.status.setProperty("class", "status")

            self._update_loading_detail("กำลังคำนวณค่าตอบแทน")
            (
                case_out,
                payroll_summary,
                pay_details,
                physician_by_hospital,
                doctor_summary,
                doctor_details,
            ) = compute_case_fees(
                cases_df=cases_df,
                employees_df=employees,
                emp_map=emp_map,
                alias_map=alias_map,
                hospital_fee_map=hospital_fee_map,
                master_dir=master_dir,
                as_of=as_of,
            )

            # ✅ keep for detail join
            self.case_out = case_out.copy() if case_out is not None else pd.DataFrame()

            self.pay_details = (
                pay_details.copy() if pay_details is not None else pd.DataFrame()
            )
            payroll_summary = self._build_employee_role_summary(self.pay_details)
            self.emp_summary_all = (
                payroll_summary.copy()
                if payroll_summary is not None
                else pd.DataFrame()
            )

            self.doctor_details = (
                doctor_details.copy() if doctor_details is not None else pd.DataFrame()
            )
            self.doctor_summary_all = (
                doctor_summary.copy() if doctor_summary is not None else pd.DataFrame()
            )

            self._enrich_doctor_summary()
            self.apply_summary_filter()

            self._update_loading_detail("กำลังเตรียมไฟล์ผลลัพธ์")
            out_path, _ = QFileDialog.getSaveFileName(
                self, "บันทึกผลลัพธ์ Excel", "fee_result.xlsx", "Excel Files (*.xlsx)"
            )
            if not out_path:
                if unknown_report_path:
                    self.status.setText(
                        "⚠️ คำนวณเสร็จ (ยังไม่บันทึก) "
                        f"และพบชื่อที่ต้องผูก alias {unknown_count} รายการ"
                    )
                else:
                    self.status.setText("⏸️ คำนวณเสร็จ (ยังไม่บันทึก)")
                self.status.setProperty("class", "status")
                return

            case_out_export = case_out.drop(columns=["_join_key"], errors="ignore")
            pay_details_export = self.pay_details.drop(
                columns=["join_key"], errors="ignore"
            )
            doctor_details_export = self.doctor_details.drop(
                columns=["join_key"], errors="ignore"
            )

            export_sheets = {
                "CaseResult": case_out_export,
                "PayrollSummary": self.emp_summary_all,
                "PayDetails": pay_details_export,
                "PhysicianByHospital": physician_by_hospital,
                "DoctorSummary": self.doctor_summary_all,
                "DoctorDetails": doctor_details_export,
            }
            required_blank_columns = {}
            if unknown_df is not None and not unknown_df.empty:
                review_export = unknown_df.copy()
                review_export["emp_code"] = ""
                review_export["หมายเหตุ"] = (
                    "กรอก emp_code ให้ตรงกับพนักงาน แล้วนำข้อมูลนี้ไปเพิ่มใน aliases.xlsx"
                )
                export_sheets["AliasesToReview"] = review_export
                required_blank_columns["AliasesToReview"] = ["emp_code"]

            out_bytes = to_excel_bytes(
                export_sheets,
                highlights={
                    "CaseResult": self._build_case_result_highlights(
                        case_out_export, unknown_df
                    )
                },
                required_blank_columns=required_blank_columns,
            )
            self._update_loading_detail("กำลังบันทึกไฟล์ Excel")
            try:
                with open(out_path, "wb") as f:
                    f.write(out_bytes)
            except PermissionError:
                show_info(
                    self,
                    "Export ไม่สำเร็จ",
                    "ไม่สามารถเขียนไฟล์ได้ เพราะไฟล์ปลายทางอาจกำลังถูกใช้งานอยู่\n"
                    f"ไฟล์: {out_path}\n\n"
                    "กรุณาปิดไฟล์ใน Excel (หรือโปรแกรมอื่น) แล้วลองใหม่",
                )
                return

            success_msg = f"✅ คำนวณและบันทึกผลลัพธ์แล้ว:\n{out_path}"
            if unknown_report_path:
                success_msg += (
                    "\n\n⚠️ มีชื่อ Monitor ที่ยังจับคู่กับพนักงานไม่ได้ "
                    f"{unknown_count} รายการ\n"
                    "ระบบคำนวณเฉพาะรายชื่อที่จับคู่ได้ และสร้างไฟล์ตรวจสอบไว้ที่:\n"
                    f"{unknown_report_path}\n\n"
                    "ในไฟล์ผลลัพธ์ แถวที่ต้องตรวจสอบจะถูกไฮไลต์สีเหลือง/ส้มในชีท CaseResult"
                )

            show_info(self, "สำเร็จ", success_msg)
            if unknown_report_path:
                self.status.setText(
                    f"✅ คำนวณเสร็จแล้ว (มีชื่อรอผูก alias {unknown_count} รายการ)"
                )
            else:
                self.status.setText("✅ คำนวณเสร็จแล้ว")
            self.status.setProperty("class", "status-success")

        except Exception as e:
            show_error(self, "Error", str(e), details=traceback.format_exc())
            self.status.setText("❌ เกิดข้อผิดพลาด")
            self.status.setProperty("class", "status-error")
        finally:
            self._set_loading_visible(False)

    # -------------------------
    # Summary: switch + filter + render
    # -------------------------
    def _build_employee_role_summary(self, pay_details: pd.DataFrame) -> pd.DataFrame:
        cols = [
            "emp_code",
            "display_name",
            "monitor_count",
            "monitor_amount",
            "scoring_count",
            "scoring_amount",
            "found_count",
            "total_amount",
        ]
        if pay_details is None or pay_details.empty:
            return pd.DataFrame(columns=cols)

        df = pay_details.copy()
        for col in ["emp_code", "display_name", "role", "amount"]:
            if col not in df.columns:
                df[col] = "" if col != "amount" else 0.0

        df["role"] = df["role"].astype(str).str.strip()
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0.0)

        df["summary_role"] = ""
        df.loc[df["role"].isin(["MonitorTech", "MonitorTech2"]), "summary_role"] = "monitor"
        df.loc[df["role"] == "ScoringTech", "summary_role"] = "scoring"
        df = df[df["summary_role"].isin(["monitor", "scoring"])].copy()
        if df.empty:
            return pd.DataFrame(columns=cols)

        grouped = (
            df.groupby(["emp_code", "display_name", "summary_role"], as_index=False)
            .agg(case_count=("amount", "size"), amount=("amount", "sum"))
        )
        pivot = grouped.pivot_table(
            index=["emp_code", "display_name"],
            columns="summary_role",
            values=["case_count", "amount"],
            aggfunc="sum",
            fill_value=0,
        )
        pivot.columns = [
            f"{role}_{metric}" for metric, role in pivot.columns.to_flat_index()
        ]
        pivot = pivot.reset_index()

        rename_map = {
            "monitor_case_count": "monitor_count",
            "monitor_amount": "monitor_amount",
            "scoring_case_count": "scoring_count",
            "scoring_amount": "scoring_amount",
        }
        pivot.rename(columns=rename_map, inplace=True)

        for col in ["monitor_count", "monitor_amount", "scoring_count", "scoring_amount"]:
            if col not in pivot.columns:
                pivot[col] = 0

        pivot["monitor_count"] = pd.to_numeric(
            pivot["monitor_count"], errors="coerce"
        ).fillna(0).astype(int)
        pivot["scoring_count"] = pd.to_numeric(
            pivot["scoring_count"], errors="coerce"
        ).fillna(0).astype(int)
        pivot["monitor_amount"] = pd.to_numeric(
            pivot["monitor_amount"], errors="coerce"
        ).fillna(0.0)
        pivot["scoring_amount"] = pd.to_numeric(
            pivot["scoring_amount"], errors="coerce"
        ).fillna(0.0)
        pivot["found_count"] = pivot["monitor_count"] + pivot["scoring_count"]
        pivot["total_amount"] = pivot["monitor_amount"] + pivot["scoring_amount"]

        return pivot[cols]

    def _setup_table_headers_for_mode(self, mode: str):
        if mode == "DOC":
            headers = ["doctor_key", "doctor_name", "found_count", "total_amount"]
        else:
            headers = [
                "emp_code",
                "display_name",
                "monitor_count",
                "monitor_amount",
                "scoring_count",
                "scoring_amount",
                "found_count",
                "total_amount",
            ]

        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.resizeColumnsToContents()

    def _create_metric_card(self, label_text: str):
        frame = QFrame()
        frame.setProperty("class", "metric-card")
        frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(2)

        value = QLabel("-")
        value.setProperty("class", "metric-value")
        layout.addWidget(value)

        label = QLabel(label_text)
        label.setProperty("class", "metric-label")
        layout.addWidget(label)

        return frame, value, label

    def _set_metric_tooltips(self, mode: str):
        if mode == "DOC":
            self.metric_monitor_label.setText("พบทั้งหมด")
            self.metric_scoring_label.setText("แยกประเภท")
            tips = {
                self.metric_mode_card: "มุมมองปัจจุบัน: แพทย์",
                self.metric_rows_card: (
                    "จำนวนแถวแพทย์ที่แสดงอยู่ในตารางตอนนี้ "
                    "หลังใช้ช่องค้นหา"
                ),
                self.metric_monitor_card: (
                    "ผลรวมคอลัมน์ found_count ของแพทย์ที่แสดงอยู่\n"
                    "นับจาก Interpret MD เป็นหลัก\n"
                    "Scoring Tech จะนับเป็นแพทย์เฉพาะแถวที่คำนำหน้ามีค่า"
                ),
                self.metric_scoring_card: (
                    "มุมมองแพทย์ยังไม่ได้แยก Monitor/Scoring\n"
                    "ใช้การ์ดพบทั้งหมดสำหรับจำนวนเคสแพทย์รวม"
                ),
                self.metric_total_card: (
                    "ผลรวมคอลัมน์ total_amount ของแพทย์ที่แสดงอยู่\n"
                    "มาจาก Physician Fee และบวก Scoring Fee เพิ่มเมื่อแพทย์อยู่ใน Scoring Tech "
                    "พร้อมคำนำหน้า"
                ),
            }
        else:
            self.metric_monitor_label.setText("Monitor")
            self.metric_scoring_label.setText("Scoring")
            tips = {
                self.metric_mode_card: "มุมมองปัจจุบัน: พนักงาน",
                self.metric_rows_card: (
                    "จำนวนแถวพนักงานที่แสดงอยู่ในตารางตอนนี้ "
                    "หลังใช้ช่องค้นหา\n"
                    "มาจาก PayrollSummary หนึ่งแถวต่อ emp_code/display_name"
                ),
                self.metric_monitor_card: (
                    "ผลรวมคอลัมน์ monitor_count ของพนักงานที่แสดงอยู่\n"
                    "monitor_count มาจาก PayDetails role MonitorTech/MonitorTech2\n"
                    "จำนวนนี้มาจากเคสที่มี Monitor Fee หรือ Monitor Fee2"
                ),
                self.metric_scoring_card: (
                    "ผลรวมคอลัมน์ scoring_count ของพนักงานที่แสดงอยู่\n"
                    "scoring_count มาจาก PayDetails role ScoringTech\n"
                    "จำนวนนี้มาจากเคสที่มี Scoring Fee"
                ),
                self.metric_total_card: (
                    "ผลรวมคอลัมน์ total_amount ของพนักงานที่แสดงอยู่\n"
                    "total_amount = monitor_amount + scoring_amount\n"
                    "monitor_amount มาจาก Monitor Fee/Monitor Fee2\n"
                    "scoring_amount มาจาก Scoring Fee"
                ),
            }

        for widget, tip in tips.items():
            widget.setToolTip(tip)
            for child in widget.findChildren(QWidget):
                child.setToolTip(tip)

    def _relayout_action_buttons(self, compact: bool):
        if compact:
            positions = [
                (self.btn_add_emp, 0, 0, 1, 1),
                (self.btn_view_emp, 0, 1, 1, 1),
                (self.btn_summary_emp, 1, 0, 1, 1),
                (self.btn_free_emp, 1, 1, 1, 1),
                (self.btn_free_physician, 2, 0, 1, 1),
                (self.btn_hospital, 2, 1, 1, 1),
                (self.btn_calc, 3, 0, 1, 2),
            ]
            stretch_count = 2
        else:
            positions = [
                (self.btn_add_emp, 0, 0, 1, 1),
                (self.btn_view_emp, 0, 1, 1, 1),
                (self.btn_summary_emp, 0, 2, 1, 1),
                (self.btn_free_emp, 1, 0, 1, 1),
                (self.btn_free_physician, 1, 1, 1, 1),
                (self.btn_hospital, 1, 2, 1, 1),
                (self.btn_calc, 0, 3, 2, 1),
            ]
            stretch_count = 4

        for widget, row, col, row_span, col_span in positions:
            self.actions_layout.addWidget(widget, row, col, row_span, col_span)

        for i in range(4):
            self.actions_layout.setColumnStretch(i, 1 if i < stretch_count else 0)

    def _relayout_metric_cards(self, compact: bool):
        if compact:
            positions = [
                (self.metric_mode_card, 0, 0),
                (self.metric_rows_card, 0, 1),
                (self.metric_monitor_card, 1, 0),
                (self.metric_scoring_card, 1, 1),
                (self.metric_total_card, 2, 0),
            ]
            stretch_count = 2
        else:
            positions = [
                (self.metric_mode_card, 0, 0),
                (self.metric_rows_card, 0, 1),
                (self.metric_monitor_card, 0, 2),
                (self.metric_scoring_card, 0, 3),
                (self.metric_total_card, 0, 4),
            ]
            stretch_count = 5

        for widget, row, col in positions:
            self.stats_layout.addWidget(widget, row, col)

        for i in range(5):
            self.stats_layout.setColumnStretch(i, 1 if i < stretch_count else 0)

    def _apply_responsive_layout(self):
        width = max(self.width(), 0)
        compact = width < 1180
        narrow = width < 980

        self.top_workspace_layout.setDirection(
            QBoxLayout.TopToBottom if narrow else QBoxLayout.LeftToRight
        )
        self._relayout_action_buttons(compact=compact)
        self._relayout_metric_cards(compact=narrow)

        header = self.table.horizontalHeader()
        if narrow:
            header.setSectionResizeMode(QHeaderView.Stretch)
        else:
            header.setSectionResizeMode(QHeaderView.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.Stretch)

    def _update_summary_metrics(self, df: pd.DataFrame, mode: str):
        mode_text = "พนักงาน" if mode == "EMP" else "แพทย์"
        row_count = len(df) if df is not None else 0

        self._set_metric_tooltips(mode)

        if df is None or df.empty:
            found_total = 0
            monitor_total = 0
            scoring_total = 0
            amount_total = 0.0
        else:
            found_total = int(
                pd.to_numeric(df.get("found_count", 0), errors="coerce").fillna(0).sum()
            )
            monitor_total = int(
                pd.to_numeric(df.get("monitor_count", 0), errors="coerce").fillna(0).sum()
            )
            scoring_total = int(
                pd.to_numeric(df.get("scoring_count", 0), errors="coerce").fillna(0).sum()
            )
            amount_total = float(
                pd.to_numeric(df.get("total_amount", 0.0), errors="coerce").fillna(0).sum()
            )

        self.metric_mode_value.setText(mode_text)
        self.metric_rows_value.setText(f"{row_count:,}")
        if mode == "DOC":
            self.metric_monitor_value.setText(f"{found_total:,}")
            self.metric_scoring_value.setText("-")
        else:
            self.metric_monitor_value.setText(f"{monitor_total:,}")
            self.metric_scoring_value.setText(f"{scoring_total:,}")
        self.metric_total_value.setText(f"{amount_total:,.2f}")

    def _enrich_doctor_summary(self):
        df = self.doctor_summary_all
        det = self.doctor_details
        if df is None or df.empty:
            return

        if "doctor_key" not in df.columns:
            df["doctor_key"] = ""

        if det is None or det.empty or "doctor_key" not in det.columns:
            df["doctor_name"] = df["doctor_key"]
            self.doctor_summary_all = df
            return

        det2 = det.copy()
        for c in ["doctor_key", "doctor_name_raw"]:
            if c not in det2.columns:
                det2[c] = ""

        name_map = {}
        for key, g in det2.groupby(det2["doctor_key"].astype(str)):
            if not key:
                continue
            source = g.get("source_col", pd.Series("", index=g.index)).astype(str)
            preferred = g[source.eq("Interpret MD")]
            if preferred.empty:
                preferred = g
            vc = preferred["doctor_name_raw"].astype(str).value_counts()
            name_map[key] = vc.index[0] if len(vc.index) else key

        df = df.copy()
        df["doctor_key"] = df["doctor_key"].astype(str)
        df["doctor_name"] = df["doctor_key"].map(name_map).fillna(df["doctor_key"])

        case_df = self.case_out.copy() if self.case_out is not None else pd.DataFrame()
        if case_df is not None and not case_df.empty:
            for col in ["คำนำหน้า", "Scoring Tech", "Interpret MD"]:
                if col not in case_df.columns:
                    case_df[col] = ""
            for col in ["Scoring Fee", "Physician Fee"]:
                if col not in case_df.columns:
                    case_df[col] = 0.0

            case_df["Scoring Fee"] = pd.to_numeric(
                case_df["Scoring Fee"], errors="coerce"
            ).fillna(0.0)
            case_df["Physician Fee"] = pd.to_numeric(
                case_df["Physician Fee"], errors="coerce"
            ).fillna(0.0)
            scoring_has_prefix = case_df["คำนำหน้า"].map(has_scoring_doctor_prefix)
            case_df["_scoring_key"] = case_df["Scoring Tech"].where(
                scoring_has_prefix, ""
            ).map(normalize_doctor_name)
            case_df["_interpret_key"] = case_df["Interpret MD"].map(normalize_doctor_name)

            doctor_case_rows = []
            for idx, row in case_df.iterrows():
                case_id = str(row.get("_join_key", idx))
                scoring_key = str(row.get("_scoring_key", "") or "").strip()
                interpret_key = str(row.get("_interpret_key", "") or "").strip()
                scoring_fee = float(row.get("Scoring Fee", 0.0) or 0.0)
                physician_fee = float(row.get("Physician Fee", 0.0) or 0.0)

                if scoring_key:
                    total_amount = scoring_fee
                    if interpret_key == scoring_key:
                        total_amount += physician_fee
                    doctor_case_rows.append(
                        {
                            "doctor_key": scoring_key,
                            "case_id": case_id,
                            "found_count": 1,
                            "total_amount": total_amount,
                        }
                    )

                if interpret_key and interpret_key != scoring_key:
                    doctor_case_rows.append(
                        {
                            "doctor_key": interpret_key,
                            "case_id": case_id,
                            "found_count": 1,
                            "total_amount": physician_fee,
                        }
                    )

            if doctor_case_rows:
                doctor_case_df = pd.DataFrame(doctor_case_rows)
                totals_df = doctor_case_df.groupby("doctor_key", as_index=False).agg(
                    found_count=("case_id", "size"),
                    total_amount=("total_amount", "sum"),
                )
                df = df.drop(columns=["found_count", "total_amount"], errors="ignore")
                df = df.merge(totals_df, on="doctor_key", how="left")
                df["found_count"] = (
                    pd.to_numeric(df["found_count"], errors="coerce")
                    .fillna(0)
                    .astype(int)
                )
                df["total_amount"] = pd.to_numeric(
                    df["total_amount"], errors="coerce"
                ).fillna(0.0)

        self.doctor_summary_all = df

    def apply_summary_filter(self):
        mode = self.view_mode.currentData() or "EMP"
        q = (self.summary_search.text() or "").strip().lower()

        if mode == "DOC":
            self._setup_table_headers_for_mode("DOC")
            df = (
                self.doctor_summary_all.copy()
                if self.doctor_summary_all is not None
                else pd.DataFrame()
            )

            if "doctor_key" not in df.columns:
                df["doctor_key"] = ""
            if "doctor_name" not in df.columns:
                df["doctor_name"] = df["doctor_key"]
            if "found_count" not in df.columns:
                df["found_count"] = 0
            if "total_amount" not in df.columns:
                df["total_amount"] = 0.0

            df["found_count"] = (
                pd.to_numeric(df["found_count"], errors="coerce").fillna(0).astype(int)
            )
            df["total_amount"] = pd.to_numeric(
                df["total_amount"], errors="coerce"
            ).fillna(0.0)

            if q:
                mask = df["doctor_key"].astype(str).str.lower().str.contains(
                    q, na=False
                ) | df["doctor_name"].astype(str).str.lower().str.contains(q, na=False)
                df = df[mask].copy()

            df.sort_values(
                ["total_amount", "found_count"], ascending=False, inplace=True
            )
            self.doctor_summary_view = df
            self._update_summary_metrics(df, mode="DOC")
            self._render_summary_table(df, mode="DOC")
        else:
            self._setup_table_headers_for_mode("EMP")
            df = (
                self.emp_summary_all.copy()
                if self.emp_summary_all is not None
                else pd.DataFrame()
            )

            if "emp_code" not in df.columns:
                df["emp_code"] = ""
            if "display_name" not in df.columns:
                df["display_name"] = ""
            for col in [
                "monitor_count",
                "monitor_amount",
                "scoring_count",
                "scoring_amount",
            ]:
                if col not in df.columns:
                    df[col] = 0
            if "found_count" not in df.columns:
                df["found_count"] = 0
            if "total_amount" not in df.columns:
                df["total_amount"] = 0.0

            for col in ["monitor_count", "scoring_count"]:
                df[col] = (
                    pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
                )
            for col in ["monitor_amount", "scoring_amount"]:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
            df["found_count"] = (
                pd.to_numeric(df["found_count"], errors="coerce").fillna(0).astype(int)
            )
            df["total_amount"] = pd.to_numeric(
                df["total_amount"], errors="coerce"
            ).fillna(0.0)

            if q:
                mask = df["emp_code"].astype(str).str.lower().str.contains(
                    q, na=False
                ) | df["display_name"].astype(str).str.lower().str.contains(q, na=False)
                df = df[mask].copy()

            df.sort_values(
                ["total_amount", "found_count"], ascending=False, inplace=True
            )
            self.emp_summary_view = df
            self._update_summary_metrics(df, mode="EMP")
            self._render_summary_table(df, mode="EMP")

    def _render_summary_table(self, df: pd.DataFrame, mode: str):
        if mode == "DOC":
            cols = ["doctor_key", "doctor_name", "found_count", "total_amount"]
        else:
            cols = [
                "emp_code",
                "display_name",
                "monitor_count",
                "monitor_amount",
                "scoring_count",
                "scoring_amount",
                "found_count",
                "total_amount",
            ]

        self.table.setRowCount(len(df))
        for r in range(len(df)):
            for c, col_name in enumerate(cols):
                val = df.iloc[r][col_name]

                if col_name in (
                    "monitor_count",
                    "monitor_amount",
                    "scoring_count",
                    "scoring_amount",
                    "found_count",
                    "total_amount",
                ):
                    if col_name.endswith("_count") or col_name == "found_count":
                        item = QTableWidgetItem(str(int(val)))
                        item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    else:
                        item = QTableWidgetItem(f"{float(val):,.2f}")
                        item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    item = QTableWidgetItem(str(val))

                self.table.setItem(r, c, item)

        self.table.resizeColumnsToContents()

    # -------------------------
    # Drill-down: open detail by mode
    # -------------------------
    def open_detail_by_mode(self, row: int, col: int):
        mode = self.view_mode.currentData() or "EMP"
        try:
            if mode == "DOC":
                if self.doctor_summary_view is None or self.doctor_summary_view.empty:
                    return
                if row < 0 or row >= len(self.doctor_summary_view):
                    return
                doctor_key = str(self.doctor_summary_view.iloc[row]["doctor_key"])
                dlg = DoctorDetailDialog(
                    doctor_key, self.doctor_details, self.case_out, self
                )
                dlg.exec()
            else:
                if self.emp_summary_view is None or self.emp_summary_view.empty:
                    return
                if row < 0 or row >= len(self.emp_summary_view):
                    return
                emp_code = str(self.emp_summary_view.iloc[row]["emp_code"])
                display_name = str(self.emp_summary_view.iloc[row]["display_name"])

                # ✅ pass case_out for join
                dlg = PayrollDetailDialog(
                    emp_code, display_name, self.pay_details, self.case_out, self
                )
                dlg.exec()
        except Exception as e:
            show_error(self, "Error", f"เปิดรายละเอียดไม่สำเร็จ: {e}")


def main():
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "kasidid.salarycalc.desktop"
        )
    except Exception:
        pass

    app = QApplication(sys.argv)
    app.setApplicationName("Salary & Fee Calculator")
    app.setOrganizationName("Kasidid")
    apply_light_app_palette(app)
    icon = _load_app_icon()
    if not icon.isNull():
        app.setWindowIcon(icon)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
