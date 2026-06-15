# gen_case_500.py
from __future__ import annotations

import argparse
import calendar
import random
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Optional

import pandas as pd

# UI is optional: script works as CLI even if PySide6 not installed
try:
    from PySide6 import QtWidgets
except Exception:  # pragma: no cover
    QtWidgets = None  # type: ignore


# -----------------------------
# Columns (match your main app)
# -----------------------------
HEADERS = [
    "ลำดับ",
    "จำนวนเคส",
    "วันที่ตรวจ",
    "ชื่อคนไข้",
    "HN",
    "โรงพยาบาล",
    "Type",
    "Program",
    "สิทธิ์",
    "หมอส่ง",
    "Monitor Tech",
    "Monitor Fee",
    "Monitor Tech (2)",
    "Monitor Fee2",
    "Scoring Fee",
    "คำนำหน้า",
    "Scoring Tech",
    "Scoring Fee2",
    "Interpret MD",
    "Physician Fee",
    "IV",
    "ราคารวมIV",
    "15 เคสแรก",
    "ส่วนลด",
]

DOC_POOL = [
    ("นพ", "นพ.ชาญสุรี"),
    ("พญ", "พญ.นาถยศ"),
    ("", "ชาญสุรี"),
    ("", "นาถยศ"),
    ("", "ประภาพร"),
    ("", "ชยพล กวางดี"),
]

HOSP_POOL = [
    ("พญาไท 3", 1, "SN/FN"),
    ("บำรุงราษฎร์", 2, "SN/FN"),
    ("สมิติเวช", 1, "SN/FN"),
]


# -----------------------------
# Utils
# -----------------------------
def pick_writable_path(path: Path) -> Path:
    """
    ถ้าไฟล์ถูกล็อก/เขียนทับไม่ได้ (เช่น เปิดใน Excel) ให้หาชื่อใหม่ _1, _2, ...
    """
    if not path.exists():
        return path

    stem, suf, parent = path.stem, path.suffix, path.parent
    for i in range(1, 1000):
        cand = parent / f"{stem}_{i}{suf}"
        if not cand.exists():
            return cand
    raise RuntimeError("ไม่สามารถหา filename ว่างได้ (ลองเปลี่ยนโฟลเดอร์/ชื่อไฟล์)")


def gen_dates_in_month(
    n: int,
    year: int,
    month: int,
    *,
    force_first_days: int,
    rng: random.Random,
) -> list[date]:
    """
    ✅ สร้างวันที่ให้อยู่ในเดือนเดียวกันเสมอ (ไม่ข้ามเดือน)
    - force_first_days จะทำให้ช่วงต้นกระจุกวัน 1..force_first_days (วนซ้ำ) เพื่อให้ลำดับชัด
    """
    last_day = calendar.monthrange(year, month)[1]
    days = list(range(1, last_day + 1))

    out: list[date] = []
    for i in range(n):
        if force_first_days > 0 and i < (force_first_days * 2):
            d = 1 + (i % min(force_first_days, last_day))
        else:
            d = rng.choice(days)
        out.append(date(year, month, d))
    out.sort()
    return out


def load_employee_names(employees_xlsx: Path) -> dict[str, list[str]]:
    """
    Return {"all": [...], "cond15": [...], "non_cond15": [...]}
    Uses display_name first then full_name.
    Filters active==True if exists.
    """
    if not employees_xlsx.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์: {employees_xlsx.resolve()}")

    df = pd.read_excel(employees_xlsx)

    if "active" in df.columns:
        df = df[df["active"] == True].copy()

    name_col = "display_name" if "display_name" in df.columns else None
    if name_col is None:
        name_col = "full_name" if "full_name" in df.columns else None
    if name_col is None:
        raise RuntimeError("employees.xlsx ต้องมีคอลัมน์ display_name หรือ full_name")

    df[name_col] = df[name_col].astype(str).str.strip()
    df = df[df[name_col] != ""].copy()

    if df.empty:
        raise RuntimeError("employees.xlsx ไม่มีพนักงาน active หรือไม่มีชื่อให้ใช้")

    if "rate_mode" not in df.columns:
        all_names = df[name_col].tolist()
        return {"all": all_names, "cond15": [], "non_cond15": all_names}

    df["rate_mode"] = df["rate_mode"].astype(str).str.strip().str.upper()
    cond15 = df[df["rate_mode"] == "COND15"][name_col].tolist()
    non_cond15 = df[df["rate_mode"] != "COND15"][name_col].tolist()
    all_names = df[name_col].tolist()
    return {"all": all_names, "cond15": cond15, "non_cond15": non_cond15}


def pick_monitor1(
    i: int,
    dt: date,
    emp_groups: dict[str, list[str]],
    *,
    mode: str,
    force_cond15_rows: int,
    p_cond15_random: float,
    target_month: int,
    rng: random.Random,
) -> str:
    cond15 = emp_groups["cond15"]
    non_cond15 = emp_groups["non_cond15"]
    all_names = emp_groups["all"]

    # force early rows to first COND15 (only in COND15/MIX)
    if (
        mode in ("COND15", "MIX")
        and dt.month == target_month
        and i <= force_cond15_rows
        and cond15
    ):
        return cond15[0]

    # random some COND15
    if mode in ("COND15", "MIX") and cond15 and rng.random() < p_cond15_random:
        return rng.choice(cond15)

    if non_cond15:
        return rng.choice(non_cond15)
    return rng.choice(all_names)


def pick_monitor2(
    mt1: str,
    emp_groups: dict[str, list[str]],
    *,
    p_monitor2: float,
    rng: random.Random,
) -> str:
    if rng.random() >= p_monitor2:
        return ""

    # 15% chance same as monitor1 (test duplicates)
    if rng.random() < 0.15:
        return mt1

    non_cond15 = emp_groups["non_cond15"]
    all_names = emp_groups["all"]
    pool = non_cond15 if non_cond15 else all_names
    return rng.choice(pool)


@dataclass
class GenConfig:
    master_dir: str = "master"
    employees_xlsx: str = ""  # override if set
    out: str = "case_sample_500.xlsx"
    sheet: str = "Cases"

    mode: str = "COND15"  # COND15 / DISCOUNT / MIX
    n: int = 500
    year: int = 2026
    month: int = 2

    force_cond15: int = 25
    p_cond15_random: float = 0.12
    p_monitor2: float = 0.35
    p_discount: float = 0.10
    seed: int = 42
    force_first_days: int = 15


def generate_cases(cfg: GenConfig) -> Path:
    rng = random.Random(cfg.seed)

    master_dir = Path(cfg.master_dir)
    employees_xlsx = (
        Path(cfg.employees_xlsx)
        if cfg.employees_xlsx
        else (master_dir / "employees.xlsx")
    )
    emp_groups = load_employee_names(employees_xlsx)

    dates = gen_dates_in_month(
        cfg.n, cfg.year, cfg.month, force_first_days=cfg.force_first_days, rng=rng
    )

    rows = []
    for i in range(1, cfg.n + 1):
        dt = dates[i - 1]
        hosp, typ, program = rng.choice(HOSP_POOL)

        hn = 100000 + i
        patient = f"คนไข้ทดสอบ{i:05d}"
        iv = f"IV{dt.strftime('%y%m')}{i:04d}"
        total_iv = rng.choice([1500, 2000, 2500, 3000])

        mt1 = pick_monitor1(
            i,
            dt,
            emp_groups,
            mode=cfg.mode,
            force_cond15_rows=cfg.force_cond15,
            p_cond15_random=cfg.p_cond15_random,
            target_month=cfg.month,
            rng=rng,
        )
        mt2 = pick_monitor2(mt1, emp_groups, p_monitor2=cfg.p_monitor2, rng=rng)

        prefix, docname = rng.choice(DOC_POOL)
        scoring_tech = docname
        interpret_md = rng.choice([docname, rng.choice(DOC_POOL)[1]])

        # ส่วนลด:
        # - ในโปรแกรมหลักคุณตีความ 0<x<1 เป็น "เปอร์เซ็นต์"
        # - ดังนั้น 50% ให้ใส่ 0.5
        disc_first15 = 0  # ปล่อย 0 (คอลัมน์นี้คุณใช้เป็น "บาท" อยู่)
        disc_other = 0
        if cfg.mode in ("DISCOUNT", "MIX") and rng.random() < cfg.p_discount:
            disc_other = rng.choice([0.5, 0.25])

        row = {
            "ลำดับ": i,
            "จำนวนเคส": 1,
            "วันที่ตรวจ": dt,  # ✅ เขียนเป็น date object เพื่อกัน parse สลับวัน/เดือน
            "ชื่อคนไข้": patient,
            "HN": str(hn),
            "โรงพยาบาล": hosp,
            "Type": typ,
            "Program": program,
            "สิทธิ์": "",
            "หมอส่ง": docname,
            "Monitor Tech": mt1,
            "Monitor Fee": 0,
            "Monitor Tech (2)": mt2,
            "Monitor Fee2": 0,
            "Scoring Fee": 0,
            "คำนำหน้า": prefix,
            "Scoring Tech": scoring_tech,
            "Scoring Fee2": 0,  # เผื่อ template เก่ายังมี
            "Interpret MD": interpret_md,
            "Physician Fee": 0,
            "IV": iv,
            "ราคารวมIV": total_iv,
            "15 เคสแรก": disc_first15,
            "ส่วนลด": disc_other,
        }
        rows.append(row)

    df = pd.DataFrame(rows, columns=HEADERS)

    out_path = pick_writable_path(Path(cfg.out))
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=cfg.sheet)

    return out_path


# -----------------------------
# CLI
# -----------------------------
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Generate case_sample.xlsx for testing salary_calc (UI/CLI)."
    )
    p.add_argument("--cli", action="store_true", help="รันแบบ CLI (ถ้าไม่ใส่จะเปิด UI)")
    p.add_argument(
        "--master-dir", default="master", help="โฟลเดอร์ master ที่มี employees.xlsx"
    )
    p.add_argument("--employees-xlsx", default="", help="ระบุไฟล์ employees.xlsx โดยตรง")
    p.add_argument("--out", default="case_sample_500.xlsx", help="ชื่อไฟล์ output .xlsx")
    p.add_argument("--sheet", default="Cases", help="ชื่อชีทใน output")

    p.add_argument("--mode", choices=["COND15", "DISCOUNT", "MIX"], default="COND15")
    p.add_argument("--n", type=int, default=500, help="จำนวนแถวทั้งหมด")
    p.add_argument("--year", type=int, default=2026)
    p.add_argument("--month", type=int, default=2)

    p.add_argument(
        "--force-cond15",
        type=int,
        default=25,
        help="บังคับแถวแรกๆ ให้เป็น COND15 (เฉพาะ COND15/MIX)",
    )
    p.add_argument(
        "--p-cond15-random",
        type=float,
        default=0.12,
        help="โอกาสสุ่มเจอ COND15 (เฉพาะ COND15/MIX)",
    )

    p.add_argument(
        "--p-monitor2", type=float, default=0.35, help="โอกาสมี Monitor Tech (2)"
    )
    # ❗ หลีกเลี่ยง % ใน help (argparse ใช้ % เป็น format)
    p.add_argument(
        "--p-discount",
        type=float,
        default=0.10,
        help="โอกาสมีส่วนลด (เปอร์เซ็นต์) (เฉพาะ DISCOUNT/MIX)",
    )

    p.add_argument("--seed", type=int, default=42)
    p.add_argument(
        "--force-first-days",
        type=int,
        default=15,
        help="ทำให้วันต้นเดือนหนาแน่น 1..N เพื่อให้ลำดับเคสชัด",
    )
    return p.parse_args()


# -----------------------------
# UI (single-file)
# -----------------------------
class GeneratorUI(QtWidgets.QDialog):  # type: ignore[misc]
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🧪 Case Generator (Test)")
        self.resize(560, 420)

        # widgets
        self.master_dir = QtWidgets.QLineEdit("master")
        self.btn_pick_master = QtWidgets.QPushButton("เลือกโฟลเดอร์ master")
        self.btn_pick_master.clicked.connect(self._pick_master_dir)

        self.employees_xlsx = QtWidgets.QLineEdit("")
        self.btn_pick_emp = QtWidgets.QPushButton("เลือก employees.xlsx")
        self.btn_pick_emp.clicked.connect(self._pick_employees_xlsx)

        self.mode = QtWidgets.QComboBox()
        self.mode.addItem("เน้น COND15 (ให้เกิน 15/เดือนชัวร์)", "COND15")
        self.mode.addItem("เน้นส่วนลด (0.xx เช่น 0.5)", "DISCOUNT")
        self.mode.addItem("ผสม (COND15 + ส่วนลด)", "MIX")

        self.n = QtWidgets.QSpinBox()
        self.n.setRange(10, 50000)
        self.n.setValue(500)

        self.year = QtWidgets.QSpinBox()
        self.year.setRange(2020, 2100)
        self.year.setValue(2026)

        self.month = QtWidgets.QSpinBox()
        self.month.setRange(1, 12)
        self.month.setValue(2)

        self.force_cond15 = QtWidgets.QSpinBox()
        self.force_cond15.setRange(0, 50000)
        self.force_cond15.setValue(25)

        self.p_monitor2 = QtWidgets.QDoubleSpinBox()
        self.p_monitor2.setRange(0.0, 1.0)
        self.p_monitor2.setSingleStep(0.05)
        self.p_monitor2.setValue(0.35)

        self.p_discount = QtWidgets.QDoubleSpinBox()
        self.p_discount.setRange(0.0, 1.0)
        self.p_discount.setSingleStep(0.05)
        self.p_discount.setValue(0.10)

        self.p_cond15_random = QtWidgets.QDoubleSpinBox()
        self.p_cond15_random.setRange(0.0, 1.0)
        self.p_cond15_random.setSingleStep(0.01)
        self.p_cond15_random.setValue(0.12)

        self.seed = QtWidgets.QSpinBox()
        self.seed.setRange(0, 999999)
        self.seed.setValue(42)

        self.force_first_days = QtWidgets.QSpinBox()
        self.force_first_days.setRange(0, 31)
        self.force_first_days.setValue(15)

        self.out = QtWidgets.QLineEdit("case_sample_500.xlsx")
        self.btn_pick_out = QtWidgets.QPushButton("เลือกที่เซฟ")
        self.btn_pick_out.clicked.connect(self._pick_out)

        self.sheet = QtWidgets.QLineEdit("Cases")

        self.btn_gen = QtWidgets.QPushButton("📄 Generate")
        self.btn_gen.clicked.connect(self._generate)

        self.btn_close = QtWidgets.QPushButton("ปิด")
        self.btn_close.clicked.connect(self.close)

        # layout
        form = QtWidgets.QFormLayout()

        row_master = QtWidgets.QHBoxLayout()
        row_master.addWidget(self.master_dir, 1)
        row_master.addWidget(self.btn_pick_master)
        form.addRow("master dir:", row_master)

        row_emp = QtWidgets.QHBoxLayout()
        row_emp.addWidget(self.employees_xlsx, 1)
        row_emp.addWidget(self.btn_pick_emp)
        form.addRow("employees.xlsx (optional):", row_emp)

        form.addRow("โหมด:", self.mode)
        form.addRow("จำนวนแถว:", self.n)

        row_ym = QtWidgets.QHBoxLayout()
        row_ym.addWidget(QtWidgets.QLabel("ปี"))
        row_ym.addWidget(self.year)
        row_ym.addSpacing(10)
        row_ym.addWidget(QtWidgets.QLabel("เดือน"))
        row_ym.addWidget(self.month)
        row_ym.addStretch(1)
        form.addRow("ปี/เดือน:", row_ym)

        form.addRow("บังคับ COND15 แถวแรก:", self.force_cond15)
        form.addRow("โอกาสสุ่ม COND15:", self.p_cond15_random)
        form.addRow("โอกาสมี Monitor(2):", self.p_monitor2)
        form.addRow("โอกาสมีส่วนลด (0.xx):", self.p_discount)
        form.addRow("กระจุกวันต้นเดือน 1..N:", self.force_first_days)
        form.addRow("Random seed:", self.seed)

        row_out = QtWidgets.QHBoxLayout()
        row_out.addWidget(self.out, 1)
        row_out.addWidget(self.btn_pick_out)
        form.addRow("output:", row_out)

        form.addRow("sheet:", self.sheet)

        btns = QtWidgets.QHBoxLayout()
        btns.addWidget(self.btn_close)
        btns.addStretch(1)
        btns.addWidget(self.btn_gen)

        hint = QtWidgets.QLabel(
            "หมายเหตุ:\n"
            "- วันที่ตรวจจะเป็น Date จริง (ไม่สลับวัน/เดือน)\n"
            "- ส่วนลดใช้ 0.xx (เช่น 0.5 = 50%)\n"
            "- ถ้าไฟล์ output ถูกเปิดใน Excel จะสร้างชื่อใหม่ _1 อัตโนมัติ"
        )
        hint.setStyleSheet("color:#555;")
        hint.setWordWrap(True)

        layout = QtWidgets.QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(hint)
        layout.addStretch(1)
        layout.addLayout(btns)

    def _pick_master_dir(self):
        d = QtWidgets.QFileDialog.getExistingDirectory(self, "เลือกโฟลเดอร์ master")
        if d:
            self.master_dir.setText(d)

    def _pick_employees_xlsx(self):
        p, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "เลือก employees.xlsx", "", "Excel (*.xlsx)"
        )
        if p:
            self.employees_xlsx.setText(p)

    def _pick_out(self):
        p, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "เลือกไฟล์ output", self.out.text(), "Excel (*.xlsx)"
        )
        if p:
            if not p.lower().endswith(".xlsx"):
                p = p + ".xlsx"
            self.out.setText(p)

    def _generate(self):
        try:
            cfg = GenConfig(
                master_dir=self.master_dir.text().strip() or "master",
                employees_xlsx=self.employees_xlsx.text().strip(),
                out=self.out.text().strip() or "case_sample_500.xlsx",
                sheet=self.sheet.text().strip() or "Cases",
                mode=str(self.mode.currentData()),
                n=int(self.n.value()),
                year=int(self.year.value()),
                month=int(self.month.value()),
                force_cond15=int(self.force_cond15.value()),
                p_cond15_random=float(self.p_cond15_random.value()),
                p_monitor2=float(self.p_monitor2.value()),
                p_discount=float(self.p_discount.value()),
                seed=int(self.seed.value()),
                force_first_days=int(self.force_first_days.value()),
            )

            out_path = generate_cases(cfg)

            QtWidgets.QMessageBox.information(
                self,
                "Done",
                f"✅ สร้างไฟล์สำเร็จ\n{out_path.resolve()}",
            )
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", str(e))


def run_ui():
    if QtWidgets is None:
        raise RuntimeError("PySide6 ยังไม่ได้ติดตั้ง: pip install PySide6")
    app = QtWidgets.QApplication([])
    dlg = GeneratorUI()
    dlg.show()
    app.exec()


def run_cli(args: argparse.Namespace):
    cfg = GenConfig(
        master_dir=args.master_dir,
        employees_xlsx=args.employees_xlsx,
        out=args.out,
        sheet=args.sheet,
        mode=args.mode,
        n=args.n,
        year=args.year,
        month=args.month,
        force_cond15=args.force_cond15,
        p_cond15_random=args.p_cond15_random,
        p_monitor2=args.p_monitor2,
        p_discount=args.p_discount,
        seed=args.seed,
        force_first_days=args.force_first_days,
    )
    out_path = generate_cases(cfg)
    print(f"✅ Generated: {out_path.resolve()} (rows={cfg.n})")


def main():
    args = parse_args()
    if args.cli:
        run_cli(args)
    else:
        run_ui()


if __name__ == "__main__":
    main()
