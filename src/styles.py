"""
Modern UI Styles for Salary Calculator Application
"""

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QLabel, QPushButton, QGraphicsDropShadowEffect, QStyle

# Modern color palette
COLORS = {
    'primary': '#0f766e',
    'primary_hover': '#115e59',
    'primary_light': '#ccfbf1',
    'secondary': '#2563eb',
    'secondary_hover': '#1d4ed8',
    'danger': '#dc2626',
    'danger_hover': '#b91c1c',
    'warning': '#d97706',
    'success': '#16a34a',
    'background': '#f3f6fb',
    'surface': '#ffffff',
    'surface_alt': '#f8fafc',
    'text_primary': '#0f172a',
    'text_secondary': '#475569',
    'border': '#dbe4f0',
    'border_strong': '#c3d2e6',
    'shadow': 'rgba(15, 23, 42, 0.08)',
}

# Modern stylesheet
MODERN_STYLE = f"""
/* Main Window */
QMainWindow {{
    background-color: {COLORS['background']};
    color: {COLORS['text_primary']};
    font-family: 'Segoe UI', 'Microsoft Sans Serif', sans-serif;
    font-size: 14px;
}}

QDialog {{
    background-color: {COLORS['background']};
    color: {COLORS['text_primary']};
    font-family: 'Segoe UI', 'Microsoft Sans Serif', sans-serif;
    font-size: 14px;
}}

QScrollArea {{
    background-color: {COLORS['background']};
    border: none;
}}

QScrollArea > QWidget {{
    background-color: {COLORS['background']};
}}

/* Group Box - Modern Card Style */
QGroupBox {{
    background-color: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px;
    margin-top: 10px;
    padding-top: 18px;
    font-weight: 600;
    font-size: 15px;
    color: {COLORS['text_primary']};
}}

QGroupBox::title {{
    subcontrol-origin: margin;
    left: 16px;
    padding: 0 8px 0 8px;
    background-color: {COLORS['surface']};
    color: {COLORS['text_primary']};
    font-weight: 600;
}}

/* Buttons - Modern Style */
QPushButton {{
    background-color: qlineargradient(
        x1: 0, y1: 0, x2: 0, y2: 1,
        stop: 0 #149286,
        stop: 1 {COLORS['primary']}
    );
    color: white;
    border: 1px solid rgba(15, 23, 42, 0.10);
    border-radius: 8px;
    padding: 10px 18px;
    font-weight: 600;
    font-size: 14px;
    min-height: 40px;
    text-align: center;
}}

QPushButton:hover {{
    background-color: qlineargradient(
        x1: 0, y1: 0, x2: 0, y2: 1,
        stop: 0 #1aa092,
        stop: 1 {COLORS['primary_hover']}
    );
    border: 1px solid rgba(15, 23, 42, 0.18);
}}

QPushButton:pressed {{
    background-color: {COLORS['primary_hover']};
    padding-top: 11px;
    padding-bottom: 9px;
}}

QPushButton:disabled {{
    background-color: {COLORS['border']};
    color: {COLORS['text_secondary']};
}}

/* Secondary Button */
QPushButton[class="secondary"] {{
    background-color: {COLORS['surface']};
    color: {COLORS['secondary']};
    border: 1px solid {COLORS['border_strong']};
}}

QPushButton[class="secondary"]:hover {{
    background-color: #eff6ff;
    border: 1px solid {COLORS['secondary']};
}}

/* Success Button */
QPushButton[class="success"] {{
    background-color: qlineargradient(
        x1: 0, y1: 0, x2: 0, y2: 1,
        stop: 0 #22c55e,
        stop: 1 {COLORS['success']}
    );
}}

QPushButton[class="success"]:hover {{
    background-color: qlineargradient(
        x1: 0, y1: 0, x2: 0, y2: 1,
        stop: 0 #1fb155,
        stop: 1 #15803d
    );
}}

/* Danger Button */
QPushButton[class="danger"] {{
    background-color: {COLORS['danger']};
}}

QPushButton[class="danger"]:hover {{
    background-color: {COLORS['danger_hover']};
}}

/* Line Edit - Modern Input */
QLineEdit {{
    background-color: {COLORS['surface']};
    border: 1px solid {COLORS['border_strong']};
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 14px;
    color: {COLORS['text_primary']};
    text-align: left;
}}

QLineEdit:focus {{
    border-color: {COLORS['primary']};
    outline: none;
}}

QLineEdit:disabled {{
    background-color: {COLORS['background']};
    color: {COLORS['text_secondary']};
}}

QTextEdit {{
    background-color: {COLORS['surface']};
    border: 1px solid {COLORS['border_strong']};
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 14px;
    color: {COLORS['text_primary']};
    selection-background-color: {COLORS['primary_light']};
    selection-color: {COLORS['text_primary']};
}}

QTextEdit:focus {{
    border-color: {COLORS['primary']};
}}

/* Combo Box */
QComboBox {{
    background-color: {COLORS['surface']};
    border: 1px solid {COLORS['border_strong']};
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 14px;
    color: {COLORS['text_primary']};
    min-height: 20px;
    text-align: left;
}}

QComboBox:focus {{
    border-color: {COLORS['primary']};
}}

QComboBox::drop-down {{
    border: none;
    width: 30px;
}}

QComboBox::down-arrow {{
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 5px solid {COLORS['text_secondary']};
    margin-right: 10px;
}}

QComboBox QAbstractItemView {{
    background-color: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px;
    selection-background-color: {COLORS['primary_light']};
    selection-color: {COLORS['text_primary']};
    padding: 4px;
}}

/* Date Edit */
QDateEdit {{
    background-color: {COLORS['surface']};
    border: 1px solid {COLORS['border_strong']};
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 14px;
    color: {COLORS['text_primary']};
    text-align: left;
}}

QDateEdit:focus {{
    border-color: {COLORS['primary']};
}}

/* Table Widget */
QTableWidget {{
    background-color: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px;
    gridline-color: {COLORS['border']};
    selection-background-color: {COLORS['primary_light']};
    selection-color: {COLORS['text_primary']};
    alternate-background-color: {COLORS['surface_alt']};
}}

QTableWidget::item {{
    padding: 8px 12px;
    border: none;
    border-bottom: 1px solid {COLORS['border']};
}}

QTableWidget::item:selected {{
    background-color: {COLORS['primary_light']};
    color: {COLORS['text_primary']};
}}

QTableCornerButton::section {{
    background-color: {COLORS['surface_alt']};
    border: none;
    border-bottom: 1px solid {COLORS['border_strong']};
    border-right: 1px solid {COLORS['border']};
}}

QHeaderView::section {{
    background-color: {COLORS['surface_alt']};
    color: {COLORS['text_primary']};
    font-weight: 600;
    padding: 10px 8px;
    border: none;
    border-bottom: 1px solid {COLORS['border_strong']};
    border-right: 1px solid {COLORS['border']};
}}

QHeaderView::section:last {{
    border-right: none;
}}

/* Labels */
QLabel {{
    color: {COLORS['text_primary']};
    font-weight: 500;
}}

QLabel[class="header"] {{
    font-size: 20px;
    font-weight: 700;
    color: {COLORS['text_primary']};
    margin: 6px 0;
}}

QLabel[class="subheader"] {{
    font-size: 16px;
    font-weight: 600;
    color: {COLORS['text_primary']};
    margin: 4px 0;
}}

QLabel[class="helper"] {{
    font-size: 13px;
    color: {COLORS['text_secondary']};
    padding: 0 2px 4px 2px;
}}

QFrame[class="metric-card"] {{
    background-color: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px;
}}

QLabel[class="metric-value"] {{
    font-size: 18px;
    font-weight: 700;
    color: {COLORS['text_primary']};
}}

QLabel[class="metric-label"] {{
    font-size: 12px;
    font-weight: 500;
    color: {COLORS['text_secondary']};
}}

QLabel[class="status"] {{
    font-size: 13px;
    color: {COLORS['text_secondary']};
    padding: 8px 12px;
    background-color: {COLORS['surface_alt']};
    border-radius: 6px;
    border: 1px solid {COLORS['border']};
}}

QLabel[class="status-success"] {{
    color: {COLORS['success']};
    background-color: #dcfce7;
}}

QLabel[class="status-error"] {{
    color: {COLORS['danger']};
    background-color: #fee2e2;
}}

/* Status Bar */
QStatusBar {{
    background-color: {COLORS['surface']};
    border-top: 1px solid {COLORS['border']};
    color: {COLORS['text_secondary']};
}}

QTabWidget::pane {{
    background-color: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px;
    top: -1px;
}}

QTabBar::tab {{
    background-color: {COLORS['surface_alt']};
    color: {COLORS['text_secondary']};
    border: 1px solid {COLORS['border']};
    border-bottom: none;
    padding: 8px 14px;
    margin-right: 2px;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    font-weight: 600;
}}

QTabBar::tab:selected {{
    background-color: {COLORS['surface']};
    color: {COLORS['text_primary']};
    border-color: {COLORS['border_strong']};
}}

QTabBar::tab:hover {{
    color: {COLORS['secondary']};
}}

QListWidget {{
    background-color: {COLORS['surface']};
    border: 1px solid {COLORS['border_strong']};
    border-radius: 8px;
    color: {COLORS['text_primary']};
    selection-background-color: {COLORS['primary_light']};
    selection-color: {COLORS['text_primary']};
}}

QProgressBar {{
    background-color: #e7eef8;
    border: none;
    border-radius: 4px;
}}

QProgressBar::chunk {{
    background-color: {COLORS['primary']};
    border-radius: 4px;
}}

/* Scroll Bar */
QScrollBar:vertical {{
    background-color: {COLORS['background']};
    width: 12px;
    border-radius: 6px;
}}

QScrollBar::handle:vertical {{
    background-color: {COLORS['border']};
    border-radius: 6px;
    min-height: 20px;
}}

QScrollBar::handle:vertical:hover {{
    background-color: {COLORS['text_secondary']};
}}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}

/* Tool Tip */
QToolTip {{
    background-color: {COLORS['text_primary']};
    color: {COLORS['surface']};
    border: none;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 12px;
}}

/* Calendar Widget */
QCalendarWidget {{
    background-color: {COLORS['surface']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px;
}}

QCalendarWidget QToolButton {{
    background-color: {COLORS['surface']};
    color: {COLORS['text_primary']};
    border: none;
    border-radius: 4px;
    padding: 4px 8px;
    font-weight: 500;
}}

QCalendarWidget QToolButton:hover {{
    background-color: {COLORS['primary_light']};
}}

QCalendarWidget QAbstractItemView:enabled {{
    background-color: {COLORS['surface']};
    color: {COLORS['text_primary']};
    selection-background-color: {COLORS['primary']};
    selection-color: white;
}}
"""

def get_style_sheet():
    """Return the modern style sheet"""
    return MODERN_STYLE


def apply_light_app_palette(app):
    """Keep the app's light theme stable on macOS dark appearance."""
    if hasattr(app.styleHints(), "setColorScheme") and hasattr(Qt, "ColorScheme"):
        app.styleHints().setColorScheme(Qt.ColorScheme.Light)

    palette = app.palette()
    palette.setColor(QPalette.Window, QColor(COLORS['background']))
    palette.setColor(QPalette.WindowText, QColor(COLORS['text_primary']))
    palette.setColor(QPalette.Base, QColor(COLORS['surface']))
    palette.setColor(QPalette.AlternateBase, QColor(COLORS['surface_alt']))
    palette.setColor(QPalette.ToolTipBase, QColor(COLORS['text_primary']))
    palette.setColor(QPalette.ToolTipText, QColor(COLORS['surface']))
    palette.setColor(QPalette.Text, QColor(COLORS['text_primary']))
    palette.setColor(QPalette.Button, QColor(COLORS['surface']))
    palette.setColor(QPalette.ButtonText, QColor(COLORS['text_primary']))
    palette.setColor(QPalette.BrightText, QColor("#ffffff"))
    palette.setColor(QPalette.Highlight, QColor(COLORS['primary_light']))
    palette.setColor(QPalette.HighlightedText, QColor(COLORS['text_primary']))

    disabled_text = QColor(COLORS['text_secondary'])
    palette.setColor(QPalette.Disabled, QPalette.WindowText, disabled_text)
    palette.setColor(QPalette.Disabled, QPalette.Text, disabled_text)
    palette.setColor(QPalette.Disabled, QPalette.ButtonText, disabled_text)
    palette.setColor(QPalette.Disabled, QPalette.Button, QColor(COLORS['background']))
    palette.setColor(QPalette.Disabled, QPalette.Base, QColor(COLORS['background']))

    app.setPalette(palette)


def _icon_for_button(style, text: str):
    text = (text or "").strip().lower()
    mapping = [
        (("เลือกไฟล์", "โหลด"), QStyle.SP_DialogOpenButton),
        (("เปิดโฟลเดอร์",), QStyle.SP_DirOpenIcon),
        (("คำนวณ", "ส่งออก", "บันทึก"), QStyle.SP_DialogSaveButton),
        (("เพิ่มพนักงาน", "เพิ่ม "), QStyle.SP_FileDialogNewFolder),
        (("จัดการพนักงาน",), QStyle.SP_FileDialogDetailedView),
        (("สรุป",), QStyle.SP_FileDialogInfoView),
        (("ล้าง",), QStyle.SP_DialogResetButton),
        (("ไม่รับเงิน", "ไม่คิด"), QStyle.SP_DialogCancelButton),
        (("โรงพยาบาล", "รพ"), QStyle.SP_DriveHDIcon),
        (("ปิด", "ยกเลิก"), QStyle.SP_DialogCloseButton),
        (("export",), QStyle.SP_DialogSaveButton),
        (("ลบ",), QStyle.SP_TrashIcon),
    ]

    for keywords, icon_type in mapping:
        if any(keyword in text for keyword in keywords):
            return style.standardIcon(icon_type)
    return None


def _apply_button_depth(button: QPushButton):
    if button.property("depth_applied"):
        return

    shadow = QGraphicsDropShadowEffect(button)
    shadow.setBlurRadius(18)
    shadow.setOffset(0, 4)
    shadow.setColor(QColor(15, 23, 42, 32))
    button.setGraphicsEffect(shadow)
    button.setProperty("depth_applied", True)


def apply_modern_style(widget):
    """Apply modern styling to a widget and its children"""
    widget.setStyleSheet(MODERN_STYLE)
    style = widget.style()

    for child in widget.findChildren(QPushButton):
        text = child.text()
        if "เพิ่ม" in text:
            child.setProperty("class", "success")
        elif "ลบ" in text or "danger" in text.lower():
            child.setProperty("class", "danger")
        elif "ล้าง" in text or "เปิด" in text or "จัดการ" in text or "สรุป" in text:
            child.setProperty("class", "secondary")

        icon = _icon_for_button(style, text)
        if icon is not None:
            child.setIcon(icon)
            child.setIconSize(QSize(18, 18))
            child.setCursor(Qt.PointingHandCursor)

        _apply_button_depth(child)
        child.style().unpolish(child)
        child.style().polish(child)

    for child in widget.findChildren(QLabel):
        if "สรุป" in child.text() or "📌" in child.text():
            child.setProperty("class", "header")
        elif "สถานะ" in child.text().lower():
            child.setProperty("class", "status")
