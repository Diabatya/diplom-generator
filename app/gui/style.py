"""Application stylesheet in the FNS palette: light background, dark-blue
primary elements, light-blue accents.

Font sizes use pt so the whole UI scales correctly under Windows display
scaling (125 % / 150 %) together with Qt's DPI awareness.
"""

from __future__ import annotations

from PySide6.QtWidgets import QApplication

COLOR_PRIMARY = "#0B3D91"      # dark blue
COLOR_PRIMARY_HOVER = "#1557B8"
COLOR_ACCENT = "#3E8FD5"       # light blue
COLOR_BG = "#F4F7FB"
COLOR_SURFACE = "#FFFFFF"
COLOR_TEXT = "#1A2733"
COLOR_BORDER = "#C9D6E4"
COLOR_DISABLED = "#9AAABD"

_QSS = f"""
QWidget {{
    background: {COLOR_BG};
    color: {COLOR_TEXT};
    font-size: 10pt;
}}

QLabel {{
    background: transparent;
}}

QLabel[role="hint"] {{
    color: #55677C;
}}

QGroupBox {{
    background: {COLOR_SURFACE};
    border: 1px solid {COLOR_BORDER};
    border-radius: 6px;
    margin-top: 14px;
    padding-top: 10px;
    font-weight: bold;
    color: {COLOR_PRIMARY};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 10px;
    padding: 0 4px;
}}

QPushButton {{
    background: {COLOR_PRIMARY};
    color: white;
    border: none;
    border-radius: 5px;
    padding: 6px 14px;
    font-size: 10pt;
    min-height: 20px;
}}
QPushButton:hover {{ background: {COLOR_PRIMARY_HOVER}; }}
QPushButton:pressed {{ background: #093372; }}
QPushButton:disabled {{
    background: #D7E0EA;
    color: {COLOR_DISABLED};
}}
QPushButton[secondary="true"] {{
    background: {COLOR_SURFACE};
    color: {COLOR_PRIMARY};
    border: 1px solid {COLOR_ACCENT};
}}
QPushButton[secondary="true"]:hover {{ background: #E7F0FA; }}
QPushButton[secondary="true"]:disabled {{
    color: {COLOR_DISABLED};
    border-color: {COLOR_BORDER};
}}
QPushButton:checked {{
    background: {COLOR_ACCENT};
}}

QLineEdit, QComboBox, QDoubleSpinBox, QPlainTextEdit {{
    background: {COLOR_SURFACE};
    border: 1px solid {COLOR_BORDER};
    border-radius: 4px;
    padding: 4px 6px;
    selection-background-color: {COLOR_ACCENT};
}}
QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus, QPlainTextEdit:focus {{
    border: 1px solid {COLOR_ACCENT};
}}
QComboBox::drop-down {{
    border: none;
    width: 22px;
}}
QComboBox::down-arrow {{
    image: none;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 5px solid {COLOR_PRIMARY};
    margin-right: 6px;
}}

QListWidget {{
    background: {COLOR_SURFACE};
    border: 1px solid {COLOR_BORDER};
    border-radius: 4px;
}}
QListWidget::item {{
    padding: 3px 6px;
}}
QListWidget::item:selected {{
    background: {COLOR_ACCENT};
    color: white;
}}
QListWidget::item:hover:!selected {{
    background: #E7F0FA;
}}

QTableWidget, QTableView {{
    background: {COLOR_SURFACE};
    gridline-color: {COLOR_BORDER};
    border: 1px solid {COLOR_BORDER};
}}
QHeaderView::section {{
    background: {COLOR_PRIMARY};
    color: white;
    padding: 4px 6px;
    border: none;
    font-weight: bold;
}}

QProgressBar {{
    background: {COLOR_SURFACE};
    border: 1px solid {COLOR_BORDER};
    border-radius: 5px;
    text-align: center;
    height: 16px;
}}
QProgressBar::chunk {{
    background: {COLOR_ACCENT};
    border-radius: 4px;
}}

QCheckBox::indicator, QRadioButton::indicator {{
    width: 15px;
    height: 15px;
}}

QToolTip {{
    background: {COLOR_PRIMARY};
    color: white;
    border: none;
    padding: 4px;
}}

QSplitter::handle {{
    background: {COLOR_BORDER};
    width: 2px;
}}

QDialog {{ background: {COLOR_BG}; }}
QMessageBox QLabel {{ font-size: 10pt; }}
"""


def apply(app: QApplication) -> None:
    app.setStyleSheet(_QSS)
