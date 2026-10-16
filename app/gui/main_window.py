"""Main window: wizard with an FNS-styled header (emblem, step
breadcrumb) and a developer footer."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QMainWindow, QStackedWidget,
    QVBoxLayout, QWidget,
)

from app import resources
from app.gui.state import AppState
from app.gui.step_excel import StepExcel
from app.gui.step_generate import StepGenerate
from app.gui.step_template import StepTemplate

_STEPS = ("Шаблон и поля", "Участники (Excel)", "Проверка и генерация")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Генератор дипломов и грамот")
        self.resize(1160, 800)

        self.state = AppState()
        central = QWidget()
        v = QVBoxLayout(central)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        # ---- header: emblem + titles + breadcrumb ----------------------
        header = QFrame()
        header.setStyleSheet(
            "background: #FFFFFF; border-bottom: 1px solid #C9D6E4;")
        hh = QHBoxLayout(header)
        hh.setContentsMargins(14, 8, 14, 8)
        self.logo = QLabel()
        emblem = resources.emblem_path()
        pix = QPixmap(emblem)
        if not pix.isNull():
            self.logo.setPixmap(pix.scaled(
                44, 46, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        self.logo.setFixedSize(48, 50)
        hh.addWidget(self.logo)
        titles = QVBoxLayout()
        self.lbl_title = QLabel("Генератор дипломов и грамот")
        self.lbl_title.setStyleSheet(
            "font-size: 15pt; font-weight: bold; color: #0B3D91;")
        self.lbl_breadcrumb = QLabel("")
        self.lbl_breadcrumb.setTextFormat(Qt.RichText)
        titles.addWidget(self.lbl_title)
        titles.addWidget(self.lbl_breadcrumb)
        hh.addLayout(titles, 1)
        v.addWidget(header)

        # ---- wizard pages ----------------------------------------------
        self.stack = QStackedWidget()
        body = QWidget()
        bv = QVBoxLayout(body)
        bv.setContentsMargins(14, 10, 14, 8)
        bv.addWidget(self.stack, 1)
        v.addWidget(body, 1)

        # ---- footer -----------------------------------------------------
        footer = QFrame()
        footer.setStyleSheet(
            "background: #FFFFFF; border-top: 1px solid #C9D6E4;")
        fh = QHBoxLayout(footer)
        fh.setContentsMargins(14, 4, 14, 4)
        dev = QLabel("Разработчик — УФНС России по Ставропольскому краю")
        dev.setStyleSheet("color: #55677C; font-size: 9pt;")
        fh.addWidget(dev)
        fh.addStretch(1)
        v.addWidget(footer)

        self.setCentralWidget(central)

        self.step1 = StepTemplate(self.state, go_next=lambda: self._go(1))
        self.step2 = StepExcel(self.state, go_next=lambda: self._go(2), go_back=lambda: self._go(0))
        self.step3 = StepGenerate(self.state, go_back=lambda: self._go(1), restart=self._restart)
        self.stack.addWidget(self.step1)
        self.stack.addWidget(self.step2)
        self.stack.addWidget(self.step3)
        self._go(0)

    def _go(self, idx: int):
        self.stack.setCurrentIndex(idx)
        parts = []
        for i, name in enumerate(_STEPS):
            if i == idx:
                parts.append(f'<b><span style="color:#0B3D91">{i + 1}. {name}</span></b>')
            else:
                parts.append(f'<span style="color:#9AAABD">{i + 1}. {name}</span>')
        self.lbl_breadcrumb.setText("&nbsp;&nbsp;→&nbsp;&nbsp;".join(parts))
        if idx == 1:
            self.step2.on_enter()
        if idx == 2:
            self.step3.on_enter()

    def closeEvent(self, event):
        self.step3.stop_worker()
        super().closeEvent(event)

    def _restart(self):
        self.state.excel_path = None
        self.state.headers = []
        self.state.rows = []
        self.state.colmap = {}
        self._go(0)
