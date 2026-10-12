"""Main window: 3-step wizard."""

from __future__ import annotations

from PySide6.QtWidgets import QLabel, QMainWindow, QStackedWidget, QVBoxLayout, QWidget

from app.gui.state import AppState
from app.gui.step_excel import StepExcel
from app.gui.step_generate import StepGenerate
from app.gui.step_template import StepTemplate


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Генератор дипломов и грамот")
        self.resize(1100, 750)

        self.state = AppState()
        central = QWidget()
        v = QVBoxLayout(central)
        self.header = QLabel("")
        self.header.setStyleSheet("font-size: 16px; font-weight: bold; padding: 4px;")
        v.addWidget(self.header)
        self.stack = QStackedWidget()
        v.addWidget(self.stack, 1)
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
        self.header.setText(
            ["Шаг 1. Шаблон и поля",
             "Шаг 2. Список Excel и соответствие столбцов",
             "Шаг 3. Генерация"][idx]
        )
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
