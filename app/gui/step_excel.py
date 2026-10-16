"""Step 2: load the Excel list and map columns to fields."""

from __future__ import annotations

import os
import tempfile

import openpyxl

try:
    import pymupdf as fitz_mod
except ImportError:
    import fitz as fitz_mod
from PySide6.QtCore import QUrl, Qt
from PySide6.QtGui import QDesktopServices, QImage, QPixmap
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFormLayout,
    QGroupBox, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPlainTextEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from app.core import excel_reader
from app.core.generator import (
    MODE_SEPARATE, choose_filename_field, generate, sanitize_filename,
)
from app.gui.state import AppState

MAX_PREVIEW_ROWS = 100


class CompareDialog(QDialog):
    """Side-by-side preview: template page vs generated sample page."""

    def __init__(self, template_pdf: str, generated_pdf: str, out_path: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Сравнение: шаблон и результат")
        self._out_path = out_path
        v = QVBoxLayout(self)
        row = QHBoxLayout()
        for title, path in (("Шаблон", template_pdf), ("Результат", generated_pdf)):
            col = QVBoxLayout()
            col.addWidget(QLabel(title))
            lbl = QLabel()
            lbl.setAlignment(Qt.AlignCenter)
            try:
                with fitz_mod.open(path) as doc:
                    pix = doc[0].get_pixmap(dpi=110)
                img = QImage(pix.samples, pix.width, pix.height,
                             pix.stride, QImage.Format_RGB888)
                pm = QPixmap.fromImage(img.copy())
                lbl.setPixmap(pm)
            except Exception as exc:
                lbl.setText(f"не удалось отрисовать:\n{exc}")
            col.addWidget(lbl, 1)
            row.addLayout(col)
        v.addLayout(row, 1)
        bb = QDialogButtonBox(QDialogButtonBox.Close)
        btn_open = bb.addButton("Открыть файл", QDialogButtonBox.AcceptRole)
        btn_open.clicked.connect(self._open_file)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self.resize(1100, 700)

    def _open_file(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile(self._out_path))


class StepExcel(QWidget):
    def __init__(self, state: AppState, go_next, go_back, parent=None):
        super().__init__(parent)
        self.state = state
        self._go_next = go_next
        self._go_back = go_back
        self._combos: dict[str, QComboBox] = {}
        self._constants = {}
        self._sheets: list[str] = []
        self._colmap_titles: dict[str, str] = {}
        self._name_col_title = None

        root = QVBoxLayout(self)
        top = QHBoxLayout()
        btn = QPushButton("Выбрать Excel-файл (.xlsx)…")
        btn.clicked.connect(self._open_excel)
        self.lbl_file = QLabel("Файл не выбран")
        self.cmb_sheet = QComboBox()
        self.cmb_sheet.setVisible(False)
        self.cmb_sheet.currentIndexChanged.connect(self._sheet_changed)
        top.addWidget(btn)
        btn_template = QPushButton("Сохранить пустой Excel…")
        btn_template.clicked.connect(self._save_blank_excel)
        top.addWidget(btn_template)
        top.addWidget(self.lbl_file, 1)
        top.addWidget(QLabel("Лист:"))
        top.addWidget(self.cmb_sheet)
        root.addLayout(top)

        root.addWidget(QLabel("Предпросмотр (первые строки):"))
        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        root.addWidget(self.table, 2)

        self.grp_map = QGroupBox("Соответствие полей шаблона столбцам Excel")
        self.map_layout = QFormLayout(self.grp_map)
        root.addWidget(self.grp_map)

        name_row = QHBoxLayout()
        name_row.addWidget(QLabel("Поле для имени файла:"))
        self.cmb_name = QComboBox()
        self.cmb_name.currentIndexChanged.connect(self._name_src_changed)
        name_row.addWidget(self.cmb_name)
        self.chk_first_word = QCheckBox("Взять только первое слово")
        self.chk_first_word.stateChanged.connect(self._first_word_changed)
        name_row.addWidget(self.chk_first_word)
        self.lbl_name_preview = QLabel("")
        name_row.addWidget(self.lbl_name_preview, 1)
        root.addLayout(name_row)

        root.addWidget(QLabel("Проверка данных:"))
        self.txt_issues = QPlainTextEdit()
        self.txt_issues.setReadOnly(True)
        self.txt_issues.setMaximumHeight(120)
        root.addWidget(self.txt_issues)

        row = QHBoxLayout()
        self.cmb_sample = QComboBox()
        self.btn_sample = QPushButton("Сгенерировать пример выбранной строки")
        self.btn_sample.clicked.connect(self._sample)
        self.btn_back = QPushButton("← Назад")
        self.btn_back.clicked.connect(self._go_back)
        self.btn_next = QPushButton("Далее →")
        self.btn_next.clicked.connect(self._next)
        row.addWidget(self.btn_back)
        row.addStretch(1)
        row.addWidget(self.cmb_sample)
        row.addWidget(self.btn_sample)
        row.addWidget(self.btn_next)
        root.addLayout(row)

    def _save_blank_excel(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить пустой Excel", "participants-template.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Участники"
        ws.append(["ФИО", "Номинация", "Управление"])
        try:
            wb.save(path)
        except Exception as exc:
            QMessageBox.critical(self, "Excel", f"Не удалось сохранить файл:\n{exc}")
        finally:
            wb.close()

    def _open_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Список Excel", "", "Excel (*.xlsx)")
        if not path:
            return
        try:
            sheets = excel_reader.sheet_names(path)
        except Exception as exc:
            QMessageBox.critical(self, "Excel", f"Не удалось прочитать файл:\n{exc}")
            return
        self._sheets = sheets
        if self.state.excel_path != path:
            self.state.colmap = {}
            self.state.filename_source = ("auto", None)
            self.state.headers = []
            self.state.rows = []
            self.state.sheet_name = None
            self._colmap_titles = {}
            self._name_col_title = None
        self.state.excel_path = path
        self.cmb_sheet.blockSignals(True)
        self.cmb_sheet.clear()
        self.cmb_sheet.addItems(sheets)
        self.cmb_sheet.setVisible(len(sheets) > 1)
        self.cmb_sheet.blockSignals(False)
        self._load_sheet(sheets[0])

    def _sheet_changed(self, idx: int):
        if 0 <= idx < len(self._sheets):
            self._load_sheet(self._sheets[idx])

    def _clear_sheet(self, sheet_name: str, headers=None):
        if self.state.rows:
            # Preserve the mapping by column title — indices into the old
            # header list are meaningless once a different schema is shown.
            self._colmap_titles = {name: self.state.headers[i]
                                   for name, i in self.state.colmap.items()
                                   if 0 <= i < len(self.state.headers)}
            kind, val = self.state.filename_source
            if kind == "col" and 0 <= val < len(self.state.headers):
                self._name_col_title = self.state.headers[val]
        self.state.sheet_name = sheet_name
        self.state.headers = headers or []
        self.state.rows = []
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.cmb_sample.clear()
        self.btn_next.setEnabled(False)
        self.btn_sample.setEnabled(False)
        self.lbl_file.setText(f"{os.path.basename(self.state.excel_path)} — 0 строк")
        self.txt_issues.setPlainText("Excel-файл пуст: заполните строки участников.")

    def _load_sheet(self, sheet_name: str):
        try:
            headers, rows = excel_reader.read_xlsx(self.state.excel_path, sheet_name)
        except Exception as exc:
            self._clear_sheet(sheet_name)
            QMessageBox.critical(self, "Excel", f"Не удалось прочитать лист:\n{exc}")
            return
        if not headers or not rows:
            self._clear_sheet(sheet_name, headers)
            QMessageBox.warning(self, "Excel", "Лист пуст или нет данных после строки заголовков.")
            return
        if self._colmap_titles:
            previous = self._colmap_titles
            self._colmap_titles = {}
        else:
            previous = {name: self.state.headers[i]
                        for name, i in self.state.colmap.items()
                        if 0 <= i < len(self.state.headers)}
        self.state.colmap = {name: headers.index(title) for name, title in previous.items()
                             if title in headers}
        self.state.sheet_name = sheet_name
        self.state.headers = headers
        self.state.rows = rows
        self.btn_next.setEnabled(True)
        self.btn_sample.setEnabled(True)
        self.lbl_file.setText(
            f"{os.path.basename(self.state.excel_path)} — {len(rows)} строк(и)")
        self._fill_table()
        self._build_mapping()
        self._revalidate()
        self.cmb_sample.clear()
        longest = max(range(len(rows)), key=lambda i: sum(len(v) for v in rows[i]))
        for i, row in enumerate(rows):
            label = f"Excel строка {i + 2}"
            if i == longest:
                label += " (длинные значения)"
            self.cmb_sample.addItem(label, i)

    def _fill_table(self):
        h, rows = self.state.headers, self.state.rows
        self.table.setColumnCount(len(h))
        self.table.setHorizontalHeaderLabels(h)
        self.table.setRowCount(min(len(rows), MAX_PREVIEW_ROWS))
        for r, row in enumerate(rows[:MAX_PREVIEW_ROWS]):
            for c, val in enumerate(row):
                self.table.setItem(r, c, QTableWidgetItem(val))

    # ---- mapping -----------------------------------------------------------
    def _build_mapping(self):
        while self.map_layout.rowCount():
            self.map_layout.removeRow(0)
        self._combos = {}
        self._constants = {}
        for f in self.state.fields:
            cmb = QComboBox()
            cmb.addItem("Сохранить исходный текст", -1)
            cmb.addItem("Постоянный текст", -2)
            for i, htext in enumerate(self.state.headers):
                cmb.addItem(htext, i)
            pre = self.state.colmap.get(f.name, self._guess_column(f.name))
            if f.source == "keep":
                cmb.setCurrentIndex(0)
            elif f.source == "constant":
                cmb.setCurrentIndex(1)
            elif pre >= 0 and pre < len(self.state.headers):
                cmb.setCurrentIndex(pre + 2)
            constant = QLineEdit(f.constant)
            constant.setPlaceholderText("Текст для всех дипломов")
            constant.setVisible(cmb.currentData() == -2)
            constant.textChanged.connect(lambda text, field=f: self._constant_changed(field, text))
            required = QCheckBox("Обязательно")
            required.setChecked(f.required or f.name.casefold() == "фио")
            required.toggled.connect(lambda checked, field=f: self._required_changed(field, checked))
            row = QWidget()
            layout = QHBoxLayout(row)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.addWidget(cmb, 2)
            layout.addWidget(constant, 2)
            layout.addWidget(required)
            cmb.currentIndexChanged.connect(lambda _i, name=f.name: self._map_changed(name))
            self._combos[f.name] = cmb
            self._constants[f.name] = constant
            self.map_layout.addRow(f.name, row)
            f.required = required.isChecked()
        for f in self.state.fields:
            self._map_changed(f.name)
        self._build_name_combo()

    def _guess_column(self, field_name: str) -> int:
        fl = field_name.lower()
        for i, htext in enumerate(self.state.headers):
            hl = htext.lower()
            if fl == hl or fl in hl or hl in fl:
                return i
        return -1

    def _constant_changed(self, field, text):
        field.constant = text
        self._revalidate()

    def _required_changed(self, field, checked):
        field.required = checked
        self._revalidate()

    def _map_changed(self, name: str):
        cmb = self._combos[name]
        col = cmb.currentData()
        field = next(f for f in self.state.fields if f.name == name)
        field.source = "keep" if col == -1 else "constant" if col == -2 else "column"
        self._constants[name].setVisible(col == -2)
        if col is not None and col >= 0:
            self.state.colmap[name] = col
        else:
            self.state.colmap.pop(name, None)
        self._update_name_preview()
        self._revalidate()

    # ---- filename source -----------------------------------------------------
    def _build_name_combo(self):
        """Populate the 'field for file name' combo: template fields first,
        then raw Excel columns. Keeps the previous choice when possible."""
        prev = self.state.filename_source
        if self._name_col_title is not None:
            if self._name_col_title in self.state.headers:
                prev = ("col", self.state.headers.index(self._name_col_title))
                self.state.filename_source = prev
            self._name_col_title = None
        self.cmb_name.blockSignals(True)
        self.cmb_name.clear()
        default_idx = choose_filename_field(self.state.fields) if self.state.fields else -1
        sel = 0
        for i, f in enumerate(self.state.fields):
            self.cmb_name.addItem(f"Поле: {f.name}", ("field", f.name))
            if prev == ("field", f.name):
                sel = self.cmb_name.count() - 1
            elif prev[0] == "auto" and i == default_idx:
                sel = self.cmb_name.count() - 1
        for i, htext in enumerate(self.state.headers):
            self.cmb_name.addItem(f"Столбец: {htext}", ("col", i))
            if prev == ("col", i):
                sel = self.cmb_name.count() - 1
        self.cmb_name.setCurrentIndex(sel)
        self.cmb_name.blockSignals(False)
        self.chk_first_word.blockSignals(True)
        self.chk_first_word.setChecked(self.state.filename_first_word)
        self.chk_first_word.blockSignals(False)
        self._name_src_changed(sel)

    def _first_word_changed(self, _state):
        self.state.filename_first_word = self.chk_first_word.isChecked()
        self._update_name_preview()

    def _name_src_changed(self, _idx: int):
        src = self.cmb_name.currentData()
        if src is not None:
            self.state.filename_source = src
        self._update_name_preview()
        self._revalidate()

    def _filename_col(self) -> int:
        kind, val = self.state.filename_source
        if kind == "field":
            return self.state.colmap.get(val, -1)
        if kind == "col":
            return val
        if self.state.fields:
            idx = choose_filename_field(self.state.fields)
            return self.state.colmap.get(self.state.fields[idx].name, -1)
        return -1

    def _update_name_preview(self):
        col = self._filename_col()
        if self.state.rows and 0 <= col:
            raw = self.state.rows[0][col] if col < len(self.state.rows[0]) else ""
            raw = raw.strip()
            if self.state.filename_first_word and raw:
                raw = raw.split()[0]
            ext = ".pptx" if self.state.template_type == "pptx" else ".pdf"
            self.lbl_name_preview.setText(f"Пример: {sanitize_filename(raw)}{ext}")
        else:
            self.lbl_name_preview.setText("Пример: —")

    # ---- validation ----------------------------------------------------------
    def current_problems(self) -> list[tuple[int, str]]:
        if not self.state.rows:
            return []
        cols = [self.state.colmap[f.name] for f in self.state.fields
                if f.required and f.source == "column" and f.name in self.state.colmap]
        problems = excel_reader.validate_rows(
            self.state.rows, cols, None, self.state.headers)
        for f in self.state.fields:
            if f.required and (f.source == "keep" or
                               f.source == "column" and f.name not in self.state.colmap or
                               f.source == "constant" and not f.constant.strip()):
                problems.append((0, f"обязательное поле «{f.name}» не заполнено"))
        return problems

    def _revalidate(self):
        problems = self.current_problems()
        if not self.state.rows:
            self.txt_issues.setPlainText("Excel-файл не загружен.")
            return
        kept = [f.name for f in self.state.fields if f.source == "keep"]
        lines = []
        if kept:
            lines.append("Сохранён исходный текст шаблона: " + ", ".join(kept))
        if problems:
            lines.append(f"Найдено проблем: {len(problems)}")
            for row_no, prob in problems[:200]:
                prefix = f"  строка {row_no}: " if row_no else "  "
                lines.append(prefix + prob)
            if len(problems) > 200:
                lines.append(f"  …и ещё {len(problems) - 200}")
        if not lines:
            lines.append("Проблем не найдено.")
        self.txt_issues.setPlainText("\n".join(lines))

    # ---- sample -----------------------------------------------------------
    def _sample(self):
        if not self.state.template_path or not self.state.rows:
            QMessageBox.warning(self, "Пример", "Нужны шаблон и Excel-файл.")
            return
        problems = self.current_problems()
        if any(r == 0 for r, _ in problems):
            QMessageBox.warning(self, "Пример", "Исправьте ошибки столбцов и обязательных полей.")
            return
        idx = self.cmb_sample.currentData()
        if idx is None:
            return
        row_problems = [p for r, p in problems if r == idx + 2]
        if row_problems:
            QMessageBox.warning(self, "Пример", "Строка содержит ошибки: " + "; ".join(row_problems))
            return
        tmpdir = tempfile.mkdtemp(prefix="diplom_sample_")
        try:
            rep = generate(
                self.state.template_path, self.state.template_type, self.state.fields,
                self.state.colmap, [self.state.rows[idx]], tmpdir, MODE_SEPARATE,
                self._filename_col(), set(),
                first_word=self.state.filename_first_word,
                slide_sel=self.state.slide_sel,
                signature=self.state.signature,
            )
        except Exception as exc:
            QMessageBox.critical(self, "Пример", f"Ошибка генерации:\n{exc}")
            return
        if not rep["created"]:
            reason = rep["skipped"][0][1] if rep["skipped"] else "неизвестная ошибка"
            QMessageBox.warning(self, "Пример", "Файл не был создан: " + reason)
            return
        name = rep["created"][0][1]
        out_path = os.path.join(tmpdir, name)
        note = ""
        if rep["review"]:
            note = "\n\nТребует проверки:\n" + "\n".join(rep["review"][0][1])
        if self.state.template_type == "pdf":
            CompareDialog(self.state.template_path, out_path, out_path, self).exec()
            if note:
                QMessageBox.information(self, "Пример", note.strip())
        else:
            QDesktopServices.openUrl(QUrl.fromLocalFile(out_path))
            if note:
                QMessageBox.information(self, "Пример", "Файл создан." + note)

    def _next(self):
        if not self.state.rows:
            QMessageBox.warning(self, "Excel", "Сначала выберите Excel-файл.")
            return
        self.state.filename_col = self._filename_col()
        # persist naming choice into the template's .mapping.json
        if self.state.template_path:
            try:
                from app.core import mapping_io
                mapping_io.save_mapping(
                    self.state.template_path, self.state.template_type,
                    self.state.fields,
                    filename_source=self.state.filename_source,
                    filename_first_word=self.state.filename_first_word,
                    slide_sel=self.state.slide_sel,
                    signature=self.state.signature,
                )
            except Exception:
                pass
        self._go_next()

    def on_enter(self):
        """Called when the wizard navigates to this step."""
        if self.state.headers:
            self._build_mapping()
            self._revalidate()
