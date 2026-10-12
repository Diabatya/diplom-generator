"""Step 3: choose output folder, run batch generation with progress."""

from __future__ import annotations

import copy
import os

from PySide6.QtCore import QThread, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QDialog, QDialogButtonBox, QFileDialog, QHBoxLayout, QLabel,
    QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QRadioButton,
    QVBoxLayout, QWidget,
)

from app.core import excel_reader, pptx_engine
from app.core.models import TEMPLATE_PPTX
from app.core.generator import (
    EXISTS_OVERWRITE, EXISTS_SKIP, EXISTS_SUFFIX,
    MODE_COMBINED, MODE_SEPARATE, generate, planned_filenames,
)
from app.gui.state import AppState


class _Worker(QThread):
    sig_progress = Signal(int, int, str)
    sig_done = Signal(dict)

    def __init__(self, kwargs, parent=None):
        super().__init__(parent)
        self._kwargs = kwargs
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        self._kwargs["progress_cb"] = lambda i, t, lbl: self.sig_progress.emit(i, t, lbl)
        self._kwargs["cancel_cb"] = lambda: self._cancel
        try:
            rep = generate(**self._kwargs)
        except Exception as exc:
            rep = {"created": [], "skipped": [(0, str(exc))], "review": [],
                   "cancelled": False, "fatal": str(exc), "report_path": None}
        self.sig_done.emit(rep)


class ModeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Режим вывода")
        v = QVBoxLayout(self)
        v.addWidget(QLabel("Как сохранить результаты?"))
        self.rb_sep = QRadioButton("Отдельный файл на каждую строку")
        self.rb_comb = QRadioButton("Один файл со всеми записями")
        self.rb_comb.setChecked(True)
        v.addWidget(self.rb_sep)
        v.addWidget(self.rb_comb)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

    @property
    def mode(self):
        return MODE_COMBINED if self.rb_comb.isChecked() else MODE_SEPARATE


class ExistsDialog(QDialog):
    """One-shot resolution for output files that already exist."""

    def __init__(self, count: int, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Файлы уже существуют")
        v = QVBoxLayout(self)
        v.addWidget(QLabel(
            f"В папке результатов уже есть {count} файл(ов) с такими же именами.\n"
            "Что сделать со всеми совпадениями?"))
        self.rb_over = QRadioButton("Перезаписать все")
        self.rb_skip = QRadioButton("Пропустить существующие (создать только новые)")
        self.rb_suf = QRadioButton("Добавить суффикс «_2», «_3» к именам")
        self.rb_suf.setChecked(True)
        v.addWidget(self.rb_over)
        v.addWidget(self.rb_skip)
        v.addWidget(self.rb_suf)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

    @property
    def exists_mode(self):
        if self.rb_over.isChecked():
            return EXISTS_OVERWRITE
        if self.rb_skip.isChecked():
            return EXISTS_SKIP
        return EXISTS_SUFFIX


class StepGenerate(QWidget):
    def __init__(self, state: AppState, go_back, restart, parent=None):
        super().__init__(parent)
        self.state = state
        self._go_back = go_back
        self._restart = restart
        self._worker: _Worker | None = None

        root = QVBoxLayout(self)
        top = QHBoxLayout()
        self.btn_dir = QPushButton("Выбрать папку для результатов…")
        self.btn_dir.clicked.connect(self._pick_dir)
        self.lbl_dir = QLabel("Папка не выбрана")
        top.addWidget(self.btn_dir)
        top.addWidget(self.lbl_dir, 1)
        root.addLayout(top)

        self.lbl_summary = QLabel("")
        root.addWidget(self.lbl_summary)

        row = QHBoxLayout()
        self.btn_back = QPushButton("← Назад")
        self.btn_back.clicked.connect(self._go_back)
        self.btn_start = QPushButton("Начать генерацию")
        self.btn_start.clicked.connect(self._start)
        self.btn_cancel = QPushButton("Отменить генерацию")
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._cancel)
        row.addWidget(self.btn_back)
        row.addStretch(1)
        row.addWidget(self.btn_start)
        row.addWidget(self.btn_cancel)
        root.addLayout(row)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        root.addWidget(self.progress)
        self.lbl_status = QLabel("")
        root.addWidget(self.lbl_status)

        self.txt_log = QPlainTextEdit()
        self.txt_log.setReadOnly(True)
        root.addWidget(self.txt_log, 1)

        tail = QHBoxLayout()
        self.btn_open_dir = QPushButton("Открыть папку с результатами")
        self.btn_open_dir.setEnabled(False)
        self.btn_open_dir.clicked.connect(self._open_dir)
        self.btn_restart = QPushButton("В начало")
        self.btn_restart.clicked.connect(self._restart)
        tail.addStretch(1)
        tail.addWidget(self.btn_open_dir)
        tail.addWidget(self.btn_restart)
        root.addLayout(tail)

    def on_enter(self):
        self.lbl_summary.setText(
            f"Шаблон: {os.path.basename(self.state.template_path or '—')}   |   "
            f"Строк в Excel: {len(self.state.rows)}   |   "
            f"Полей: {len(self.state.fields)}"
        )

    def _pick_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Папка для результатов")
        if d:
            self.state.out_dir = d
            self.lbl_dir.setText(d)

    def _start(self):
        if self._worker is not None:
            return
        if not self.state.out_dir:
            QMessageBox.warning(self, "Папка", "Выберите папку для результатов.")
            return
        if not self.state.template_path or not self.state.rows:
            QMessageBox.warning(self, "Данные", "Нужны шаблон и Excel-файл.")
            return
        if self.state.template_type == TEMPLATE_PPTX:
            try:
                issues = pptx_engine.required_placement_issues(
                    pptx_engine.open_template(self.state.template_path),
                    self.state.fields, self.state.slide_sel)
            except Exception as exc:
                issues = [f"Не удалось проверить шаблон: {exc}"]
            if issues:
                QMessageBox.warning(self, "Привязки", "\n".join(issues))
                return

        dlg = ModeDialog(self)
        if dlg.exec() != QDialog.Accepted:
            return
        mode = dlg.mode

        # pre-generation validation
        cols = [self.state.colmap[f.name] for f in self.state.fields
                if f.required and f.source == "column" and f.name in self.state.colmap]
        problems = excel_reader.validate_rows(
            self.state.rows, cols,
            self.state.filename_col if mode == MODE_SEPARATE else None,
            self.state.headers)
        problems.extend((0, f"обязательное поле «{f.name}» не заполнено")
                        for f in self.state.fields if f.required and
                        (f.source == "keep" or
                         f.source == "column" and f.name not in self.state.colmap or
                         f.source == "constant" and not f.constant.strip()))
        columns = [p for r, p in problems if r == 0]
        if columns:
            QMessageBox.warning(self, "Проверка данных",
                                "Исправьте ошибки перед генерацией:\n" + "\n".join(columns))
            return
        skip: set[int] = set()
        if problems:
            preview = "\n".join(
                (f"  строка {r}: {p}" if r else f"  {p}") for r, p in problems[:20])
            if len(problems) > 20:
                preview += f"\n  …и ещё {len(problems) - 20}"
            box = QMessageBox(self)
            box.setWindowTitle("Проверка данных")
            box.setIcon(QMessageBox.Warning)
            box.setText(f"Найдено проблемных строк: {len({r for r, _ in problems if r})}\n{preview}")
            btn_skip = box.addButton("Пропустить эти строки и продолжить", QMessageBox.AcceptRole)
            box.addButton("Отмена (исправлю Excel)", QMessageBox.RejectRole)
            box.exec()
            if box.clickedButton() is not btn_skip:
                return
            skip = {r - 2 for r, _ in problems if r}

        # collision check: files with the same names already in out_dir
        planned = planned_filenames(
            self.state.template_path, self.state.template_type,
            self.state.rows, mode, self.state.filename_col, skip,
            first_word=self.state.filename_first_word,
            out_dir=self.state.out_dir)
        on_disk = ({n.casefold() for n in os.listdir(self.state.out_dir)}
                   if os.path.isdir(self.state.out_dir) else set())
        existing = [n for n in planned if n.casefold() in on_disk]
        exists_mode = EXISTS_OVERWRITE
        if existing:
            dlg2 = ExistsDialog(len(existing), self)
            if dlg2.exec() != QDialog.Accepted:
                return
            exists_mode = dlg2.exists_mode

        self.btn_start.setEnabled(False)
        self.btn_back.setEnabled(False)
        self.btn_restart.setEnabled(False)
        self.btn_dir.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.txt_log.clear()
        self.progress.setMaximum(len(self.state.rows))
        self.progress.setValue(0)

        kwargs = dict(
            template_path=self.state.template_path,
            template_type=self.state.template_type,
            fields=self.state.fields,
            colmap=self.state.colmap,
            rows=self.state.rows,
            out_dir=self.state.out_dir,
            mode=mode,
            filename_col=self.state.filename_col,
            skip_rows=skip,
            exists_mode=exists_mode,
            first_word=self.state.filename_first_word,
            slide_sel=self.state.slide_sel,
        )
        self._worker = _Worker(copy.deepcopy(kwargs))
        self._worker.sig_progress.connect(self._on_progress)
        self._worker.sig_done.connect(self._on_done)
        self._worker.finished.connect(self._worker_finished)
        self._worker.start()

    def _cancel(self):
        if self._worker:
            self._worker.cancel()
            self.lbl_status.setText("Отмена…")

    def _on_progress(self, i, total, label):
        self.progress.setMaximum(max(1, total))
        self.progress.setValue(i)
        self.lbl_status.setText(f"{label} — {i} из {total}")

    def _on_done(self, rep: dict):
        self.btn_start.setEnabled(True)
        self.btn_back.setEnabled(True)
        self.btn_restart.setEnabled(True)
        self.btn_dir.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.btn_open_dir.setEnabled(bool(rep.get("created")))

        lines = []
        if rep.get("fatal"):
            lines.append(f"Критическая ошибка: {rep['fatal']}")
        lines.append(f"Обработано: {self.progress.value()} из {len(self.state.rows)}")
        lines.append(f"Сохранено записей: {len(rep['created'])}")
        lines.append(f"Папка результатов: {self.state.out_dir}")
        if rep["cancelled"]:
            lines.append("Генерация была отменена пользователем.")
        if rep["skipped"]:
            lines.append(f"Пропущено строк: {len(rep['skipped'])}")
            for r, reason in rep["skipped"][:50]:
                lines.append(f"  строка {r}: {reason}")
            if len(rep["skipped"]) > 50:
                lines.append(f"  …и ещё {len(rep['skipped']) - 50}")
        if rep["review"]:
            lines.append(f"Требуют ручной проверки: {len(rep['review'])}")
            for r, notes in rep["review"][:50]:
                for n in notes:
                    lines.append(f"  строка {r}: {n}")
        if rep.get("report_path"):
            lines.append(f"Отчёт сохранён: {rep['report_path']}")
        text = "\n".join(lines)
        self.txt_log.setPlainText(text)
        self.lbl_status.setText("Отменено" if rep["cancelled"] else
                                "Завершено с ошибками" if rep.get("fatal") or
                                (rep["skipped"] and not rep["created"]) else "Готово")
        QMessageBox.information(self, "Генерация завершена", text[:3000])

    def _worker_finished(self):
        worker = self._worker
        if worker:
            worker.wait()
            self._worker = None

    def stop_worker(self):
        if self._worker and self._worker.isRunning():
            self._worker.cancel()
            self._worker.wait()

    def _open_dir(self):
        if self.state.out_dir:
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.state.out_dir))
