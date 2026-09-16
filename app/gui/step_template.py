"""Step 1: load template, mark fields and placements."""

from __future__ import annotations

import os

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QImage, QPixmap
from PySide6.QtWidgets import (
    QColorDialog, QComboBox, QDialog, QDialogButtonBox, QDoubleSpinBox,
    QFileDialog, QGroupBox, QHBoxLayout, QInputDialog, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMessageBox, QPushButton, QRadioButton,
    QSplitter, QVBoxLayout, QWidget,
)

from app.core import mapping_io, pdf_engine, pptx_engine
from app.core.fonts import list_system_fonts
from app.core.models import (
    ALIGN_CENTER, ALIGN_LEFT, ALIGN_RIGHT, Field, PdfPlacement,
    PptxPlacement, TEMPLATE_PDF, TEMPLATE_PPTX,
)
from app.gui.canvas import MODE_DRAW, MODE_PICK, PreviewCanvas
from app.gui import pptx_render
from app.gui.state import AppState

_ALIGN_LABELS = {"left": "Слева", "center": "По центру", "right": "Справа"}
_ALIGN_VALUES = {v: k for k, v in _ALIGN_LABELS.items()}


class BindDialog(QDialog):
    """Choose how a bound PPTX text block is substituted."""

    def __init__(self, pattern: str, detected_token: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Привязка текстового блока")
        self.pattern = pattern
        self.detected_token = detected_token
        self.mode = "whole"
        self.token = ""

        v = QVBoxLayout(self)
        v.addWidget(QLabel("Полный текст блока:"))
        self.txt_pattern = QLineEdit(pattern)
        self.txt_pattern.setReadOnly(True)
        v.addWidget(self.txt_pattern)

        self.rb_whole = QRadioButton("Заменить весь текст блока")
        self.rb_frag = QRadioButton("Заменить только фрагмент:")
        self.rb_token = QRadioButton(f"Заменить токен «{detected_token}»")
        v.addWidget(self.rb_whole)
        v.addWidget(self.rb_frag)

        row = QHBoxLayout()
        row.addSpacing(24)
        self.ed_frag = QLineEdit(detected_token)
        self.ed_frag.setPlaceholderText("подстрока из текста блока")
        btn_take = QPushButton("Взять выделенное")
        btn_take.clicked.connect(self._take_selection)
        row.addWidget(self.ed_frag, 1)
        row.addWidget(btn_take)
        v.addLayout(row)
        v.addWidget(self.rb_token)
        if not detected_token:
            self.rb_token.setVisible(False)

        if detected_token:
            self.rb_token.setChecked(True)
        else:
            self.rb_whole.setChecked(True)
        self.ed_frag.textChanged.connect(
            lambda _t: self.rb_frag.setChecked(True))

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

    def _take_selection(self):
        sel = self.txt_pattern.selectedText()
        if sel:
            self.ed_frag.setText(sel)
            self.rb_frag.setChecked(True)

    def _accept(self):
        if self.rb_frag.isChecked():
            frag = self.ed_frag.text()
            if not frag or frag not in self.pattern:
                QMessageBox.warning(
                    self, "Фрагмент",
                    "Фрагмент должен быть непустой подстрокой текста блока.")
                return
            self.mode, self.token = "fragment", frag
        elif self.rb_token.isChecked() and self.detected_token:
            self.mode, self.token = "token", self.detected_token
        else:
            self.mode, self.token = "whole", ""
        self.accept()


class StepTemplate(QWidget):
    def __init__(self, state: AppState, go_next, parent=None):
        super().__init__(parent)
        self.state = state
        self._go_next = go_next
        self._prs = None               # open pptx template
        self._page = 0                 # current page/slide index
        self._page_count = 0
        self._doc_w = 0.0
        self._doc_h = 0.0
        self._blocks: list[dict] = []  # pptx text blocks on current slide
        self._loading_style = False

        root = QVBoxLayout(self)
        top = QHBoxLayout()
        self.btn_open = QPushButton("Выбрать файл шаблона (.pptx или .pdf)…")
        self.btn_open.clicked.connect(self._open_template)
        self.lbl_template = QLabel("Шаблон не выбран")
        top.addWidget(self.btn_open)
        top.addWidget(self.lbl_template, 1)
        root.addLayout(top)

        split = QSplitter(Qt.Horizontal)

        # left: preview + page nav
        left = QWidget()
        lv = QVBoxLayout(left)
        nav = QHBoxLayout()
        self.btn_prev = QPushButton("◀")
        self.btn_next = QPushButton("▶")
        self.lbl_page = QLabel("—")
        self.btn_prev.clicked.connect(lambda: self._set_page(self._page - 1))
        self.btn_next.clicked.connect(lambda: self._set_page(self._page + 1))
        nav.addWidget(self.btn_prev)
        nav.addWidget(self.lbl_page)
        nav.addWidget(self.btn_next)
        nav.addStretch(1)
        lv.addLayout(nav)
        self.canvas = PreviewCanvas()
        self.canvas.sigRectDrawn.connect(self._on_rect_drawn)
        self.canvas.sigShapePicked.connect(self._on_shape_picked)
        lv.addWidget(self.canvas, 1)
        split.addWidget(left)

        # right: fields panel
        right = QWidget()
        rv = QVBoxLayout(right)

        rv.addWidget(QLabel("Поля (имена используются для привязки к столбцам Excel):"))
        self.lst_fields = QListWidget()
        self.lst_fields.currentRowChanged.connect(self._on_field_selected)
        rv.addWidget(self.lst_fields, 1)
        fbtns = QHBoxLayout()
        self.btn_add_field = QPushButton("Добавить поле…")
        self.btn_del_field = QPushButton("Удалить поле")
        self.btn_add_field.clicked.connect(self._add_field)
        self.btn_del_field.clicked.connect(self._del_field)
        fbtns.addWidget(self.btn_add_field)
        fbtns.addWidget(self.btn_del_field)
        rv.addLayout(fbtns)

        self.grp_bind = QGroupBox("Места выбранного поля в шаблоне")
        bv = QVBoxLayout(self.grp_bind)
        self.lbl_hint = QLabel("")
        self.lbl_hint.setWordWrap(True)
        bv.addWidget(self.lbl_hint)
        self.lst_blocks = QListWidget()   # pptx: detected text blocks
        self.lst_blocks.itemClicked.connect(self._on_block_clicked)
        self.lst_blocks.setVisible(False)
        bv.addWidget(self.lst_blocks)
        self.lst_places = QListWidget()
        self.lst_places.currentRowChanged.connect(self._on_place_selected)
        bv.addWidget(self.lst_places, 1)
        self.btn_del_place = QPushButton("Удалить выбранное место")
        self.btn_del_place.clicked.connect(self._del_place)
        bv.addWidget(self.btn_del_place)
        rv.addWidget(self.grp_bind, 2)

        # PDF style box
        self.grp_style = QGroupBox("Стиль текста (для PDF-рамок)")
        sv = QVBoxLayout(self.grp_style)
        sv.addWidget(QLabel("Шрифт:"))
        self.cmb_font = QComboBox()
        self.cmb_font.setEditable(False)
        sv.addWidget(self.cmb_font)
        self.btn_custom_font = QPushButton("Указать свой .ttf файл…")
        self.btn_custom_font.clicked.connect(self._pick_custom_font)
        sv.addWidget(self.btn_custom_font)
        row = QHBoxLayout()
        row.addWidget(QLabel("Размер:"))
        self.spin_size = QDoubleSpinBox()
        self.spin_size.setRange(4, 200)
        self.spin_size.setValue(24)
        row.addWidget(self.spin_size)
        row.addWidget(QLabel("Выравнивание:"))
        self.cmb_align = QComboBox()
        for v in (ALIGN_LEFT, ALIGN_CENTER, ALIGN_RIGHT):
            self.cmb_align.addItem(_ALIGN_LABELS[v], v)
        self.cmb_align.setCurrentIndex(1)
        row.addWidget(self.cmb_align)
        sv.addLayout(row)
        self.btn_color = QPushButton("Цвет текста…")
        self._color = "#000000"
        self.btn_color.clicked.connect(self._pick_color)
        sv.addWidget(self.btn_color)
        self.grp_style.setVisible(False)
        rv.addWidget(self.grp_style)
        self.cmb_font.currentIndexChanged.connect(self._style_changed)
        self.cmb_align.currentIndexChanged.connect(self._style_changed)
        self.spin_size.valueChanged.connect(self._style_changed)

        self.btn_save_map = QPushButton("Сохранить разметку (.mapping.json)")
        self.btn_save_map.clicked.connect(self._save_mapping)
        self.btn_save_map.setEnabled(False)
        rv.addWidget(self.btn_save_map)

        self.btn_next = QPushButton("Далее →")
        self.btn_next.clicked.connect(self._next)
        rv.addWidget(self.btn_next)

        split.addWidget(right)
        split.setStretchFactor(0, 3)
        split.setStretchFactor(1, 2)
        root.addWidget(split, 1)

    # ---- template loading ----------------------------------------------
    def _open_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите шаблон", "", "Шаблоны (*.pptx *.pdf)"
        )
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        if ext not in (".pptx", ".pdf"):
            QMessageBox.warning(self, "Шаблон", "Поддерживаются только .pptx и .pdf")
            return
        try:
            if ext == ".pptx":
                self._prs = pptx_engine.open_template(path)
                self._page_count = pptx_engine.slide_count(self._prs)
                self._doc_w = self._prs.slide_width / 12700.0
                self._doc_h = self._prs.slide_height / 12700.0
                self.state.template_type = TEMPLATE_PPTX
            else:
                self._prs = None
                self._page_count = pdf_engine.page_count(path)
                self._doc_w, self._doc_h = pdf_engine.page_size_pt(path, 0)
                self.state.template_type = TEMPLATE_PDF
        except Exception as exc:
            QMessageBox.critical(self, "Шаблон", f"Не удалось открыть файл:\n{exc}")
            return

        self.state.template_path = path
        self.state.fields = []
        self._page = 0
        self.lbl_template.setText(os.path.basename(path))
        self.btn_save_map.setEnabled(True)

        if mapping_io.has_saved_mapping(path):
            ans = QMessageBox.question(
                self, "Разметка",
                "Рядом с шаблоном найдена сохранённая разметка (.mapping.json).\n"
                "Загрузить её вместо разметки заново?",
            )
            if ans == QMessageBox.Yes:
                try:
                    ttype, fields, extra = mapping_io.load_mapping(path)
                    if ttype == self.state.template_type:
                        self.state.fields = fields
                        if "filename_source" in extra:
                            self.state.filename_source = tuple(extra["filename_source"])
                        if "filename_first_word" in extra:
                            self.state.filename_first_word = extra["filename_first_word"]
                    else:
                        QMessageBox.warning(self, "Разметка", "Тип шаблона в разметке не совпадает.")
                except Exception as exc:
                    QMessageBox.warning(self, "Разметка", f"Не удалось загрузить разметку:\n{exc}")

        self._configure_for_type()
        self._set_page(0)
        self._refresh_fields()

    def _configure_for_type(self):
        is_pdf = self.state.template_type == TEMPLATE_PDF
        self.canvas.set_mode(MODE_DRAW if is_pdf else MODE_PICK)
        self.grp_style.setVisible(is_pdf)
        self.lst_blocks.setVisible(not is_pdf)
        if is_pdf:
            self.lbl_hint.setText(
                "Выделите поле в списке выше и растяните рамку мышью поверх "
                "текста-плейсхолдера в превью. Стиль настраивается ниже.\n\n"
                "ВНИМАНИЕ: затирание подбирает цвет фона автоматически; на "
                "градиентном или текстурном фоне результат может отличаться. "
                "Проверьте пример перед массовой генерацией."
            )
            self._populate_fonts()
        else:
            self.lbl_hint.setText(
                "Выделите поле и кликните по текстовому блоку в превью "
                "(или в списке блоков ниже), чтобы привязать его к полю."
            )

    def _populate_fonts(self):
        self.cmb_font.blockSignals(True)
        self.cmb_font.clear()
        try:
            fonts = list_system_fonts(cyrillic_only=True)
        except Exception:
            fonts = []
        for name, path in fonts:
            self.cmb_font.addItem(name, path)
        # prefer Arial-like default
        for i in range(self.cmb_font.count()):
            if "arial" in self.cmb_font.itemText(i).lower():
                self.cmb_font.setCurrentIndex(i)
                break
        self.cmb_font.blockSignals(False)

    def _set_page(self, idx: int):
        if not self.state.template_path or self._page_count == 0:
            return
        self._page = max(0, min(idx, self._page_count - 1))
        self.lbl_page.setText(f"Страница/слайд {self._page + 1} из {self._page_count}")
        if self.state.template_type == TEMPLATE_PDF:
            png = pdf_engine.render_page_png(self.state.template_path, self._page, dpi=150)
            img = QImage.fromData(png)
            self.canvas.set_image(QPixmap.fromImage(img), self._doc_w, self._doc_h)
            self.canvas.set_pickable_shapes([])
        else:
            pix = pptx_render.render_slide(self._prs, self._page, dpi=150)
            self.canvas.set_image(pix, self._doc_w, self._doc_h)
            self._blocks = [b for b in pptx_engine.list_text_shapes(self._prs)
                            if b["slide"] == self._page]
            self.canvas.set_pickable_shapes(self._blocks)
            self._refresh_blocks()
        self._refresh_overlays()

    # ---- fields ----------------------------------------------------------
    def _cur_field(self) -> Field | None:
        row = self.lst_fields.currentRow()
        if 0 <= row < len(self.state.fields):
            return self.state.fields[row]
        return None

    def _add_field(self):
        name, ok = QInputDialog.getText(self, "Новое поле", "Имя поля (например, «ФИО»):")
        name = name.strip()
        if not ok or not name:
            return
        if any(f.name == name for f in self.state.fields):
            QMessageBox.warning(self, "Поле", "Поле с таким именем уже есть.")
            return
        self.state.fields.append(Field(name=name))
        self._refresh_fields()
        self.lst_fields.setCurrentRow(len(self.state.fields) - 1)

    def _del_field(self):
        row = self.lst_fields.currentRow()
        if 0 <= row < len(self.state.fields):
            del self.state.fields[row]
            self._refresh_fields()
            self._refresh_overlays()

    def _refresh_fields(self):
        self.lst_fields.blockSignals(True)
        self.lst_fields.clear()
        for f in self.state.fields:
            self.lst_fields.addItem(f"{f.name}  ({len(f.placements)} место/мест)")
        self.lst_fields.blockSignals(False)
        if self.state.fields:
            self.lst_fields.setCurrentRow(0)
        self._refresh_places()

    def _refresh_places(self):
        self.lst_places.blockSignals(True)
        self.lst_places.clear()
        f = self._cur_field()
        if f:
            for pl in f.placements:
                if isinstance(pl, PptxPlacement):
                    mode_lbl = {"whole": "весь текст", "fragment": "фрагмент",
                                "token": "токен"}.get(pl.mode, pl.mode)
                    extra = f' ({pl.token[:30]})' if pl.mode != "whole" and pl.token else ""
                    self.lst_places.addItem(
                        f"слайд {pl.slide + 1}, блок #{pl.shape_id} — {mode_lbl}{extra}")
                else:
                    self.lst_places.addItem(
                        f"стр. {pl.page + 1}, рамка "
                        f"({pl.rect[0]:.0f},{pl.rect[1]:.0f})–({pl.rect[2]:.0f},{pl.rect[3]:.0f})"
                    )
        self.lst_places.blockSignals(False)

    def _refresh_blocks(self):
        self.lst_blocks.clear()
        bound = {pl.shape_id for f in self.state.fields for pl in f.placements
                 if isinstance(pl, PptxPlacement)}
        for b in self._blocks:
            mark = " ✓" if b["shape_id"] in bound else ""
            text = b["text"].replace("\n", " ")[:60]
            item = QListWidgetItem(f"#{b['shape_id']}: {text}{mark}")
            item.setData(Qt.UserRole, b["shape_id"])
            self.lst_blocks.addItem(item)

    def _on_field_selected(self, _row):
        self._refresh_places()
        self._refresh_overlays()

    # ---- placement binding ------------------------------------------------
    def _on_shape_picked(self, shape_id: int):
        f = self._cur_field()
        if f is None:
            QMessageBox.information(self, "Поле", "Сначала добавьте и выберите поле в списке справа.")
            return
        if any(isinstance(p, PptxPlacement) and p.shape_id == shape_id and p.slide == self._page
               for p in f.placements):
            return
        block = next((b for b in self._blocks if b["shape_id"] == shape_id), None)
        pattern = block["text"] if block else ""
        detected = pptx_engine.detect_token(pattern) if pattern else ""
        dlg = BindDialog(pattern, detected, self)
        if dlg.exec() != QDialog.Accepted:
            return
        f.placements.append(PptxPlacement(
            slide=self._page, shape_id=shape_id,
            pattern=pattern, token=dlg.token, mode=dlg.mode,
        ))
        self._refresh_fields_keep_sel()
        self._refresh_blocks()
        self._refresh_overlays()

    def _on_block_clicked(self, item: QListWidgetItem):
        self._on_shape_picked(item.data(Qt.UserRole))

    def _on_rect_drawn(self, x0, y0, x1, y1):
        f = self._cur_field()
        if f is None:
            QMessageBox.information(self, "Поле", "Сначала добавьте и выберите поле в списке справа.")
            return
        pl = PdfPlacement(
            page=self._page,
            rect=(x0, y0, x1, y1),
            font_path=self.cmb_font.currentData() or "",
            font_size=self.spin_size.value(),
            color=self._color,
            align=self.cmb_align.currentData(),
        )
        f.placements.append(pl)
        self._refresh_fields_keep_sel()
        self._refresh_overlays()

    def _del_place(self):
        f = self._cur_field()
        row = self.lst_places.currentRow()
        if f and 0 <= row < len(f.placements):
            del f.placements[row]
            self._refresh_fields_keep_sel()
            self._refresh_overlays()
            self._refresh_blocks()

    def _refresh_fields_keep_sel(self):
        row = self.lst_fields.currentRow()
        self._refresh_fields()
        if 0 <= row < self.lst_fields.count():
            self.lst_fields.setCurrentRow(row)

    def _refresh_overlays(self):
        sel, other = [], []
        cur = self._cur_field()
        for f in self.state.fields:
            target = sel if f is cur else other
            for pl in f.placements:
                if isinstance(pl, PdfPlacement) and pl.page == self._page:
                    target.append(pl.rect)
                elif isinstance(pl, PptxPlacement) and pl.slide == self._page:
                    b = next((b for b in self._blocks if b["shape_id"] == pl.shape_id), None)
                    if b:
                        target.append((b["left_pt"], b["top_pt"],
                                       b["left_pt"] + b["width_pt"], b["top_pt"] + b["height_pt"]))
        self.canvas.set_overlays(sel, other)

    # ---- PDF style ---------------------------------------------------------
    def _on_place_selected(self, row):
        f = self._cur_field()
        if not f or not (0 <= row < len(f.placements)):
            return
        pl = f.placements[row]
        if isinstance(pl, PdfPlacement):
            self._loading_style = True
            idx = self.cmb_font.findData(pl.font_path)
            if idx >= 0:
                self.cmb_font.setCurrentIndex(idx)
            elif pl.font_path:
                self.cmb_font.addItem(os.path.basename(pl.font_path), pl.font_path)
                self.cmb_font.setCurrentIndex(self.cmb_font.count() - 1)
            self.spin_size.setValue(pl.font_size)
            self._color = pl.color
            self.cmb_align.setCurrentIndex(
                max(0, self.cmb_align.findData(pl.align)))
            self._loading_style = False

    def _style_changed(self, *_a):
        if self._loading_style:
            return
        f = self._cur_field()
        row = self.lst_places.currentRow()
        if f and 0 <= row < len(f.placements):
            pl = f.placements[row]
            if isinstance(pl, PdfPlacement):
                pl.font_path = self.cmb_font.currentData() or pl.font_path
                pl.font_size = self.spin_size.value()
                pl.color = self._color
                pl.align = self.cmb_align.currentData()
        self._refresh_overlays()

    def _pick_custom_font(self):
        path, _ = QFileDialog.getOpenFileName(self, "Шрифт", "", "Шрифты (*.ttf *.otf *.ttc)")
        if path:
            self.cmb_font.addItem(os.path.basename(path), path)
            self.cmb_font.setCurrentIndex(self.cmb_font.count() - 1)

    def _pick_color(self):
        c = QColorDialog.getColor(QColor(self._color), self, "Цвет текста")
        if c.isValid():
            self._color = c.name()
            self._style_changed()

    # ---- save / next --------------------------------------------------------
    def _save_mapping(self):
        if not self.state.template_path:
            return
        try:
            path = mapping_io.save_mapping(
                self.state.template_path, self.state.template_type, self.state.fields,
                filename_source=self.state.filename_source,
                filename_first_word=self.state.filename_first_word,
            )
            QMessageBox.information(self, "Разметка", f"Разметка сохранена:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "Разметка", f"Не удалось сохранить:\n{exc}")

    def _next(self):
        if not self.state.template_path:
            QMessageBox.warning(self, "Шаблон", "Сначала выберите файл шаблона.")
            return
        if not self.state.fields:
            QMessageBox.warning(self, "Поля", "Создайте хотя бы одно поле.")
            return
        empty = [f.name for f in self.state.fields if not f.placements]
        if empty:
            ans = QMessageBox.question(
                self, "Поля",
                "У полей нет привязанных мест: " + ", ".join(empty) +
                ".\nОни будут подставляться «в никуда». Продолжить?",
            )
            if ans != QMessageBox.Yes:
                return
        self._go_next()
