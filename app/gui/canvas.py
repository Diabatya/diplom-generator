"""Preview canvas: displays a page/slide image, supports drawing PDF
rectangles with the mouse, clicking PPTX text blocks and editing the
facsimile-signature area (draw / move / resize)."""

from __future__ import annotations

from PySide6.QtCore import QPointF, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QPainter, QPen, QPixmap
from PySide6.QtWidgets import QWidget

MODE_NONE = "none"
MODE_DRAW = "draw"      # PDF: drag to draw a rect
MODE_PICK = "pick"      # PPTX: click a text block
MODE_EYEDROP = "eyedrop"  # click to pick the pixel colour under the cursor
MODE_SIGNATURE = "signature"  # draw/move/resize the signature area

_COLOR_SEL = QColor(30, 144, 255)      # placements of selected field
_COLOR_OTHER = QColor(120, 120, 120)   # other placements
_COLOR_DRAG = QColor(255, 140, 0)
_COLOR_HOVER = QColor(60, 180, 75)
_COLOR_SIG = QColor(0, 82, 155)        # signature frame (FNS dark blue)

_HANDLE = 12  # view px, side of the resize handle square


class PreviewCanvas(QWidget):
    sigRectDrawn = Signal(float, float, float, float)  # doc coords (pt)
    sigShapePicked = Signal(object)                    # (shape_id, cell)
    sigPointClicked = Signal(float, float)             # doc coords (pt), eyedrop
    sigSignatureChanged = Signal(float, float, float, float)  # doc pt

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(400, 300)
        self.setMouseTracking(True)
        self._pix: QPixmap | None = None
        self._doc_w = 1.0   # document size in points
        self._doc_h = 1.0
        self._mode = MODE_NONE
        self._shapes: list[dict] = []          # pickable pptx blocks (pt bboxes)
        self._placements_sel: list[QRectF] = []   # pt rects
        self._placements_other: list[QRectF] = []
        self._hover_id: int | None = None
        self._drag_start: QPointF | None = None
        self._drag_cur: QPointF | None = None
        # facsimile signature overlay (doc-pt rect + image to preview)
        self._sig_rect: tuple | None = None
        self._sig_pix: QPixmap | None = None
        self._sig_action: str | None = None    # "draw" | "move" | "resize"
        self._sig_grab: QPointF | None = None  # cursor pos minus rect origin (view px)

    # ---- setup ---------------------------------------------------------
    def set_image(self, pix: QPixmap, doc_w_pt: float, doc_h_pt: float):
        self._pix = pix
        self._doc_w = doc_w_pt or 1.0
        self._doc_h = doc_h_pt or 1.0
        self._drag_start = self._drag_cur = None
        self.update()

    def clear(self):
        self._pix = None
        self._shapes = []
        self._placements_sel = []
        self._placements_other = []
        self._sig_rect = None
        self._sig_pix = None
        self._sig_action = None
        self.update()

    def set_mode(self, mode: str):
        self._mode = mode
        self._sig_action = None
        self.setCursor(Qt.CrossCursor if mode == MODE_EYEDROP
                       else Qt.ArrowCursor)
        self.update()

    def set_pickable_shapes(self, shapes: list[dict]):
        self._shapes = shapes
        self.update()

    def set_overlays(self, selected: list, other: list):
        """selected/other: lists of (x0,y0,x1,y1) rects in doc points."""
        self._placements_sel = [QRectF(QPointF(r[0], r[1]), QPointF(r[2], r[3])) for r in selected]
        self._placements_other = [QRectF(QPointF(r[0], r[1]), QPointF(r[2], r[3])) for r in other]
        self.update()

    # ---- facsimile signature -------------------------------------------
    def set_signature(self, rect_pt: tuple | None, pix: QPixmap | None):
        """Show the signature image aspect-fit inside ``rect_pt``
        (x0,y0,x1,y1 in doc points). ``None`` rect removes the overlay."""
        self._sig_rect = tuple(rect_pt) if rect_pt else None
        self._sig_pix = pix
        self.update()

    def _sig_view_rect(self) -> QRectF | None:
        if self._sig_rect is None:
            return None
        return self._to_view(self._sig_rect)

    def _sig_handle_rect(self) -> QRectF | None:
        rv = self._sig_view_rect()
        if rv is None:
            return None
        return QRectF(rv.right() - _HANDLE, rv.bottom() - _HANDLE,
                      _HANDLE, _HANDLE)

    def _sig_hit(self, pos: QPointF) -> str:
        """"resize" | "move" | "draw" for a press at ``pos`` (view px)."""
        hr = self._sig_handle_rect()
        if hr is not None and hr.contains(pos):
            return "resize"
        rv = self._sig_view_rect()
        if rv is not None and rv.contains(pos):
            return "move"
        return "draw"

    def _sig_apply_drag(self, pos: QPointF):
        """Update ``_sig_rect`` for the active drag gesture."""
        if self._sig_action == "draw":
            r = QRectF(self._drag_start, pos).normalized()
            x0, y0 = self._to_doc(r.topLeft())
            x1, y1 = self._to_doc(r.bottomRight())
            self._sig_rect = (x0, y0, x1, y1)
            return
        if self._sig_rect is None:
            return
        if self._sig_action == "move" and self._sig_grab is not None:
            top_left = pos - self._sig_grab
            x0, y0 = self._to_doc(top_left)
            w = self._sig_rect[2] - self._sig_rect[0]
            h = self._sig_rect[3] - self._sig_rect[1]
            self._sig_rect = (x0, y0, x0 + w, y0 + h)
        elif self._sig_action == "resize":
            x1, y1 = self._to_doc(pos)
            x0, y0 = self._sig_rect[0], self._sig_rect[1]
            self._sig_rect = (x0, y0,
                              max(x1, x0 + 10.0), max(y1, y0 + 10.0))

    # ---- geometry helpers ----------------------------------------------
    def _scale(self) -> float:
        """displayed px per document point."""
        if self._pix is None:
            return 1.0
        sx = self.width() / self._doc_w
        sy = self.height() / self._doc_h
        return min(sx, sy)

    def _img_rect(self) -> QRectF:
        s = self._scale()
        w, h = self._doc_w * s, self._doc_h * s
        return QRectF((self.width() - w) / 2, (self.height() - h) / 2, w, h)

    def _to_doc(self, pos: QPointF) -> tuple[float, float]:
        r = self._img_rect()
        s = self._scale()
        return (pos.x() - r.x()) / s, (pos.y() - r.y()) / s

    def _to_view(self, rect_pt) -> QRectF:
        r = self._img_rect()
        s = self._scale()
        if isinstance(rect_pt, QRectF):
            x0, y0, x1, y1 = rect_pt.x(), rect_pt.y(), rect_pt.right(), rect_pt.bottom()
        else:
            x0, y0, x1, y1 = rect_pt
        return QRectF(r.x() + x0 * s, r.y() + y0 * s, (x1 - x0) * s, (y1 - y0) * s)

    def _shape_at(self, pos: QPointF) -> dict | None:
        x, y = self._to_doc(pos)
        for shp in reversed(self._shapes):  # topmost first
            if (shp["left_pt"] <= x <= shp["left_pt"] + shp["width_pt"]
                    and shp["top_pt"] <= y <= shp["top_pt"] + shp["height_pt"]):
                return shp
        return None

    # ---- painting ------------------------------------------------------
    def paintEvent(self, _ev):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(70, 70, 70))
        if self._pix is None:
            p.setPen(QColor(220, 220, 220))
            p.drawText(self.rect(), Qt.AlignCenter, "Шаблон не загружен")
            p.end()
            return
        r = self._img_rect()
        p.drawPixmap(r.toRect(), self._pix)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setRenderHint(QPainter.SmoothPixmapTransform, True)

        def stroke(rect_view: QRectF, color: QColor, fill_alpha: int):
            p.setPen(QPen(color, 2))
            c = QColor(color)
            c.setAlpha(fill_alpha)
            p.setBrush(c)
            p.drawRect(rect_view)

        for rect in self._placements_other:
            stroke(self._to_view(rect), _COLOR_OTHER, 40)
        for rect in self._placements_sel:
            stroke(self._to_view(rect), _COLOR_SEL, 60)

        if self._mode == MODE_PICK:
            for shp in self._shapes:
                rv = self._to_view((shp["left_pt"], shp["top_pt"],
                                    shp["left_pt"] + shp["width_pt"],
                                    shp["top_pt"] + shp["height_pt"]))
                col = _COLOR_HOVER if shp["shape_id"] == self._hover_id else QColor(200, 60, 60)
                p.setPen(QPen(col, 1, Qt.DashLine))
                p.setBrush(Qt.NoBrush)
                p.drawRect(rv)

        rv = self._sig_view_rect()
        if rv is not None:
            if self._sig_pix is not None and not self._sig_pix.isNull():
                target = rv
                pa = self._sig_pix.width() / max(1, self._sig_pix.height())
                if target.width() / max(1e-6, target.height()) > pa:
                    w = target.height() * pa
                    target = QRectF(target.x() + (target.width() - w) / 2,
                                    target.y(), w, target.height())
                else:
                    h = target.width() / pa
                    target = QRectF(target.x(), target.y() + (target.height() - h) / 2,
                                    target.width(), h)
                p.drawPixmap(target.toRect(), self._sig_pix)
            p.setPen(QPen(_COLOR_SIG, 2, Qt.DashLine))
            p.setBrush(Qt.NoBrush)
            p.drawRect(rv)
            hr = self._sig_handle_rect()
            if self._mode == MODE_SIGNATURE and hr is not None:
                p.setPen(QPen(QColor("white"), 1))
                p.setBrush(_COLOR_SIG)
                p.drawRect(hr)

        if self._drag_start is not None and self._drag_cur is not None \
                and self._mode in (MODE_DRAW, MODE_SIGNATURE) \
                and self._sig_action != "resize":
            stroke(QRectF(self._drag_start, self._drag_cur).normalized(),
                   _COLOR_DRAG if self._mode == MODE_DRAW else _COLOR_SIG, 60)
        p.end()

    # ---- mouse ---------------------------------------------------------
    def mousePressEvent(self, ev):
        if self._pix is None:
            return
        if self._mode == MODE_EYEDROP and ev.button() == Qt.LeftButton:
            x, y = self._to_doc(ev.position())
            self.sigPointClicked.emit(x, y)
        elif self._mode == MODE_SIGNATURE and ev.button() == Qt.LeftButton:
            self._sig_action = self._sig_hit(ev.position())
            if self._sig_action == "move":
                rv = self._sig_view_rect()
                self._sig_grab = ev.position() - rv.topLeft()
            elif self._sig_action == "draw":
                self._drag_start = ev.position()
                self._drag_cur = ev.position()
                self._sig_rect = None
            self.update()
        elif self._mode == MODE_DRAW and ev.button() == Qt.LeftButton:
            self._drag_start = ev.position()
            self._drag_cur = ev.position()
            self.update()
        elif self._mode == MODE_PICK and ev.button() == Qt.LeftButton:
            shp = self._shape_at(ev.position())
            if shp is not None:
                self.sigShapePicked.emit((shp["shape_id"], shp.get("cell")))

    def mouseMoveEvent(self, ev):
        if self._pix is None:
            return
        if self._mode == MODE_SIGNATURE:
            if self._sig_action is not None:
                if self._sig_action == "draw":
                    self._drag_cur = ev.position()
                self._sig_apply_drag(ev.position())
                self.update()
            else:
                hit = self._sig_hit(ev.position())
                self.setCursor(
                    Qt.SizeFDiagCursor if hit == "resize"
                    else Qt.SizeAllCursor if hit == "move"
                    else Qt.CrossCursor)
            return
        if self._mode == MODE_DRAW and self._drag_start is not None:
            self._drag_cur = ev.position()
            self.update()
        elif self._mode == MODE_PICK:
            shp = self._shape_at(ev.position())
            hid = shp["shape_id"] if shp else None
            if hid != self._hover_id:
                self._hover_id = hid
                self.setCursor(Qt.PointingHandCursor if hid else Qt.ArrowCursor)
                self.update()

    def mouseReleaseEvent(self, ev):
        if self._mode == MODE_SIGNATURE and ev.button() == Qt.LeftButton \
                and self._sig_action is not None:
            if self._sig_action == "draw":
                self._drag_cur = ev.position()
            self._sig_apply_drag(ev.position())
            self._drag_start = self._drag_cur = None
            self._sig_action = None
            self._sig_grab = None
            self.update()
            if self._sig_rect is not None:
                x0, y0, x1, y1 = self._sig_rect
                if x1 - x0 > 5 and y1 - y0 > 5:
                    self.sigSignatureChanged.emit(x0, y0, x1, y1)
            return
        if self._mode == MODE_DRAW and ev.button() == Qt.LeftButton and self._drag_start:
            self._drag_cur = ev.position()
            x0, y0 = self._to_doc(self._drag_start)
            x1, y1 = self._to_doc(self._drag_cur)
            self._drag_start = self._drag_cur = None
            self.update()
            if abs(x1 - x0) > 5 and abs(y1 - y0) > 5:
                self.sigRectDrawn.emit(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))
