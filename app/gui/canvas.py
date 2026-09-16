"""Preview canvas: displays a page/slide image, supports drawing PDF
rectangles with the mouse and clicking PPTX text blocks."""

from __future__ import annotations

from PySide6.QtCore import QPointF, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QPainter, QPen, QPixmap
from PySide6.QtWidgets import QWidget

MODE_NONE = "none"
MODE_DRAW = "draw"      # PDF: drag to draw a rect
MODE_PICK = "pick"      # PPTX: click a text block
MODE_EYEDROP = "eyedrop"  # click to pick the pixel colour under the cursor

_COLOR_SEL = QColor(30, 144, 255)      # placements of selected field
_COLOR_OTHER = QColor(120, 120, 120)   # other placements
_COLOR_DRAG = QColor(255, 140, 0)
_COLOR_HOVER = QColor(60, 180, 75)


class PreviewCanvas(QWidget):
    sigRectDrawn = Signal(float, float, float, float)  # doc coords (pt)
    sigShapePicked = Signal(int)                       # pptx shape_id
    sigPointClicked = Signal(float, float)             # doc coords (pt), eyedrop

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
        self.update()

    def set_mode(self, mode: str):
        self._mode = mode
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

        if self._drag_start is not None and self._drag_cur is not None:
            stroke(QRectF(self._drag_start, self._drag_cur).normalized(), _COLOR_DRAG, 60)
        p.end()

    # ---- mouse ---------------------------------------------------------
    def mousePressEvent(self, ev):
        if self._pix is None:
            return
        if self._mode == MODE_EYEDROP and ev.button() == Qt.LeftButton:
            x, y = self._to_doc(ev.position())
            self.sigPointClicked.emit(x, y)
        elif self._mode == MODE_DRAW and ev.button() == Qt.LeftButton:
            self._drag_start = ev.position()
            self._drag_cur = ev.position()
            self.update()
        elif self._mode == MODE_PICK and ev.button() == Qt.LeftButton:
            shp = self._shape_at(ev.position())
            if shp is not None:
                self.sigShapePicked.emit(shp["shape_id"])

    def mouseMoveEvent(self, ev):
        if self._pix is None:
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
        if self._mode == MODE_DRAW and ev.button() == Qt.LeftButton and self._drag_start:
            self._drag_cur = ev.position()
            x0, y0 = self._to_doc(self._drag_start)
            x1, y1 = self._to_doc(self._drag_cur)
            self._drag_start = self._drag_cur = None
            self.update()
            if abs(x1 - x0) > 5 and abs(y1 - y0) > 5:
                self.sigRectDrawn.emit(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))
