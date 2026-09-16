"""Approximate PPTX slide renderer for the preview pane.

PowerPoint rendering is not available offline, so we draw an approximate
preview: pictures, solid shape fills and text boxes at their real
positions. It is sufficient for clicking a text block to bind it.
"""

from __future__ import annotations

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.oxml.ns import qn
from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QImage, QPainter, QPixmap

EMU_PER_PT = 12700.0

_ALIGN_FLAGS = {1: Qt.AlignLeft, 2: Qt.AlignHCenter, 3: Qt.AlignRight}


def _emu2px(v, dpi):
    return v / EMU_PER_PT * dpi / 72.0


def _shape_rect(shp, dpi):
    try:
        return QRectF(
            _emu2px(shp.left, dpi), _emu2px(shp.top, dpi),
            _emu2px(shp.width, dpi), _emu2px(shp.height, dpi),
        )
    except Exception:
        return QRectF()


def _solid_color(fill):
    try:
        if fill.type == MSO_FILL_TYPE.SOLID:
            rgb = fill.fore_color.rgb
            if rgb is not None:
                return QColor("#" + str(rgb))
    except Exception:
        pass
    return None


def _group_transform(shp):
    """Return (off_x, off_y, ch_off_x, ch_off_y, scale_x, scale_y) in EMU."""
    try:
        xfrm = shp.element.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
        off = xfrm.find(qn("a:off"))
        ext = xfrm.find(qn("a:ext"))
        choff = xfrm.find(qn("a:chOff"))
        chext = xfrm.find(qn("a:chExt"))
        sx = int(ext.get("cx")) / max(1, int(chext.get("cx")))
        sy = int(ext.get("cy")) / max(1, int(chext.get("cy")))
        return (
            int(off.get("x")), int(off.get("y")),
            int(choff.get("x")), int(choff.get("y")), sx, sy,
        )
    except Exception:
        return None


def _draw_shapes(painter, shapes, dpi, dx=0.0, dy=0.0, sx=1.0, sy=1.0):
    for shp in shapes:
        try:
            if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
                tr = _group_transform(shp)
                if tr:
                    ox, oy, cox, coy, gsx, gsy = tr
                    painter.save()
                    painter.translate(
                        _emu2px(ox - cox * gsx, dpi), _emu2px(oy - coy * gsy, dpi)
                    )
                    painter.scale(gsx, gsy)
                    _draw_shapes(painter, shp.shapes, dpi)
                    painter.restore()
                    continue
            rect = _shape_rect(shp, dpi)
            if rect.width() <= 0 or rect.height() <= 0:
                continue
            if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    pm = QPixmap()
                    pm.loadFromData(shp.image.blob)
                    painter.drawPixmap(rect.toRect(), pm)
                except Exception:
                    pass
                continue
            fill = None
            try:
                fill = _solid_color(shp.fill)
            except Exception:
                pass
            pen_color = None
            try:
                if shp.line.fill.type == MSO_FILL_TYPE.SOLID:
                    rgb = shp.line.color.rgb
                    if rgb is not None:
                        pen_color = QColor("#" + str(rgb))
            except Exception:
                pass
            if fill is not None or pen_color is not None:
                painter.save()
                if fill is not None:
                    painter.setBrush(fill)
                else:
                    painter.setBrush(Qt.NoBrush)
                if pen_color is not None:
                    painter.setPen(pen_color)
                else:
                    painter.setPen(Qt.NoPen)
                painter.drawRect(rect)
                painter.restore()
            if shp.has_text_frame and shp.text_frame.text.strip():
                _draw_text(painter, shp, rect, dpi)
        except Exception:
            continue


def _draw_text(painter, shp, rect: QRectF, dpi):
    tf = shp.text_frame
    size_pt, family, bold, color = 18.0, "Arial", False, QColor("black")
    align = Qt.AlignHCenter
    for para in tf.paragraphs:
        try:
            if para.alignment is not None:
                align = _ALIGN_FLAGS.get(int(para.alignment), Qt.AlignHCenter)
        except Exception:
            pass
        run = next((r for r in para.runs if r.text), None)
        if run is not None:
            try:
                if run.font.size is not None:
                    size_pt = run.font.size.pt
            except Exception:
                pass
            try:
                if run.font.name:
                    family = run.font.name
            except Exception:
                pass
            try:
                bold = bool(run.font.bold)
            except Exception:
                pass
            try:
                rgb = run.font.color.rgb
                if rgb is not None:
                    color = QColor("#" + str(rgb))
            except Exception:
                pass
            break
    font = QFont(family)
    font.setPixelSize(max(1, int(size_pt * dpi / 72.0)))
    font.setBold(bold)
    painter.save()
    painter.setFont(font)
    painter.setPen(color)
    painter.drawText(rect, align | Qt.AlignVCenter | Qt.TextWordWrap, tf.text)
    painter.restore()


def render_slide(prs, slide_index: int, dpi: int = 150) -> QPixmap:
    slide = prs.slides[slide_index]
    w = _emu2px(prs.slide_width, dpi)
    h = _emu2px(prs.slide_height, dpi)
    img = QImage(int(w), int(h), QImage.Format_ARGB32)
    img.fill(QColor("white"))
    painter = QPainter(img)
    painter.setRenderHint(QPainter.Antialiasing, True)
    painter.setRenderHint(QPainter.TextAntialiasing, True)
    _draw_shapes(painter, slide.shapes, dpi)
    painter.end()
    return QPixmap.fromImage(img)
