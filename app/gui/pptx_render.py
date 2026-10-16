"""Approximate PPTX slide renderer for the preview pane.

PowerPoint rendering is not available offline, so we draw an approximate
preview: the slide background (slide → layout → master), non-placeholder
master/layout graphics and the slide's own shapes — pictures, solid and
gradient fills, text boxes at their real positions. It is sufficient for
clicking a text block to bind it and for placing a facsimile signature
on the real layout.
"""

from __future__ import annotations

import colorsys

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import (QColor, QFont, QImage, QLinearGradient, QPainter,
                           QPixmap)

EMU_PER_PT = 12700.0

_ALIGN_FLAGS = {1: Qt.AlignLeft, 2: Qt.AlignHCenter, 3: Qt.AlignRight}

# fallback palette when the theme part cannot be read (Office defaults)
_OFFICE_THEME = {
    "dk1": "000000", "lt1": "FFFFFF", "dk2": "44546A", "lt2": "E7E6E6",
    "accent1": "4472C4", "accent2": "ED7D31", "accent3": "A5A5A5",
    "accent4": "FFC000", "accent5": "5B9BD5", "accent6": "70AD47",
    "hlink": "0563C1", "folHlink": "954F72",
}
_SCHEME_SYS = {"windowText": "000000", "window": "FFFFFF"}


def _emu2px(v, dpi):
    return v / EMU_PER_PT * dpi / 72.0


def _theme_colors(master) -> dict:
    """scheme-name → "rrggbb" from the master's theme part."""
    colors = dict(_OFFICE_THEME)
    try:
        theme_part = master.part.part_related_by(RT.THEME)
        from lxml import etree
        root = etree.fromstring(theme_part.blob)
        scheme = root.find(
            ".//" + qn("a:themeElements") + "/" + qn("a:clrScheme"))
        for slot in scheme:
            name = etree.QName(slot).localname
            clr = next(iter(slot), None)
            if clr is None:
                continue
            hexv = _clr_to_hex(clr, colors, transforms=False)
            if hexv:
                colors[name] = hexv
    except Exception:
        pass
    return colors


def _clr_to_hex(clr_el, theme: dict, transforms: bool = True):
    """a:srgbClr / a:schemeClr / a:sysClr element → "rrggbb" or None."""
    from lxml import etree
    tag = etree.QName(clr_el).localname
    if tag == "srgbClr":
        base = clr_el.get("val")
    elif tag == "sysClr":
        base = clr_el.get("lastClr") or _SCHEME_SYS.get(clr_el.get("val"))
    elif tag == "schemeClr":
        base = theme.get(clr_el.get("val"))
    else:
        return None
    if not base or len(base) != 6:
        return None
    r, g, b = (int(base[i:i + 2], 16) for i in (0, 2, 4))
    if transforms:
        lum_mod, lum_off, tint, shade = 1.0, 0.0, None, None
        for child in clr_el:
            ctag = etree.QName(child).localname
            v = int(child.get("val", 0)) / 100000.0
            if ctag == "lumMod":
                lum_mod = v
            elif ctag == "lumOff":
                lum_off = v
            elif ctag == "tint":
                tint = v
            elif ctag == "shade":
                shade = v
        if lum_mod != 1.0 or lum_off != 0.0:
            h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
            l = max(0.0, min(1.0, l * lum_mod + lum_off))
            r, g, b = (int(round(c * 255))
                       for c in colorsys.hls_to_rgb(h, l, s))
        if tint is not None:
            r, g, b = (int(round(c * tint + 255 * (1 - tint)))
                       for c in (r, g, b))
        if shade is not None:
            r, g, b = (int(round(c * shade)) for c in (r, g, b))
    return f"{r:02X}{g:02X}{b:02X}"


def _first_clr(parent):
    return next(iter(parent), None)


def _fill_paint(fill_el, part, theme: dict, rect: QRectF):
    """a:solidFill/gradFill/blipFill element → QColor | QLinearGradient
    | QPixmap | "phClr" | None."""
    from lxml import etree
    tag = etree.QName(fill_el).localname
    if tag == "solidFill":
        clr = _first_clr(fill_el)
        if clr is not None and etree.QName(clr).localname == "phClr":
            return "phClr"
        if clr is not None:
            hexv = _clr_to_hex(clr, theme)
            if hexv:
                return QColor("#" + hexv)
        return None
    gf = fill_el if tag == "gradFill" else None
    if gf is not None:
        stops = []
        gs_lst = gf.find(qn("a:gsLst"))
        if gs_lst is not None:
            for gs in gs_lst.findall(qn("a:gs")):
                clr = _first_clr(gs)
                hexv = _clr_to_hex(clr, theme) if clr is not None else None
                if hexv:
                    stops.append((int(gs.get("pos", 0)) / 100000.0,
                                  QColor("#" + hexv)))
        if len(stops) >= 2:
            lin = gf.find(qn("a:lin"))
            import math
            ang = math.radians(int(lin.get("ang", 0)) / 60000.0) \
                if lin is not None else 0.0
            dx, dy = math.sin(ang), -math.cos(ang)
            grad = QLinearGradient(
                rect.center().x() - dx * rect.width() / 2,
                rect.center().y() - dy * rect.height() / 2,
                rect.center().x() + dx * rect.width() / 2,
                rect.center().y() + dy * rect.height() / 2)
            for pos, col in stops:
                grad.setColorAt(max(0.0, min(1.0, pos)), col)
            return grad
        if stops:
            return stops[0][1]
        return None
    if tag != "blipFill":
        return None
    bf = fill_el
    if bf is not None:
        try:
            blip = bf.find(qn("a:blip"))
            rid = blip.get(qn("r:embed"))
            blob = part.rels[rid].target_part.blob
            pm = QPixmap()
            pm.loadFromData(blob)
            return pm
        except Exception:
            return None
    return None


def _bg_paint(bg_el, part, theme: dict, rect: QRectF):
    """Resolve p:bg (bgPr or bgRef) to a paintable value."""
    if bg_el is None:
        return None
    bgpr = bg_el.find(qn("p:bgPr"))
    if bgpr is not None:
        fill_el = next(iter(bgpr), None)
        if fill_el is not None:
            return _fill_paint(fill_el, part, theme, rect)
        return None
    bgref = bg_el.find(qn("p:bgRef"))
    if bgref is not None:
        idx = int(bgref.get("idx", 0))
        ref_clr = _first_clr(bgref)
        try:
            theme_part = part.part_related_by(RT.THEME)
            from lxml import etree
            root = etree.fromstring(theme_part.blob)
            fmt = root.find(".//" + qn("a:fmtScheme"))
            lst = fmt.find(qn("a:fillStyleLst") if idx >= 1001
                           else qn("a:bgFillStyleLst"))
            fills = list(lst)
            fill_el = fills[idx - 1001] if idx >= 1001 else fills[idx - 1]
        except Exception:
            fill_el = None
        if fill_el is not None:
            val = _fill_paint(fill_el, part, theme, rect)
            if val == "phClr" or (val is None and ref_clr is not None):
                hexv = (_clr_to_hex(ref_clr, theme)
                        if ref_clr is not None else None)
                return QColor("#" + hexv) if hexv else None
            return val
        if ref_clr is not None:
            hexv = _clr_to_hex(ref_clr, theme)
            if hexv:
                return QColor("#" + hexv)
    return None


def _slide_bg(slide, theme: dict, rect: QRectF):
    """Slide → layout → master background fill (first one defined)."""
    for obj in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        try:
            bg = obj.element.find(qn("p:cSld")).find(qn("p:bg"))
        except Exception:
            bg = None
        if bg is None:
            continue
        part = obj.part
        val = _bg_paint(bg, part, theme, rect)
        if val is not None:
            return val
    return None


def _decor_shapes(shapes):
    """Shapes that appear on the slide from a layout/master: everything
    except placeholders (placeholders themselves are not rendered)."""
    return [s for s in shapes
            if not getattr(s, "is_placeholder", False)]


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
    full = QRectF(0, 0, w, h)
    try:
        theme = _theme_colors(slide.slide_layout.slide_master)
    except Exception:
        theme = dict(_OFFICE_THEME)
    bg = _slide_bg(slide, theme, full)
    if isinstance(bg, QPixmap):
        painter.drawPixmap(full.toRect(), bg)
    elif bg is not None and bg != "phClr":
        painter.fillRect(full, bg)
    if slide.element.get("showMasterSp", "1") != "0":
        for anc in (slide.slide_layout.slide_master, slide.slide_layout):
            _draw_shapes(painter, _decor_shapes(anc.shapes), dpi)
    _draw_shapes(painter, slide.shapes, dpi)
    painter.end()
    return QPixmap.fromImage(img)
