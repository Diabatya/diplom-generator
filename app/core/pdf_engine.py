"""PDF template handling with PyMuPDF.

The original placeholder text is removed via redaction annotations and the
new value is drawn on top inside the user-defined rectangle with the
chosen TTF font, size, color and alignment.
"""

from __future__ import annotations

from statistics import median

try:
    import pymupdf as fitz
except ImportError:  # older PyMuPDF
    import fitz  # type: ignore  # PyMuPDF

from . import fonts as fontlib
from .models import Field, PdfPlacement



def open_doc(path: str) -> fitz.Document:
    return fitz.open(path)


def page_count(path: str) -> int:
    with fitz.open(path) as doc:
        return doc.page_count


def page_size_pt(path: str, page: int = 0) -> tuple[float, float]:
    with fitz.open(path) as doc:
        r = doc[page].rect
        return r.width, r.height


def render_page_png(path: str, page: int, dpi: int = 150) -> bytes:
    with fitz.open(path) as doc:
        pix = doc[page].get_pixmap(dpi=dpi)
        return pix.tobytes("png")


def _hex_to_rgb(color: str) -> tuple[float, float, float]:
    c = color.lstrip("#")
    if len(c) != 6:
        return (0.0, 0.0, 0.0)
    return tuple(int(c[i : i + 2], 16) / 255.0 for i in (0, 2, 4))


def _median_rgb(pix, pts):
    n, w = pix.n, pix.width
    rs, gs, bs = [], [], []
    for x, y in pts:
        off = (y * w + x) * n
        rs.append(pix.samples[off])
        gs.append(pix.samples[off + 1])
        bs.append(pix.samples[off + 2])
    return (median(rs) / 255, median(gs) / 255, median(bs) / 255)


def _sample_fill(page: fitz.Page, rect: fitz.Rect):
    """Median colour of the rect border pixels — used to cover old text."""
    try:
        pix = page.get_pixmap(clip=rect, dpi=72)
        w, h = pix.width, pix.height
        if w < 2 or h < 2:
            return (1, 1, 1)
        pts = [
            (0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1),
            (w // 2, 0), (w // 2, h - 1), (0, h // 2), (w - 1, h // 2),
        ]
        return _median_rgb(pix, pts)
    except Exception:
        return (1, 1, 1)


def _sample_center(page: fitz.Page, rect: fitz.Rect):
    """Median colour of the rect interior — the plate under the text."""
    inner = fitz.Rect(rect)
    inner.x0 += rect.width * 0.3
    inner.x1 -= rect.width * 0.3
    inner.y0 += rect.height * 0.3
    inner.y1 -= rect.height * 0.3
    if inner.is_empty:
        return _sample_fill(page, rect)
    try:
        pix = page.get_pixmap(clip=inner, dpi=72)
        w, h = pix.width, pix.height
        if w < 2 or h < 2:
            return _sample_fill(page, rect)
        pts = [(w // 4, h // 4), (w // 2, h // 2), (3 * w // 4, 3 * h // 4),
               (w // 2, h // 4), (w // 2, 3 * h // 4)]
        return _median_rgb(pix, pts)
    except Exception:
        return _sample_fill(page, rect)


def _lum(c) -> float:
    return 0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2]


def fill_risk_on_page(page: fitz.Page, rect: fitz.Rect) -> bool:
    """True when the auto fill colour (edge median) is much lighter than the
    area under the text — typical sign that the rect spills past a coloured
    plate onto the white page, which would punch a white hole into it."""
    edge = _sample_fill(page, rect)
    center = _sample_center(page, rect)
    return _lum(edge) - _lum(center) > 0.25


def fill_risk_at(path: str, page_no: int, rect) -> bool:
    """GUI helper: check the auto-fill risk for a placement rect."""
    try:
        with fitz.open(path) as doc:
            return fill_risk_on_page(doc[page_no], fitz.Rect(rect))
    except Exception:
        return False


_FILL_RISK_NOTE = ("края рамки светлее фона под текстом — возможно, рамка "
                   "захватывает фон за пределами цветной плашки; задайте "
                   "цвет заливки вручную")


def _apply_placements(doc: fitz.Document, fields: list[Field], values: dict[str, str]):
    """Edit ``doc`` in place. Returns list of review notes."""
    notes: list[str] = []
    per_page: dict[int, list[tuple[Field, PdfPlacement, str]]] = {}
    for f in fields:
        value = values.get(f.name, "")
        for pl in f.placements:
            if isinstance(pl, PdfPlacement):
                per_page.setdefault(pl.page, []).append((f, pl, value))

    font_names: dict[str, str] = {}

    def fname(path: str) -> str:
        if path not in font_names:
            font_names[path] = f"UF{len(font_names)}"
        return font_names[path]

    for page_no, items in per_page.items():
        if page_no >= doc.page_count:
            notes.append(f"страница {page_no + 1} отсутствует")
            continue
        page = doc[page_no]
        for _f, pl, _v in items:
            rect = fitz.Rect(pl.rect) & page.rect
            if pl.fill_color:
                fill = _hex_to_rgb(pl.fill_color)
            else:
                fill = _sample_fill(page, rect)
                if fill_risk_on_page(page, rect):
                    notes.append(f"{_f.name}: {_FILL_RISK_NOTE}")
            page.add_redact_annot(rect, fill=fill)
        page.apply_redactions(
            images=fitz.PDF_REDACT_IMAGE_NONE,
            graphics=fitz.PDF_REDACT_LINE_ART_NONE,
            text=fitz.PDF_REDACT_TEXT_REMOVE,
        )
        for f, pl, value in items:
            rect = fitz.Rect(pl.rect) & page.rect
            font_path = pl.font_path or fontlib.find_font_file("Arial") \
                or fontlib.find_font_file("DejaVu Sans") or ""
            size, fits = fontlib.fit_font_size(
                value, font_path, rect.width - 4, pl.font_size, allow_wrap=False
            )
            # vertical centering via baseline: base = y0 + (h - lh)/2 + asc*size
            lh = fontlib.line_height(font_path, size)
            asc = fontlib.ascender(font_path)
            baseline_y = rect.y0 + max(0.0, (rect.height - lh) / 2.0) + asc * size
            w = fontlib.measure_text_width(value, font_path, size)
            if pl.align == "left":
                x = rect.x0 + 2
            elif pl.align == "right":
                x = rect.x1 - w - 2
            else:
                x = rect.x0 + (rect.width - w) / 2.0
            kwargs = dict(fontsize=size, color=_hex_to_rgb(pl.color))
            if font_path:
                kwargs["fontname"] = fname(font_path)
                kwargs["fontfile"] = font_path
            page.insert_text(fitz.Point(x, baseline_y), value, **kwargs)
            if not fits:
                notes.append(
                    f"{f.name}: текст не поместился даже при минимальном размере ({size:.1f} pt)"
                )
            elif size < pl.font_size:
                notes.append(f"{f.name}: шрифт уменьшен до {size:.1f} pt")
    return notes


def render_pdf_doc(template_path: str, fields: list[Field], values: dict[str, str]):
    """Return an in-memory edited document plus review notes."""
    doc = fitz.open(template_path)
    notes = _apply_placements(doc, fields, values)
    return doc, notes


def render_single(template_path: str, fields: list[Field], values: dict[str, str], out_path: str):
    doc, notes = render_pdf_doc(template_path, fields, values)
    try:
        doc.save(out_path, garbage=4, deflate=True)
    finally:
        doc.close()
    return notes
