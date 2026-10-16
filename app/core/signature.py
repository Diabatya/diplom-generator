"""Facsimile signature: shared validation and geometry helpers.

The signature is described by :class:`app.core.models.SignatureSpec`:
an image file plus a rectangle in document points. Engines stamp the
image aspect-fit inside that rect on every output slide/page.
"""

from __future__ import annotations

import os

from .models import SignatureSpec

MIN_RECT_PT = 12.0          # smallest meaningful selection (~4 mm)
_SUPPORTED_EXT = (".png", ".jpg", ".jpeg")


def image_error(image_path: str) -> str | None:
    """None when ``image_path`` is a readable PNG/JPEG, else a message."""
    if not image_path:
        return "не выбран файл изображения подписи"
    if not os.path.isfile(image_path):
        return f"файл изображения не найден: {image_path}"
    if os.path.splitext(image_path)[1].lower() not in _SUPPORTED_EXT:
        return "поддерживаются только PNG и JPEG"
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz  # type: ignore
        pix = fitz.Pixmap(image_path)
        ok = pix.width > 0 and pix.height > 0
        pix = None
        if not ok:
            return "файл не является читаемым изображением"
    except Exception as exc:
        return f"не удалось прочитать изображение: {exc}"
    return None


def image_size_px(image_path: str) -> tuple[int, int]:
    """(width, height) in pixels; raises on unreadable file."""
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz  # type: ignore
    pix = fitz.Pixmap(image_path)
    return pix.width, pix.height


def rect_error(rect, doc_w_pt: float, doc_h_pt: float) -> str | None:
    """None when ``rect`` intersects the document page with enough area."""
    if rect is None:
        return "область подписи не выделена — растяните рамку на превью"
    x0, y0, x1, y1 = rect
    if x1 - x0 < MIN_RECT_PT or y1 - y0 < MIN_RECT_PT:
        return "область подписи слишком маленькая"
    if x1 <= 0 or y1 <= 0 or x0 >= doc_w_pt or y0 >= doc_h_pt:
        return "область подписи не попадает на слайд"
    return None


def issues(spec: SignatureSpec | None, doc_w_pt: float,
           doc_h_pt: float) -> list[str]:
    """All problems blocking generation with this signature spec."""
    if spec is None or not spec.enabled:
        return []
    out = []
    err = image_error(spec.image_path)
    if err:
        out.append("факсимильная подпись: " + err)
    err = rect_error(spec.rect, doc_w_pt, doc_h_pt)
    if err:
        out.append("факсимильная подпись: " + err)
    return out


def fit_box(img_w: int, img_h: int, rect) -> tuple[float, float, float, float]:
    """(x, y, w, h) in points: ``img`` aspect-fit centered inside ``rect``,
    clamped to the rect — never stretched, never cropped."""
    x0, y0, x1, y1 = rect
    rw, rh = max(0.0, x1 - x0), max(0.0, y1 - y0)
    if img_w <= 0 or img_h <= 0 or rw <= 0 or rh <= 0:
        return (x0, y0, 0.0, 0.0)
    s = min(rw / img_w, rh / img_h)
    w, h = img_w * s, img_h * s
    return (x0 + (rw - w) / 2.0, y0 + (rh - h) / 2.0, w, h)


def clamp_rect(rect, doc_w_pt: float, doc_h_pt: float):
    """Intersect ``rect`` with the page bounds (selection may overflow)."""
    x0, y0, x1, y1 = rect
    return (max(0.0, min(x0, x1)), max(0.0, min(y0, y1)),
            min(doc_w_pt, max(x0, x1)), min(doc_h_pt, max(y0, y1)))
