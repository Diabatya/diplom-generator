"""System font discovery and offline text measurement via PyMuPDF fonts."""

from __future__ import annotations

import os
import sys
from functools import lru_cache

try:
    import pymupdf as fitz
except ImportError:
    import fitz  # type: ignore  # PyMuPDF

FONT_EXTS = (".ttf", ".otf", ".ttc")

# common Windows font file names for popular families
_KNOWN_FILES = {
    "arial": ["arial.ttf", "arial"],
    "arial black": ["ariblk.ttf"],
    "calibri": ["calibri.ttf", "calibri"],
    "cambria": ["cambria.ttf"],
    "consolas": ["consola.ttf"],
    "courier new": ["cour.ttf"],
    "georgia": ["georgia.ttf"],
    "impact": ["impact.ttf"],
    "times new roman": ["times.ttf", "times"],
    "trebuchet ms": ["trebuc.ttf"],
    "verdana": ["verdana.ttf"],
    "tahoma": ["tahoma.ttf"],
    "segoe ui": ["segoeui.ttf"],
    "comic sans ms": ["comic.ttf"],
    "garamond": ["garamond.ttf"],
    "book antiqua": ["bkant.ttf", "bookos.ttf"],
    "century gothic": ["gothic.ttf"],
    "dejavu sans": ["DejaVuSans.ttf"],
    "dejavu serif": ["DejaVuSerif.ttf"],
    "liberation sans": ["LiberationSans-Regular.ttf"],
    "liberation serif": ["LiberationSerif-Regular.ttf"],
}


def system_font_dirs() -> list[str]:
    if sys.platform == "win32":
        dirs = [os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")]
        local = os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Microsoft\Windows\Fonts")
        dirs.append(local)
    elif sys.platform == "darwin":
        dirs = [
            "/System/Library/Fonts",
            "/System/Library/Fonts/Supplemental",
            "/Library/Fonts",
            os.path.expanduser("~/Library/Fonts"),
        ]
    else:
        dirs = ["/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.fonts")]
    return [d for d in dirs if d and os.path.isdir(d)]


@lru_cache(maxsize=1)
def _font_files() -> tuple[str, ...]:
    """All system font files, scanned once (Windows font dirs are huge —
    rescanning per find_font_file call was a measurable slowdown)."""
    out = []
    for d in system_font_dirs():
        for root, _sub, files in os.walk(d):
            for fn in files:
                if fn.lower().endswith(FONT_EXTS):
                    out.append(os.path.join(root, fn))
    return tuple(sorted(set(out)))


def _iter_font_files() -> list[str]:
    return list(_font_files())


@lru_cache(maxsize=1)
def _stem_map() -> dict[str, str]:
    return {os.path.splitext(os.path.basename(p))[0].lower(): p
            for p in _font_files()}


def font_cache_clear() -> None:
    """Rescan fonts (e.g. after the user installs a font mid-session)."""
    _font_files.cache_clear()
    _stem_map.cache_clear()
    _font_obj.cache_clear()


def get_font(font_path: str | None) -> fitz.Font | None:
    return _font_obj(font_path) if font_path else None


@lru_cache(maxsize=512)
def _font_obj(font_path: str) -> fitz.Font | None:
    try:
        return fitz.Font(fontfile=font_path)
    except Exception:
        return None


def has_cyrillic(font_path: str) -> bool:
    f = _font_obj(font_path)
    if f is None:
        return False
    try:
        return bool(f.has_glyph(ord("Д"))) and bool(f.has_glyph(ord("ф")))
    except Exception:
        # fallback: measuring a cyrillic string must produce non-zero width
        try:
            return f.text_length("Дф", fontsize=10) > 0
        except Exception:
            return False


def list_system_fonts(cyrillic_only: bool = True) -> list[tuple[str, str]]:
    """Return [(display_name, path)] of usable fonts."""
    result = []
    for path in _iter_font_files():
        if cyrillic_only and not has_cyrillic(path):
            continue
        result.append((os.path.splitext(os.path.basename(path))[0], path))
    return result


def _norm(name: str) -> str:
    return "".join(ch for ch in name.lower() if ch.isalnum())


def find_font_file(font_name: str) -> str | None:
    """Best-effort mapping of a font family name to a system font file."""
    if not font_name:
        return None
    if os.path.isfile(font_name):
        return font_name
    key = font_name.strip().lower()
    files = _font_files()
    stems = _stem_map()
    for cand in _KNOWN_FILES.get(key, []):
        stem = os.path.splitext(cand)[0].lower()
        if stem in stems:
            return stems[stem]
    norm_key = _norm(font_name)
    for p in files:
        if _norm(os.path.splitext(os.path.basename(p))[0]) == norm_key:
            return p
    # relaxed: filename contains the family name
    for p in files:
        if norm_key and norm_key in _norm(os.path.basename(p)):
            return p
    return None


def measure_text_width(text: str, font_path: str | None, size_pt: float) -> float:
    """Width of a single line of text in points. Heuristic fallback."""
    f = _font_obj(font_path) if font_path else None
    if f is not None:
        try:
            return f.text_length(text, fontsize=size_pt)
        except Exception:
            pass
    # heuristic: cyrillic/latin mix ~0.55 em average
    return size_pt * 0.55 * len(text)


def line_height(font_path: str | None, size_pt: float) -> float:
    f = _font_obj(font_path) if font_path else None
    if f is not None:
        try:
            return size_pt * (f.ascender - f.descender)
        except Exception:
            pass
    return size_pt * 1.2


def ascender(font_path: str | None) -> float:
    f = _font_obj(font_path) if font_path else None
    if f is not None:
        try:
            return f.ascender
        except Exception:
            pass
    return 0.8


MIN_FONT_RATIO = 0.6  # do not shrink below 60% of the original size


def fit_font_size(
    text: str,
    font_path: str | None,
    box_width_pt: float,
    original_size: float,
    allow_wrap: bool = False,
) -> tuple[float, bool]:
    """Reduce size until ``text`` fits ``box_width_pt``.

    Returns (size, fits). ``fits`` is False when even the minimum size does
    not fit (caller marks the row as "requires manual review").
    """
    size = float(original_size)
    min_size = max(4.0, size * MIN_FONT_RATIO)

    def ok(s: float) -> bool:
        if allow_wrap:
            return all(
                measure_text_width(w, font_path, s) <= box_width_pt
                for w in text.split() or [text]
            )
        return measure_text_width(text, font_path, s) <= box_width_pt

    while size > min_size and not ok(size):
        size -= 1.0 if size > 12 else 0.5
    return size, ok(size)
