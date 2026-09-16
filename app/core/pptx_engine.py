"""PPTX template handling with python-pptx.

Text substitution edits existing runs in place, preserving font, size,
color and position of the original placeholder. Automatic font shrinking
keeps substituted text inside the original shape width.
"""

from __future__ import annotations

import copy
import re

from lxml import etree
from pptx.oxml.ns import qn

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu, Pt

from . import fonts as fontlib
from .models import Field, PptxPlacement

EMU_PER_PT = 12700.0
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_TOKEN_RE = re.compile(
    r"(@+[^@\n]*@+|\{+[^{}\n]*\}+|<[^<>\n]*>|_+[^_\n]*_+|%+[^%\n]*%+|\[+[^\[\]\n]*\]+)"
)


def detect_token(text: str) -> str:
    """Find a placeholder token like ``@@FIO@@`` inside shape text."""
    m = _TOKEN_RE.search(text)
    return m.group(0) if m else text


def open_template(path: str) -> Presentation:
    return Presentation(path)


def slide_count(prs: Presentation) -> int:
    return len(prs.slides)


def list_text_shapes(prs: Presentation) -> list[dict]:
    """All shapes that contain non-empty text, per slide."""
    out = []
    for si, slide in enumerate(prs.slides):
        for shp in slide.shapes:
            if not shp.has_text_frame:
                continue
            text = shp.text_frame.text
            if not text.strip():
                continue
            out.append(
                {
                    "slide": si,
                    "shape_id": shp.shape_id,
                    "name": shp.name,
                    "text": text,
                    "left_pt": shp.left / EMU_PER_PT if shp.left is not None else 0.0,
                    "top_pt": shp.top / EMU_PER_PT if shp.top is not None else 0.0,
                    "width_pt": shp.width / EMU_PER_PT if shp.width is not None else 0.0,
                    "height_pt": shp.height / EMU_PER_PT if shp.height is not None else 0.0,
                }
            )
    return out


def find_shape(slide, shape_id: int):
    for shp in slide.shapes:
        if shp.shape_id == shape_id:
            return shp
    return None


def _shape_text_width_pt(shape) -> float:
    tf = shape.text_frame
    ml = tf.margin_left if tf.margin_left is not None else Emu(91440)
    mr = tf.margin_right if tf.margin_right is not None else Emu(91440)
    return (shape.width - ml - mr) / EMU_PER_PT


def _first_run(shape):
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            return run
    return None


def _effective_size_pt(shape, run) -> float:
    if run is not None and run.font.size is not None:
        return run.font.size.pt
    for para in shape.text_frame.paragraphs:
        if para.font.size is not None:
            return para.font.size.pt
    return 18.0


def _run_text(r_el) -> str:
    t = r_el.find(qn("a:t"))
    return t.text or ""


def _set_run_text(r_el, s: str) -> None:
    r_el.find(qn("a:t")).text = s


def _clone_run(r_el, text: str):
    nr = copy.deepcopy(r_el)
    _set_run_text(nr, text)
    return nr


def _rpr_key(r_el) -> str:
    rpr = r_el.find(qn("a:rPr"))
    return etree.tostring(rpr) if rpr is not None else b""


def _replace_in_paragraph(p_el, token: str, value: str):
    """Replace the first ``token`` occurrence inside one ``a:p`` element,
    splitting runs so surrounding text keeps its own formatting.

    Returns (replaced, value_run_el, formats_differ).
    """
    runs = [el for el in p_el if el.tag == qn("a:r")]
    if not runs:
        return False, None, False
    texts = [_run_text(r) for r in runs]
    idx = "".join(texts).find(token)
    if idx < 0:
        return False, None, False
    end = idx + len(token)

    pos = 0
    new_children = []
    touched = []
    value_run = None
    for r, t in zip(runs, texts):
        rs, re_ = pos, pos + len(t)
        pos = re_
        if re_ <= idx or rs >= end:
            new_children.append(r)
            continue
        touched.append(r)
        before = t[: idx - rs] if idx > rs else ""
        after = t[end - rs:] if end < re_ else ""
        if before:
            new_children.append(_clone_run(r, before))
        if value_run is None:
            value_run = _clone_run(r, value)
            new_children.append(value_run)
        # text of later fully-covered runs is dropped (replaced by value)
        if after:
            new_children.append(_clone_run(r, after))

    formats = {_rpr_key(r) for r in touched}
    first_idx = list(p_el).index(runs[0])
    for r in runs:
        p_el.remove(r)
    for i, el in enumerate(new_children):
        p_el.insert(first_idx + i, el)
    return True, value_run, len(formats) > 1


def _replace_fragment(shape, token: str, value: str):
    """Replace all ``token`` occurrences across paragraphs.

    Returns (count, value_run_el, formats_differ).
    """
    count = 0
    value_run = None
    multi = False
    for para in shape.text_frame.paragraphs:
        for _ in range(50):
            ok, vrun, diff = _replace_in_paragraph(para._p, token, value)
            if not ok:
                break
            count += 1
            multi = multi or diff
            if value_run is None:
                value_run = vrun
    return count, value_run, multi


def _run_font(r_el):
    """(size_pt|None, family|None) from the run's rPr."""
    rpr = r_el.find(qn("a:rPr"))
    if rpr is None:
        return None, None
    sz = rpr.get("sz")
    latin = rpr.find(qn("a:latin"))
    return (int(sz) / 100.0 if sz else None,
            latin.get("typeface") if latin is not None else None)


def _set_run_size(r_el, pt: float) -> None:
    rpr = r_el.find(qn("a:rPr"))
    if rpr is None:
        rpr = r_el.makeelement(qn("a:rPr"), {})
        r_el.insert(0, rpr)
    rpr.set("sz", str(int(round(pt * 100))))


def _runs_elements(shape):
    return [
        (r, _run_text(r))
        for para in shape.text_frame.paragraphs
        for r in para._p.findall(qn("a:r"))
    ]


def _set_text_keep_format(shape, new_text: str):
    """Put ``new_text`` into the first run, clear the rest."""
    tf = shape.text_frame
    placed = False
    target_run = None
    for para in tf.paragraphs:
        for run in para.runs:
            if not placed:
                run.text = new_text
                target_run = run
                placed = True
            else:
                run.text = ""
    if target_run is None:
        para = tf.paragraphs[0]
        target_run = para.add_run()
        target_run.text = new_text
    return target_run


def new_text_for(placement: PptxPlacement, value: str) -> str:
    if (
        placement.mode in ("fragment", "token")
        and placement.pattern
        and placement.token
        and placement.token in placement.pattern
    ):
        return placement.pattern.replace(placement.token, value)
    return value


def apply_field_to_slide(slide, placement: PptxPlacement, value: str):
    """Replace placeholder text in one shape. Returns list of review notes."""
    shape = find_shape(slide, placement.shape_id)
    if shape is None or not shape.has_text_frame:
        return ["текстовый блок не найден"]

    notes: list[str] = []
    value_run_el = None   # a:r element holding the inserted value
    fallback_run = None   # Run object for whole-text replacement

    if placement.mode in ("fragment", "token") and placement.token:
        cnt, value_run_el, multi = _replace_fragment(shape, placement.token, value)
        if cnt == 0:
            fallback_run = _set_text_keep_format(shape, value)
            notes.append("токен не найден в тексте блока — заменён весь текст")
        elif multi:
            notes.append(
                "форматирование фрагмента взято от первого run'а, проверьте вручную")
    else:
        new_text = new_text_for(placement, value)
        fallback_run = _set_text_keep_format(shape, new_text)

    box_w = _shape_text_width_pt(shape)
    if box_w <= 0:
        return notes

    if value_run_el is not None:
        # fragment mode: shrink the value run until the composed line fits
        infos = _runs_elements(shape)
        others_w = 0.0
        orig_size, fname = 18.0, None
        for r_el, text in infos:
            sz, name = _run_font(r_el)
            fp = fontlib.find_font_file(name or "")
            if r_el is value_run_el:
                orig_size = sz or orig_size
                fname = name
            else:
                others_w += fontlib.measure_text_width(text, fp, sz or 18.0)
        font_path = fontlib.find_font_file(fname or "")
        size = orig_size
        min_size = max(4.0, size * fontlib.MIN_FONT_RATIO)
        while size > min_size and others_w + fontlib.measure_text_width(
                value, font_path, size) > box_w:
            size -= 1.0 if size > 12 else 0.5
        fits = others_w + fontlib.measure_text_width(value, font_path, size) <= box_w
        if size < orig_size:
            _set_run_size(value_run_el, size)
            notes.append(f"шрифт уменьшен до {size:.1f} pt")
        if not fits:
            notes.append(f"текст не поместился даже при минимальном размере ({size:.1f} pt)")
        return notes

    run = fallback_run
    new_text = shape.text_frame.text
    orig_size = _effective_size_pt(shape, run)
    font_path = fontlib.find_font_file(run.font.name or "")
    allow_wrap = bool(shape.text_frame.word_wrap)
    new_size, fits = fontlib.fit_font_size(
        new_text, font_path, box_w, orig_size, allow_wrap=allow_wrap
    )
    if new_size < orig_size:
        run.font.size = Pt(new_size)
    if not fits:
        notes.append(f"текст не поместился даже при минимальном размере ({new_size:.1f} pt)")
    elif new_size < orig_size:
        notes.append(f"шрифт уменьшен до {new_size:.1f} pt")
    return notes


def substitute(prs: Presentation, fields: list[Field], values: dict[str, str]):
    """Apply field values to slides of a presentation in place.

    ``values`` maps field name -> text. Returns list of review notes
    (strings) — non-fatal issues like shrunken fonts.
    """
    notes = []
    slides = list(prs.slides)
    for f in fields:
        value = values.get(f.name, "")
        for pl in f.placements:
            if not isinstance(pl, PptxPlacement):
                continue
            if pl.slide >= len(slides):
                notes.append(f"{f.name}: слайд {pl.slide + 1} отсутствует")
                continue
            for note in apply_field_to_slide(slides[pl.slide], pl, value):
                notes.append(f"{f.name}: {note}")
    return notes


def save_prs(prs: Presentation, path_or_buf) -> None:
    prs.save(path_or_buf)


def render_single(template_path: str, fields: list[Field], values: dict[str, str], out_path: str):
    """Generate one PPTX file for one data row. Returns review notes."""
    prs = Presentation(template_path)
    notes = substitute(prs, fields, values)
    prs.save(out_path)
    return notes


def duplicate_slide(prs: Presentation, slide):
    """Append a deep copy of ``slide`` (incl. image rels) to ``prs``."""
    layout = slide.slide_layout
    new_slide = prs.slides.add_slide(layout)
    for shp in list(new_slide.shapes):
        shp.element.getparent().remove(shp.element)

    id_map: dict[str, str] = {}
    for r_id, rel in slide.part.rels.items():
        if rel.reltype in (RT.NOTES_SLIDE, RT.SLIDE_LAYOUT):
            continue
        try:
            if rel.is_external:
                new_id = new_slide.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
            else:
                new_id = new_slide.part.rels.get_or_add(rel.reltype, rel.target_part)
            id_map[r_id] = new_id
        except Exception:
            continue

    for shp in slide.shapes:
        new_el = copy.deepcopy(shp.element)
        for el in new_el.iter():
            for attr in list(el.attrib):
                if attr.startswith("{" + _REL_NS + "}"):
                    v = el.attrib[attr]
                    if v in id_map:
                        el.attrib[attr] = id_map[v]
        new_slide.shapes._spTree.append(new_el)
    return new_slide


def render_combined(
    template_path: str, fields: list[Field], rows_values: list[dict[str, str]], out_path: str
):
    """One PPTX where each data row contributes a copy of the whole deck."""
    prs = Presentation(template_path)
    template_slides = list(prs.slides)
    per_row_slides: list[list] = [template_slides]
    for _ in range(1, len(rows_values)):
        per_row_slides.append([duplicate_slide(prs, s) for s in template_slides])

    notes_per_row: list[list[str]] = []
    for slides, values in zip(per_row_slides, rows_values):
        notes = []
        for f in fields:
            value = values.get(f.name, "")
            for pl in f.placements:
                if not isinstance(pl, PptxPlacement):
                    continue
                if pl.slide >= len(slides):
                    notes.append(f"{f.name}: слайд {pl.slide + 1} отсутствует")
                    continue
                for note in apply_field_to_slide(slides[pl.slide], pl, value):
                    notes.append(f"{f.name}: {note}")
        notes_per_row.append(notes)
    prs.save(out_path)
    return notes_per_row
