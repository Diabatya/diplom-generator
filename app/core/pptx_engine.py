"""PPTX template handling with python-pptx.

Text substitution edits existing runs in place, preserving font, size,
color and position of the original placeholder. Automatic font shrinking
keeps substituted text inside the original shape width.
"""

from __future__ import annotations

import copy
import re

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
    """Replace placeholder text in one shape. Returns review note or None."""
    shape = find_shape(slide, placement.shape_id)
    if shape is None or not shape.has_text_frame:
        return "текстовый блок не найден"
    new_text = new_text_for(placement, value)
    run = _set_text_keep_format(shape, new_text)
    box_w = _shape_text_width_pt(shape)
    if box_w <= 0:
        return None
    orig_size = _effective_size_pt(shape, run)
    font_path = fontlib.find_font_file(run.font.name or "")
    allow_wrap = bool(shape.text_frame.word_wrap)
    new_size, fits = fontlib.fit_font_size(
        new_text, font_path, box_w, orig_size, allow_wrap=allow_wrap
    )
    if new_size < orig_size:
        run.font.size = Pt(new_size)
    if not fits:
        return f"текст не поместился даже при минимальном размере ({new_size:.1f} pt)"
    if new_size < orig_size:
        return f"шрифт уменьшен до {new_size:.1f} pt"
    return None


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
            note = apply_field_to_slide(slides[pl.slide], pl, value)
            if note:
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
                note = apply_field_to_slide(slides[pl.slide], pl, value)
                if note:
                    notes.append(f"{f.name}: {note}")
        notes_per_row.append(notes)
    prs.save(out_path)
    return notes_per_row
