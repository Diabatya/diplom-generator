"""PPTX template handling with python-pptx.

Text substitution edits existing runs in place, preserving font, size,
color and position of the original placeholder. Automatic font shrinking
keeps substituted text inside the original shape width.
"""

from __future__ import annotations

import copy
import re
from functools import lru_cache

from lxml import etree
from pptx.oxml.ns import qn

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu, Pt

from . import fonts as fontlib
from .models import Field, PptxPlacement

EMU_PER_PT = 12700.0
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_TOKEN_RE = re.compile(r"(@{2,}[^@\n]+@{2,})")


def detect_token(text: str) -> str:
    """Find an explicit ``@@TOKEN@@`` placeholder inside shape text.

    Only double-@ markers are auto-detected: other bracketed text
    (``<номер>``, ``[название]``, ``_x_``) is ordinary text too often to
    guess safely. Those can still be bound via fragment mode.
    Returns "" when nothing is found.
    """
    m = _TOKEN_RE.search(text)
    return m.group(0) if m else ""


def open_template(path: str) -> Presentation:
    return Presentation(path)


def slide_count(prs: Presentation) -> int:
    return len(prs.slides)


def _grp_xfrm(shp):
    """Group child-transform: (off_x, off_y, ch_off_x, ch_off_y, sx, sy) EMU."""
    try:
        xfrm = shp.element.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
        off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
        choff, chext = xfrm.find(qn("a:chOff")), xfrm.find(qn("a:chExt"))
        gsx = int(ext.get("cx")) / max(1, int(chext.get("cx")))
        gsy = int(ext.get("cy")) / max(1, int(chext.get("cy")))
        return (int(off.get("x")), int(off.get("y")),
                int(choff.get("x")), int(choff.get("y")), gsx, gsy)
    except Exception:
        return None


def _walk_shapes(shapes, dx=0.0, dy=0.0, sx=1.0, sy=1.0):
    """Yield (shape, off_x, off_y, scale_x, scale_y) — absolute position of a
    shape at child-coords (0,0) is (off_x, off_y), sizes multiply by scale.
    Recurses into group shapes applying their child transform."""
    for shp in shapes:
        yield shp, dx, dy, sx, sy
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
            tr = _grp_xfrm(shp)
            if tr:
                ox, oy, cox, coy, gsx, gsy = tr
                yield from _walk_shapes(
                    shp.shapes,
                    dx + (ox - cox * gsx) * sx, dy + (oy - coy * gsy) * sy,
                    sx * gsx, sy * gsy)


def _shape_entry(si, shp, text, left_emu, top_emu, w_emu, h_emu, cell=None):
    return {
        "slide": si,
        "shape_id": shp.shape_id,
        "name": shp.name,
        "text": text,
        "cell": cell,
        "left_pt": left_emu / EMU_PER_PT,
        "top_pt": top_emu / EMU_PER_PT,
        "width_pt": w_emu / EMU_PER_PT,
        "height_pt": h_emu / EMU_PER_PT,
    }


def list_text_shapes(prs: Presentation) -> list[dict]:
    """All shapes that contain non-empty text, per slide — including
    shapes inside groups (bounds in absolute slide points) and cells of
    tables, each cell addressed by ``cell=(row, col)``."""
    out = []
    for si, slide in enumerate(prs.slides):
        for shp, dx, dy, sx, sy in _walk_shapes(slide.shapes):
            left = shp.left or 0
            top = shp.top or 0
            if getattr(shp, "has_table", False) and shp.has_table:
                col_x = 0
                for ci, col in enumerate(shp.table.columns):
                    row_y = 0
                    for ri, row in enumerate(shp.table.rows):
                        cell = shp.table.cell(ri, ci)
                        if cell.text.strip():
                            out.append(_shape_entry(
                                si, shp, cell.text,
                                dx + (left + col_x) * sx,
                                dy + (top + row_y) * sy,
                                col.width * sx, row.height * sy,
                                (ri, ci)))
                        row_y += row.height
                    col_x += col.width
            elif shp.has_text_frame and shp.text_frame.text.strip():
                out.append(_shape_entry(
                    si, shp, shp.text_frame.text,
                    dx + left * sx, dy + top * sy,
                    (shp.width or 0) * sx, (shp.height or 0) * sy))
    return out


def find_shape(slide, shape_id: int):
    for shp, *_ in _walk_shapes(slide.shapes):
        if shp.shape_id == shape_id:
            return shp
    return None


def _shape_text_width_pt(shape) -> float:
    tf = shape.text_frame
    ml = tf.margin_left if tf.margin_left is not None else Emu(91440)
    mr = tf.margin_right if tf.margin_right is not None else Emu(91440)
    return (shape.width - ml - mr) / EMU_PER_PT


def _shape_text_height_pt(shape) -> float:
    tf = shape.text_frame
    mt = tf.margin_top if tf.margin_top is not None else Emu(45720)
    mb = tf.margin_bottom if tf.margin_bottom is not None else Emu(45720)
    return (shape.height - mt - mb) / EMU_PER_PT


def _target_frame_and_box(slide, placement: PptxPlacement):
    """Resolve the text_frame and (width_pt, height_pt) for a placement —
    either a shape's text frame or a table cell. Sizes are in the shape's
    local (child) coordinate space: a group transform scales the rendered
    text and the box by the same factor, so it cancels out of the fit
    math. Slide-absolute bounds are only needed by list_text_shapes for
    the UI overlay."""
    shape = find_shape(slide, placement.shape_id)
    if shape is None:
        return None, None, 0.0, 0.0
    if placement.cell is not None:
        if not (getattr(shape, "has_table", False) and shape.has_table):
            return shape, None, 0.0, 0.0
        r, c = placement.cell
        try:
            cell = shape.table.cell(r, c)
        except IndexError:
            return shape, None, 0.0, 0.0
        ml = cell.margin_left if cell.margin_left is not None else Emu(91440)
        mr = cell.margin_right if cell.margin_right is not None else Emu(91440)
        mt = cell.margin_top if cell.margin_top is not None else Emu(45720)
        mb = cell.margin_bottom if cell.margin_bottom is not None else Emu(45720)
        w = (shape.table.columns[c].width - ml - mr) / EMU_PER_PT
        h = (shape.table.rows[r].height - mt - mb) / EMU_PER_PT
        return shape, cell.text_frame, w, h
    if not shape.has_text_frame:
        return shape, None, 0.0, 0.0
    return shape, shape.text_frame, \
        _shape_text_width_pt(shape), _shape_text_height_pt(shape)


def required_placement_issues(prs: Presentation, fields: list[Field], slide_sel=None) -> list[str]:
    selected = set(_selected_indexes(prs, slide_sel))
    issues = []
    for field in fields:
        if not field.required:
            continue
        valid = False
        for pl in field.placements:
            if not isinstance(pl, PptxPlacement) or pl.slide not in selected:
                continue
            _shape, tf, _width, _height = _target_frame_and_box(prs.slides[pl.slide], pl)
            if tf is None:
                continue
            if pl.mode in ("fragment", "token"):
                if not pl.token or not any(
                    pl.token in "".join(text for _run, text in runs)
                    for runs in _para_run_elements(tf)
                ):
                    continue
            valid = True
            break
        if not valid:
            slides = ", ".join(str(i + 1) for i in sorted(selected))
            issues.append(f"обязательное поле «{field.name}»: нет действующей привязки "
                          f"на выбранных слайдах ({slides}); проверьте блок и фрагмент")
    return issues


def _first_run(tf):
    for para in tf.paragraphs:
        for run in para.runs:
            return run
    return None


def _effective_size_pt(tf, run) -> float:
    if run is not None and run.font.size is not None:
        return run.font.size.pt
    for para in tf.paragraphs:
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


def _replace_in_paragraph(p_el, token: str, value: str, start: int = 0):
    """Replace the first ``token`` occurrence inside one ``a:p`` element,
    splitting runs so surrounding text keeps its own formatting.

    Returns (replaced, value_run_el, formats_differ).
    """
    runs = [el for el in p_el if el.tag == qn("a:r")]
    if not runs:
        return False, None, False, start
    texts = [_run_text(r) for r in runs]
    idx = "".join(texts).find(token, start)
    if idx < 0:
        return False, None, False, start
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
    return True, value_run, len(formats) > 1, idx + len(value)


def _replace_fragment(tf, token: str, value: str):
    """Replace all ``token`` occurrences across paragraphs.

    Returns (count, [value_run_el, ...], formats_differ).
    """
    count = 0
    value_runs: list = []
    multi = False
    for para in tf.paragraphs:
        start = 0
        while True:
            ok, vrun, diff, start = _replace_in_paragraph(para._p, token, value, start)
            if not ok:
                break
            count += 1
            multi = multi or diff
            if vrun is not None:
                value_runs.append(vrun)
    return count, value_runs, multi


def _run_font(r_el):
    """(size_pt|None, family|None) from the run's rPr."""
    rpr = r_el.find(qn("a:rPr"))
    if rpr is None:
        return None, None
    sz = rpr.get("sz")
    latin = rpr.find(qn("a:latin"))
    return (int(sz) / 100.0 if sz else None,
            latin.get("typeface") if latin is not None else None)


@lru_cache(maxsize=32)
def _theme_font(blob: bytes, family: str) -> str | None:
    root = etree.fromstring(blob)
    scheme = root.find(".//" + qn("a:fontScheme"))
    if scheme is None:
        return None
    branch = scheme.find(qn("a:majorFont" if family.startswith("+mj-") else "a:minorFont"))
    if branch is None:
        return None
    cyrillic = next((f.get("typeface") for f in branch.findall(qn("a:font"))
                     if f.get("script") == "Cyrl" and f.get("typeface")), None)
    latin = branch.find(qn("a:latin"))
    return cyrillic or (latin.get("typeface") if latin is not None else None)


def _effective_run_style(slide, shape, tf, para, r_el):
    level = min(9, max(1, para.level + 1))
    candidates = [r_el.find(qn("a:rPr"))]
    ppr = para._p.find(qn("a:pPr"))
    candidates.append(ppr.find(qn("a:defRPr")) if ppr is not None else None)
    body = getattr(tf, "_txBody", None)
    style = body.find(qn("a:lstStyle")) if body is not None else None
    lvl = style.find(qn(f"a:lvl{level}pPr")) if style is not None else None
    candidates.append(lvl.find(qn("a:defRPr")) if lvl is not None else None)
    layout = slide.slide_layout
    master = layout.slide_master
    if shape.is_placeholder:
        idx = shape.placeholder_format.idx
        for parent in (layout, master):
            placeholder = next((s for s in parent.placeholders
                                if s.placeholder_format.idx == idx), None)
            if placeholder is not None and placeholder.has_text_frame:
                p = placeholder.text_frame.paragraphs[0]._p.find(qn("a:pPr"))
                candidates.append(p.find(qn("a:defRPr")) if p is not None else None)
                parent_body = placeholder.text_frame._txBody
                parent_style = parent_body.find(qn("a:lstStyle"))
                parent_lvl = (parent_style.find(qn(f"a:lvl{level}pPr"))
                              if parent_style is not None else None)
                candidates.append(parent_lvl.find(qn("a:defRPr"))
                                  if parent_lvl is not None else None)
    tx_styles = master.element.find(qn("p:txStyles"))
    if tx_styles is not None:
        if shape.is_placeholder:
            ph_name = shape.placeholder_format.type.name
            key = "titleStyle" if ph_name in ("TITLE", "CENTER_TITLE") else "bodyStyle"
        else:
            # Text in shapes and text boxes is styled by otherStyle.
            key = "otherStyle"
        branch = tx_styles.find(qn(f"p:{key}"))
        lvl = branch.find(qn(f"a:lvl{level}pPr")) if branch is not None else None
        candidates.append(lvl.find(qn("a:defRPr")) if lvl is not None else None)
    size, family, bold = None, None, None
    for rpr in candidates:
        if rpr is None:
            continue
        if size is None and rpr.get("sz"):
            size = int(rpr.get("sz")) / 100.0
        if bold is None and rpr.get("b") is not None:
            bold = rpr.get("b") in ("1", "true")
        latin = rpr.find(qn("a:latin"))
        if family is None and latin is not None:
            family = latin.get("typeface")
    if family and family.startswith(("+mn-", "+mj-")):
        relation = next((rel for rel in master.part.rels.values()
                         if rel.reltype == RT.THEME), None)
        family = _theme_font(relation.target_part.blob, family) if relation else None
    return size, family, bool(bold)


def _set_run_size(r_el, pt: float) -> None:
    rpr = r_el.find(qn("a:rPr"))
    if rpr is None:
        rpr = r_el.makeelement(qn("a:rPr"), {})
        r_el.insert(0, rpr)
    rpr.set("sz", str(int(round(pt * 100))))


def _para_run_elements(tf):
    """Run elements grouped by paragraph: [[(r_el, text), ...], ...]."""
    return [
        [(r, _run_text(r)) for r in para._p.findall(qn("a:r"))]
        for para in tf.paragraphs
    ]


def _runs_elements(tf):
    return [rt for para in _para_run_elements(tf) for rt in para]


def _set_text_keep_format(tf, new_text: str):
    """Put ``new_text`` into the first run, clear the rest."""
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


def _est_height_pt(tf, box_w, value_ids=None, vsize=None, vfp=None, styles=None) -> float:
    """Estimate line heights with wrapping, run weights and paragraph spacing."""
    total = 0.0
    for para, runs in zip(tf.paragraphs, _para_run_elements(tf)):
        line_w, line_sz, line_fp = 0.0, 0.0, None
        spacing = para.line_spacing
        def advance(fp, sz):
            base = fontlib.line_height(fp, sz)
            if isinstance(spacing, float):
                return base * spacing
            return spacing.pt if spacing is not None else base
        total += (para.space_before.pt if para.space_before else 0) + (
            para.space_after.pt if para.space_after else 0)
        any_text = False
        for r_el, text in runs:
            sz, name, bold = styles[r_el] if styles else (*_run_font(r_el), False)
            fp = fontlib.find_font_file(name or "")
            if value_ids and r_el in value_ids:
                sz, fp = vsize, vfp
            sz = sz or 18.0
            for word in re.split(r"(\s+)", text):
                if not word:
                    continue
                any_text = True
                w = fontlib.measure_text_width(word, fp, sz) * (1.06 if bold else 1.0)
                if tf.word_wrap and line_w + w > box_w and line_w > 0:
                    total += advance(line_fp, line_sz or 18.0)
                    line_w, line_sz, line_fp = 0.0, 0.0, None
                    if word.isspace():
                        continue
                if sz >= line_sz:
                    line_sz, line_fp = sz, fp
                line_w += w
        if line_w > 0 or not any_text:
            total += advance(line_fp, line_sz or 18.0)
    return total


def _fit_block(text, fp, box_w, box_h, orig, wrap, tf=None, bold=False):
    """Whole-block fit: shrink until width AND height fit. Returns (size, ok)."""
    size = float(orig)
    min_s = max(4.0, size * fontlib.MIN_FONT_RATIO)
    paras = text.split("\n")
    def width(t, s):
        return fontlib.measure_text_width(t, fp, s) * (1.06 if bold else 1.0)

    def width_ok(s):
        if wrap:
            return all(
                width(w, s) <= box_w
                for p in paras for w in (p.split() or [p]))
        return all(width(p, s) <= box_w for p in paras)

    def height_ok(s):
        if box_h <= 0:
            return True
        lh = fontlib.line_height(fp, s)
        h = 0.0
        for index, para in enumerate(paras):
            style = tf.paragraphs[index] if tf and index < len(tf.paragraphs) else None
            spacing = style.line_spacing if style else None
            line_h = lh * spacing if isinstance(spacing, float) else (
                spacing.pt if spacing is not None else lh)
            if style:
                h += (style.space_before.pt if style.space_before else 0) + (
                    style.space_after.pt if style.space_after else 0)
            if not wrap:
                h += line_h
                continue
            lw, lines = 0.0, 0
            for word in re.split(r"(\s+)", para):
                if not word:
                    continue
                w = width(word, s)
                if lw + w > box_w and lw > 0:
                    lines += 1
                    lw = 0.0
                    if word.isspace():
                        continue
                lw += w
            lines += 1 if lw > 0 or para == "" else 0
            h += lines * line_h
        return h <= box_h

    while size > min_s and not (width_ok(size) and height_ok(size)):
        size -= 1.0 if size > 12 else 0.5
    return size, width_ok(size) and height_ok(size)


def apply_field_to_slide(slide, placement: PptxPlacement, value: str):
    """Replace placeholder text in one shape or table cell.
    Returns list of review notes."""
    shape, tf, box_w, box_h = _target_frame_and_box(slide, placement)
    if tf is None:
        return ["текстовый блок не найден"]

    notes: list[str] = []
    value_run_els: list = []   # a:r elements holding the inserted value
    fallback_run = None        # Run object for whole-text replacement

    if placement.mode in ("fragment", "token") and placement.token:
        cnt, value_run_els, multi = _replace_fragment(tf, placement.token, value)
        if cnt == 0:
            # never silently wipe surrounding text on mismatched markup
            return [f"токен «{placement.token[:40]}» не найден в блоке — "
                    "разметка не совпадает с шаблоном, проверьте привязку"]
        if multi:
            notes.append(
                "форматирование фрагмента взято от первого run'а, проверьте вручную")
    else:
        new_text = new_text_for(placement, value)
        fallback_run = _set_text_keep_format(tf, new_text)

    if box_w <= 0:
        return notes

    if value_run_els:
        # fragment mode: shrink value runs until the composed block fits
        # both width and height. Other value runs (the same token repeated)
        # are fully replaced text — not "surrounding" width.
        value_ids = set(value_run_els)
        styles = {r_el: _effective_run_style(slide, shape, tf, para, r_el)
                  for para in tf.paragraphs for r_el in para._p.findall(qn("a:r"))}
        measured_size, fname, _bold = styles[value_run_els[0]]
        orig_size = measured_size or 18.0
        font_path = fontlib.find_font_file(fname or "")
        if measured_size is None or not fname:
            notes.append("размер или гарнитура шрифта не определены — расчёт приблизительный")
        elif not font_path:
            notes.append(f"шрифт «{fname}» не найден в системе — "
                         "расчёт автоподбора приблизительный")
        size = orig_size
        min_size = max(4.0, size * fontlib.MIN_FONT_RATIO)

        def fits(s):
            for runs in _para_run_elements(tf):
                widths = []
                for r_el, text in runs:
                    sz, name, bold = styles[r_el]
                    if r_el in value_ids:
                        sz, fp = s, font_path
                    else:
                        sz, fp = sz or 18.0, fontlib.find_font_file(name or "")
                    weight = 1.06 if bold else 1.0
                    widths.append(fontlib.measure_text_width(text, fp, sz) * weight)
                    if tf.word_wrap and any(fontlib.measure_text_width(word, fp, sz) * weight > box_w
                                            for word in text.split()):
                        return False
                if not tf.word_wrap and sum(widths) > box_w:
                    return False
            return box_h <= 0 or _est_height_pt(
                tf, box_w, value_ids, s, font_path, styles) <= box_h + 0.5

        while size > min_size and not fits(size):
            size -= 1.0 if size > 12 else 0.5
        ok = fits(size)
        if size < orig_size:
            for r_el in value_run_els:
                _set_run_size(r_el, size)
            notes.append(f"шрифт уменьшен до {size:.1f} pt")
        if not ok:
            notes.append(f"текст не поместился даже при минимальном размере ({size:.1f} pt)")
        return notes

    run = fallback_run
    new_text = tf.text
    para = next((p for p in tf.paragraphs if run._r.getparent() is p._p), tf.paragraphs[0])
    measured_size, fname, bold = _effective_run_style(slide, shape, tf, para, run._r)
    orig_size = measured_size or 18.0
    font_path = fontlib.find_font_file(fname or "")
    if measured_size is None or not fname:
        notes.append("размер или гарнитура шрифта не определены — расчёт приблизительный")
    elif not font_path:
        notes.append(f"шрифт «{fname}» не найден в системе — "
                     "расчёт автоподбора приблизительный")
    allow_wrap = bool(tf.word_wrap)
    new_size, fits = _fit_block(
        new_text, font_path, box_w, box_h, orig_size, allow_wrap, tf, bold)
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
        if f.name not in values:
            continue
        value = values[f.name]
        for pl in f.placements:
            if not isinstance(pl, PptxPlacement):
                continue
            if pl.slide >= len(slides):
                notes.append(f"{f.name}: слайд {pl.slide + 1} отсутствует")
                continue
            for note in apply_field_to_slide(slides[pl.slide], pl, value):
                notes.append(f"{f.name}, слайд {pl.slide + 1}: {note}")
    return notes


def save_prs(prs: Presentation, path_or_buf) -> None:
    prs.save(path_or_buf)


def _drop_all_but(prs: Presentation, orig_sld_ids, keep: set[int]):
    """Remove unselected ORIGINAL slides (by their index in the template)
    from the presentation's slide list and relationships."""
    sldIdLst = prs.slides._sldIdLst
    for i, sldId in enumerate(orig_sld_ids):
        if i not in keep:
            try:
                prs.part.drop_rel(sldId.get(qn("r:id")))
            except Exception:
                pass
            sldIdLst.remove(sldId)


def _selected_indexes(prs: Presentation, slide_sel) -> list[int]:
    n = len(prs.slides)
    if slide_sel is None:
        return list(range(n))
    keep = [i for i in dict.fromkeys(slide_sel) if 0 <= i < n]
    return keep or [0]


def render_single(template_path: str, fields: list[Field], values: dict[str, str],
                  out_path: str, slide_sel=None):
    """Generate one PPTX file for one data row. Returns review notes.

    ``slide_sel`` limits the output to the given template slides (original
    indexes); None keeps the whole deck."""
    prs = Presentation(template_path)
    notes = substitute(prs, fields, values)
    if slide_sel is not None:
        keep = set(_selected_indexes(prs, slide_sel))
        _drop_all_but(prs, list(prs.slides._sldIdLst), keep)
    prs.save(out_path)
    return notes


def duplicate_slide(prs: Presentation, slide):
    """Append a deep copy of ``slide`` (incl. image rels AND the slide-level
    background — solid fill, gradient or picture) to ``prs``."""
    layout = slide.slide_layout
    new_slide = prs.slides.add_slide(layout)
    for attr in ("showMasterSp", "showMasterPhAnim", "show"):
        if attr in slide.element.attrib:
            new_slide.element.set(attr, slide.element.get(attr))
        else:
            new_slide.element.attrib.pop(attr, None)
    src_color = slide.element.find(qn("p:clrMapOvr"))
    dst_color = new_slide.element.find(qn("p:clrMapOvr"))
    if dst_color is not None:
        new_slide.element.remove(dst_color)
    if src_color is not None:
        new_slide.element.insert(1, copy.deepcopy(src_color))
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

    def _remap_rels(el):
        for sub in el.iter():
            for attr in list(sub.attrib):
                if attr.startswith("{" + _REL_NS + "}"):
                    v = sub.attrib[attr]
                    if v in id_map:
                        sub.attrib[attr] = id_map[v]

    for shp in slide.shapes:
        new_el = copy.deepcopy(shp.element)
        _remap_rels(new_el)
        new_slide.shapes._spTree.append(new_el)

    # slide-level background (<p:bg> inside <p:cSld>): solid fill, gradient,
    # picture (its r:embed is remapped through id_map as well)
    src_csld = slide.element.find(qn("p:cSld"))
    dst_csld = new_slide.element.find(qn("p:cSld"))
    if src_csld is not None and dst_csld is not None:
        src_bg = src_csld.find(qn("p:bg"))
        if src_bg is not None:
            bg = copy.deepcopy(src_bg)
            _remap_rels(bg)
            dst_bg = dst_csld.find(qn("p:bg"))
            if dst_bg is not None:
                dst_csld.remove(dst_bg)
            dst_csld.insert(0, bg)   # bg is the first child of cSld
    return new_slide


def render_combined(
    template_path: str, fields: list[Field], rows_values: list[dict[str, str]],
    out_path: str, slide_sel=None, progress_cb=None, cancel_cb=None,
):
    """One PPTX where each data row contributes a copy of the selected
    template slides (``slide_sel`` = original indexes; None = whole deck).

    ``cancel_cb()`` is expected to raise to abort; ``progress_cb(i)`` is
    called per row. Returns per-row lists of review notes."""
    prs = Presentation(template_path)
    template_slides = list(prs.slides)
    keep = _selected_indexes(prs, slide_sel)
    orig_ids = list(prs.slides._sldIdLst)

    sel = [template_slides[i] for i in keep]
    per_row_slides: list[list] = [sel]
    for _ in range(1, len(rows_values)):
        if cancel_cb:
            cancel_cb()
        per_row_slides.append([duplicate_slide(prs, s) for s in sel])

    remap = {orig: pos for pos, orig in enumerate(keep)}
    notes_per_row: list[list[str]] = []
    for ri, (slides, values) in enumerate(zip(per_row_slides, rows_values)):
        if cancel_cb:
            cancel_cb()
        if progress_cb:
            progress_cb(ri)
        notes = []
        for f in fields:
            if f.name not in values:
                continue
            value = values[f.name]
            for pl in f.placements:
                if not isinstance(pl, PptxPlacement):
                    continue
                if pl.slide not in remap:
                    if slide_sel is None:
                        notes.append(f"{f.name}: слайд {pl.slide + 1} отсутствует")
                    continue
                for note in apply_field_to_slide(slides[remap[pl.slide]], pl, value):
                    notes.append(f"{f.name}, слайд {ri * len(keep) + remap[pl.slide] + 1}: {note}")
        notes_per_row.append(notes)

    if slide_sel is not None:
        _drop_all_but(prs, orig_ids, set(keep))
    sld_ids = prs.slides._sldIdLst
    by_part = {prs.part.related_slide(sld_id.get(qn("r:id"))).part: sld_id
               for sld_id in list(sld_ids)}
    ordered = [by_part[slide.part] for slides in per_row_slides for slide in slides]
    for sld_id in list(sld_ids):
        sld_ids.remove(sld_id)
    for sld_id in ordered:
        sld_ids.append(sld_id)
    prs.save(out_path)
    return notes_per_row
