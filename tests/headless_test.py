"""Headless acceptance tests for the core engines (cases 1-7).

Run:  .venv/bin/python tests/headless_test.py
"""

import json
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import pymupdf as fitz
except ImportError:
    import fitz
from pptx import Presentation

from app.core import excel_reader, mapping_io, pdf_engine, pptx_engine
from app.core.fonts import find_font_file
from app.core.generator import (
    EXISTS_OVERWRITE, EXISTS_SKIP, EXISTS_SUFFIX,
    MODE_COMBINED, MODE_SEPARATE, choose_filename_field, generate,
    planned_filenames,
)
from app.core.models import Field, PdfPlacement, PptxPlacement, TEMPLATE_PDF, TEMPLATE_PPTX
from make_test_data import NAMES, REGIONS
import openpyxl
from pptx.enum.shapes import MSO_SHAPE_TYPE

TD = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "testdata")
T_PPTX = os.path.join(TD, "template.pptx")
T_PDF = os.path.join(TD, "template.pdf")
XLSX = os.path.join(TD, "list.xlsx")
XLSX_BAD = os.path.join(TD, "list_bad.xlsx")
XLSX_LONG = os.path.join(TD, "list_long.xlsx")

PASS, FAIL = "PASS", "FAIL"
results = []
PERF: dict = {}


def check(name, cond, extra=""):
    results.append((name, bool(cond), extra))
    print(f"[{PASS if cond else FAIL}] {name} {extra}")


def pptx_fields():
    """Fields: 'ФИО' bound to both FIO shapes, 'Регион' bound to region shape."""
    prs = Presentation(T_PPTX)
    shapes = {s["shape_id"]: s for s in pptx_engine.list_text_shapes(prs)}
    fio_ids = [sid for sid, s in shapes.items() if "@@FIO@@" in s["text"]]
    reg_ids = [sid for sid, s in shapes.items() if "@@REGION@@" in s["text"]]
    assert len(fio_ids) == 2 and len(reg_ids) == 1, shapes
    def pl(sid):
        pat = shapes[sid]["text"]
        tok = pptx_engine.detect_token(pat)
        return PptxPlacement(0, sid, pattern=pat, token=tok,
                             mode="token" if tok != pat else "whole")

    return [
        Field("ФИО", [pl(sid) for sid in fio_ids]),
        Field("Регион", [pl(reg_ids[0])]),
    ]


def pdf_fields():
    font = find_font_file("Arial") or find_font_file("DejaVu Sans")
    doc = fitz.open(T_PDF)
    fio_rect = doc[0].search_for("@@FIO@@")[0]
    reg_rect = doc[0].search_for("@@REGION@@")[0]
    doc.close()

    def grow(r, dx=30, dy=14):
        return (r.x0 - dx, r.y0 - dy, r.x1 + dx, r.y1 + dy)

    return [
        Field("ФИО", [PdfPlacement(0, grow(fio_rect), font, 36, "#000000", "center")]),
        Field("Регион", [PdfPlacement(0, grow(reg_rect), font, 22, "#404040", "center")]),
    ]


def colmap(headers):
    return {"ФИО": headers.index("ФИО"), "Регион": headers.index("Регион")}


def pptx_texts(path):
    prs = Presentation(path)
    return [s.text_frame.text for s in prs.slides[0].shapes if s.has_text_frame]


def _sha(p):
    import hashlib
    return hashlib.sha256(open(p, "rb").read()).hexdigest()


def main():
    tpl_hash_pptx, tpl_hash_pdf = _sha(T_PPTX), _sha(T_PDF)
    headers, rows = excel_reader.read_xlsx(XLSX)
    cm = colmap(headers)
    tmp = tempfile.mkdtemp(prefix="diplom_test_")
    print("out dir:", tmp)

    # ---- case 1: pptx, 10 rows -> 10 files, design preserved -------------
    out1 = os.path.join(tmp, "c1")
    fields = pptx_fields()
    rep = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, rows, out1, MODE_SEPARATE,
                   cm["ФИО"], set())
    check("c1: 10 pptx files", len(rep["created"]) == 10, str(rep["created"][:3]))
    texts = pptx_texts(os.path.join(out1, rep["created"][0][1]))
    check("c1: FIO substituted", any("Иванов Иван Иванович" in t for t in texts))
    check("c1: no placeholder left", not any("@@" in t for t in texts))
    check("c1: title intact", any("ДИПЛОМ" in t for t in texts))
    # font preserved on region shape
    prs = Presentation(os.path.join(out1, rep["created"][0][1]))
    for shp in prs.slides[0].shapes:
        if shp.has_text_frame and "Москва" in shp.text_frame.text:
            run = shp.text_frame.paragraphs[0].runs[0]
            check("c1: region font size kept", run.font.size.pt == 24)
            check("c1: region font name kept", run.font.name == "Arial")

    # dedup: two identical ФИО -> _2 suffix
    names = sorted(fn for _r, fn in rep["created"])
    check("c2-ish: dup names suffixed",
          "Иванов Иван Иванович.pptx" in names and "Иванов Иван Иванович_2.pptx" in names,
          str(names))

    # ---- case 2: pdf, duplicates -----------------------------------------
    out2 = os.path.join(tmp, "c2")
    pfields = pdf_fields()
    rep2 = generate(T_PDF, TEMPLATE_PDF, pfields, cm, rows, out2, MODE_SEPARATE,
                    cm["ФИО"], set())
    check("c2: 10 pdf files", len(rep2["created"]) == 10)
    names2 = sorted(fn for _r, fn in rep2["created"])
    check("c2: dup names suffixed",
          "Иванов Иван Иванович.pdf" in names2 and "Иванов Иван Иванович_2.pdf" in names2)
    doc = fitz.open(os.path.join(out2, names2[0]))
    page_text = doc[0].get_text().replace("\xa0", " ").replace("\xad", "-")
    stem = os.path.splitext(names2[0])[0].replace("_2", "")
    check("c2: text inserted", stem in page_text, stem)
    check("c2: region inserted", any(r in page_text for r in REGIONS))
    check("c2: placeholder gone", "@@FIO@@" not in page_text)
    doc.close()

    # ---- case 3: one field bound to two placements ------------------------
    out3 = os.path.join(tmp, "c3")
    rep3 = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, [rows[1]], out3, MODE_SEPARATE,
                    cm["ФИО"], set())
    texts = pptx_texts(os.path.join(out3, rep3["created"][0][1]))
    hits = [t for t in texts if "Петрова Анна Сергеевна" in t]
    check("c3: field in two places", len(hits) == 2, str(texts))
    check("c3: static prefix kept", any("Выдан: Петрова" in t for t in texts))
    check("c3: region filled", any("Санкт-Петербург" in t for t in texts))

    # ---- case 4: empty cell -> validation warning -------------------------
    _h, bad_rows = excel_reader.read_xlsx(XLSX_BAD)
    probs = excel_reader.validate_rows(bad_rows, list(cm.values()), cm["ФИО"])
    check("c4: empty cell detected", any(r == 5 and "пуст" in p for r, p in probs), str(probs))

    # ---- case 5: long name shrinks ---------------------------------------
    _h2, long_rows = excel_reader.read_xlsx(XLSX_LONG)
    out5 = os.path.join(tmp, "c5")
    rep5 = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, long_rows, out5, MODE_SEPARATE,
                    cm["ФИО"], set())
    check("c5: long-name file created", len(rep5["created"]) == 4)
    check("c5: shrink noted", len(rep5["review"]) >= 1, str(rep5["review"]))
    prs5 = Presentation(os.path.join(out5, rep5["created"][0][1]))
    for shp in prs5.slides[0].shapes:
        if shp.has_text_frame and shp.text_frame.text.startswith("Абдур"):
            run = shp.text_frame.paragraphs[0].runs[0]
            check("c5: font shrunk below 40pt", run.font.size is not None and run.font.size.pt < 40)

    # ---- case 6: combined output ------------------------------------------
    out6 = os.path.join(tmp, "c6")
    rep6 = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, rows[:4], out6, MODE_COMBINED,
                    cm["ФИО"], set())
    comb = [fn for _r, fn in rep6["created"]]
    check("c6: combined pptx created", len(set(comb)) == 1)
    prs6 = Presentation(os.path.join(out6, comb[0]))
    check("c6: 4 slides", len(prs6.slides) == 4, str(len(prs6.slides)))
    slide_texts = [
        " | ".join(s.text_frame.text for s in sl.shapes if s.has_text_frame)
        for sl in prs6.slides
    ]
    check("c6: each slide has its row",
          "Иванов Иван Иванович" in slide_texts[0]
          and "Петрова Анна Сергеевна" in slide_texts[1]
          and "Сидоров Алексей Петрович" in slide_texts[2]
          and "Кузнецова Мария Игоревна" in slide_texts[3])
    check("c6: placeholders replaced", all("@@" not in t for t in slide_texts))

    rep6b = generate(T_PDF, TEMPLATE_PDF, pfields, cm, rows[:4], out6, MODE_COMBINED,
                     cm["ФИО"], set())
    combpdf = [fn for _r, fn in rep6b["created"]]
    d = fitz.open(os.path.join(out6, combpdf[0]))
    check("c6: combined pdf pages", d.page_count == 4)
    check("c6: pdf page2 has row2", "Петрова" in d[1].get_text().replace("\xa0", " "))
    d.close()

    # ---- case 7: mapping save/load ----------------------------------------
    path = mapping_io.save_mapping(T_PPTX, TEMPLATE_PPTX, fields,
                                   filename_source=("field", "ФИО"),
                                   filename_first_word=True)
    check("c7: mapping file written", os.path.isfile(path), path)
    ttype, loaded, extra = mapping_io.load_mapping(T_PPTX)
    check("c7: mapping roundtrip",
          ttype == "pptx" and len(loaded) == 2
          and loaded[0].name == "ФИО" and len(loaded[0].placements) == 2)
    check("c7: naming settings persisted",
          extra.get("filename_source") == ("field", "ФИО")
          and extra.get("filename_first_word") is True)
    # same for pdf
    mapping_io.save_mapping(T_PDF, TEMPLATE_PDF, pfields)
    tt2, ld2, _ex2 = mapping_io.load_mapping(T_PDF)
    check("c7: pdf mapping roundtrip",
          tt2 == "pdf" and isinstance(ld2[0].placements[0], PdfPlacement)
          and ld2[0].placements[0].font_size == 36)

    # ---- extra: per-row error isolation -----------------------------------
    class BadField(Field):
        pass

    bad_fields = [Field("ФИО", [PptxPlacement(0, 999999)]), fields[1]]
    outx = os.path.join(tmp, "cx")
    repx = generate(T_PPTX, TEMPLATE_PPTX, bad_fields, cm, rows[:3], outx, MODE_SEPARATE,
                    cm["ФИО"], set())
    check("extra: missing shape tolerated", len(repx["created"]) == 3)

    # cancellation
    calls = {"n": 0}

    def cancel_after_3():
        calls["n"] += 1
        return calls["n"] > 3

    outc = os.path.join(tmp, "cc")
    repc = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, rows, outc, MODE_SEPARATE,
                    cm["ФИО"], set(), cancel_cb=cancel_after_3)
    check("extra: cancel works", repc["cancelled"] and len(repc["created"]) <= 4)
    check("extra: report file", os.path.isfile(repc["report_path"]))

    # ===================== second-iteration cases ==========================

    # ---- template file must stay byte-identical ----------------------------
    check("tpl: pptx unchanged", _sha(T_PPTX) == tpl_hash_pptx)
    check("tpl: pdf unchanged", _sha(T_PDF) == tpl_hash_pdf)

    # ---- P0.2: overwrite policies ------------------------------------------
    out_e = os.path.join(tmp, "exists")
    rep_e1 = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, rows[:3], out_e,
                      MODE_SEPARATE, cm["ФИО"], set())
    check("exists: first run 3 files", len(rep_e1["created"]) == 3)
    rep_skip = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, rows[:3], out_e,
                        MODE_SEPARATE, cm["ФИО"], set(), exists_mode=EXISTS_SKIP)
    check("exists: skip mode creates 0", len(rep_skip["created"]) == 0)
    check("exists: skip mode reports 3", len(rep_skip["skipped"]) == 3)
    rep_suf = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, rows[:3], out_e,
                       MODE_SEPARATE, cm["ФИО"], set(), exists_mode=EXISTS_SUFFIX)
    names_suf = sorted(fn for _r, fn in rep_suf["created"])
    check("exists: unified _2 suffix", all("_2." in n for n in names_suf), str(names_suf))
    rep_suf2 = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, rows[:3], out_e,
                        MODE_SEPARATE, cm["ФИО"], set(), exists_mode=EXISTS_SUFFIX)
    names_suf2 = sorted(fn for _r, fn in rep_suf2["created"])
    check("exists: unified _3 suffix", all("_3." in n for n in names_suf2), str(names_suf2))
    rep_over = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, rows[:3], out_e,
                        MODE_SEPARATE, cm["ФИО"], set(), exists_mode=EXISTS_OVERWRITE)
    check("exists: overwrite creates 3", len(rep_over["created"]) == 3)
    planned = planned_filenames(T_PPTX, TEMPLATE_PPTX, rows[:3], MODE_SEPARATE,
                                cm["ФИО"], set())
    check("exists: planned names", len(planned) == 3
          and planned[0] == "Иванов Иван Иванович.pptx")

    # ---- P0.3: filename from chosen column ----------------------------------
    out_n = os.path.join(tmp, "names")
    rep_n = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, rows[:3], out_n,
                     MODE_SEPARATE, cm["Регион"], set())
    names_n = sorted(fn for _r, fn in rep_n["created"])
    check("namecol: files named by region",
          names_n == ["Москва.pptx", "Новосибирск.pptx", "Санкт-Петербург.pptx"],
          str(names_n))
    dup_rows = [["А А", "Казань"], ["Б Б", "Казань"], ["В В", "Казань"]]
    rep_nd = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, dup_rows,
                      os.path.join(tmp, "names2"), MODE_SEPARATE, 1, set())
    names_nd = sorted(fn for _r, fn in rep_nd["created"])
    check("namecol: dedup by chosen col",
          names_nd == ["Казань.pptx", "Казань_2.pptx", "Казань_3.pptx"], str(names_nd))

    # ---- P1.1: fragment mode (token is substring of the real block text) ----
    fio2_sid = next(pl.shape_id for pl in fields[0].placements
                    if pl.shape_id != fields[0].placements[0].shape_id)
    frag_fields = [
        Field("ФИО", [PptxPlacement(
            0, fio2_sid, pattern="Выдан: @@FIO@@", token="@@FIO@@",
            mode="fragment")]),
        fields[1],
    ]
    out_f = os.path.join(tmp, "frag")
    rep_f = generate(T_PPTX, TEMPLATE_PPTX, frag_fields, cm, [rows[0]], out_f,
                     MODE_SEPARATE, cm["ФИО"], set())
    texts_f = pptx_texts(os.path.join(out_f, rep_f["created"][0][1]))
    check("fragment: surrounding text kept",
          any("Выдан: Иванов Иван Иванович" in t for t in texts_f),
          str(texts_f))
    # back-compat: placement dict without "mode" behaves as whole
    reg_sid = fields[1].placements[0].shape_id
    pl_old = PptxPlacement.from_dict(
        {"type": "pptx_shape", "slide": 0, "shape_id": reg_sid,
         "pattern": "Выдан: @@FIO@@", "token": "@@FIO@@"})
    check("compat: missing mode -> whole", pl_old.mode == "whole")
    check("compat: whole ignores token",
          pptx_engine.new_text_for(pl_old, "X") == "X")

    # ---- P2.1: all-empty column hint ----------------------------------------
    wb_p = os.path.join(tmp, "empty_col.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["ФИО", "Регион"])
    for n in NAMES[:4]:
        ws.append([n, None])
    wb.save(wb_p)
    eh, er = excel_reader.read_xlsx(wb_p)
    probs_e = excel_reader.validate_rows(er, [0, 1], 0, eh)
    check("emptycol: column-level hint",
          any(r == 0 and "формул" in p for r, p in probs_e), str(probs_e))
    check("emptycol: no per-row noise",
          not any(r > 0 for r, p in probs_e))
    check("emptycol: helper", excel_reader.columns_all_empty(er, [0, 1]) == {1})

    # ---- P2.2: multiple sheets ----------------------------------------------
    wb_s = os.path.join(tmp, "sheets.xlsx")
    wb = openpyxl.Workbook()
    wb.active.title = "Первый"
    wb.active.append(["ФИО"]); wb.active.append(["Иванов А"])
    ws2 = wb.create_sheet("Второй")
    ws2.append(["ФИО"]); ws2.append(["Петров Б"])
    wb.save(wb_s)
    check("sheets: names listed",
          excel_reader.sheet_names(wb_s) == ["Первый", "Второй"])
    _h2, r_second = excel_reader.read_xlsx(wb_s, "Второй")
    check("sheets: second sheet read", r_second[0][0] == "Петров Б")

    # ---- P2.3: combined pptx keeps images ------------------------------------
    out_i = os.path.join(tmp, "img")
    rep_i = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, rows[:3], out_i,
                     MODE_COMBINED, cm["ФИО"], set())
    comb_i = rep_i["created"][0][1]
    prs_i = Presentation(os.path.join(out_i, comb_i))
    pics_per_slide = [
        sum(1 for s in sl.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE)
        for sl in prs_i.slides
    ]
    check("images: every slide keeps picture",
          len(prs_i.slides) == 3 and all(n >= 1 for n in pics_per_slide),
          str(pics_per_slide))
    # image bytes identical to template's
    tpl_prs = Presentation(T_PPTX)
    tpl_img = next(s for s in tpl_prs.slides[0].shapes
                   if s.shape_type == MSO_SHAPE_TYPE.PICTURE).image.blob
    out_img = next(s for s in prs_i.slides[1].shapes
                   if s.shape_type == MSO_SHAPE_TYPE.PICTURE).image.blob
    check("images: bytes identical", tpl_img == out_img)

    # ---- P1: first-word filename -------------------------------------------
    out_fw = os.path.join(tmp, "fw")
    rep_fw = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, rows[:5], out_fw,
                      MODE_SEPARATE, cm["ФИО"], set(), first_word=True)
    names_fw = sorted(fn for _r, fn in rep_fw["created"])
    check("firstword: surname-only names",
          names_fw == ["Иванов.pptx", "Иванов_2.pptx", "Кузнецова.pptx",
                       "Петрова.pptx", "Сидоров.pptx"], str(names_fw))

    # ---- P2: fragment formatting preserved ----------------------------------
    from pptx.util import Cm, Pt
    frag_pptx = os.path.join(tmp, "frag_tpl.pptx")
    prs_f = Presentation()
    prs_f.slide_width = Cm(29.7)
    prs_f.slide_height = Cm(21)
    sl = prs_f.slides.add_slide(prs_f.slide_layouts[6])
    tb = sl.shapes.add_textbox(Cm(2), Cm(9), Cm(26), Cm(3))
    p = tb.text_frame.paragraphs[0]
    r1 = p.add_run(); r1.text = "Награждается "; r1.font.size = Pt(18)
    r2 = p.add_run(); r2.text = "ФАМИЛИЯ"; r2.font.size = Pt(36); r2.font.bold = True
    r3 = p.add_run(); r3.text = " за победу"; r3.font.size = Pt(18)
    prs_f.save(frag_pptx)

    fshape_id = next(s.shape_id for s in sl.shapes if s.has_text_frame)
    frag_f = [Field("ФИО", [PptxPlacement(
        0, fshape_id,
        pattern="Награждается ФАМИЛИЯ за победу", token="ФАМИЛИЯ",
        mode="fragment")])]
    out_ff = os.path.join(tmp, "ff")
    rep_ff = generate(frag_pptx, TEMPLATE_PPTX, frag_f, {"ФИО": 0},
                      [["Иванов И.И."]], out_ff, MODE_SEPARATE, 0, set())
    check("fragfmt: file created", len(rep_ff["created"]) == 1)
    prs_out = Presentation(os.path.join(out_ff, rep_ff["created"][0][1]))
    fshape = next(s for s in prs_out.slides[0].shapes if s.has_text_frame)
    runs = [(r.text, r.font.size.pt if r.font.size else None, bool(r.font.bold))
            for para in fshape.text_frame.paragraphs for r in para.runs]
    check("fragfmt: composed text",
          "".join(r[0] for r in runs) == "Награждается Иванов И.И. за победу",
          str(runs))
    check("fragfmt: value keeps big bold format",
          any(r[0] == "Иванов И.И." and r[1] == 36 and r[2] for r in runs))
    check("fragfmt: surrounding keeps 18pt",
          any(r[0] == " за победу" and r[1] == 18 for r in runs)
          and any(r[0] == "Награждается " and r[1] == 18 for r in runs))

    # fragment spanning two differently-formatted runs -> flag in report
    frag2 = os.path.join(tmp, "frag2_tpl.pptx")
    prs_g = Presentation()
    prs_g.slide_width = Cm(29.7); prs_g.slide_height = Cm(21)
    sl2 = prs_g.slides.add_slide(prs_g.slide_layouts[6])
    tb2 = sl2.shapes.add_textbox(Cm(2), Cm(9), Cm(26), Cm(3))
    p2 = tb2.text_frame.paragraphs[0]
    g1 = p2.add_run(); g1.text = "Выдано: "; g1.font.size = Pt(18)
    g2 = p2.add_run(); g2.text = "ФА"; g2.font.size = Pt(36); g2.font.bold = True
    g3 = p2.add_run(); g3.text = "МИЛИЯ"; g3.font.size = Pt(20)
    g4 = p2.add_run(); g4.text = "."; g4.font.size = Pt(18)
    prs_g.save(frag2)
    gsid = next(s.shape_id for s in sl2.shapes if s.has_text_frame)
    frag_g = [Field("ФИО", [PptxPlacement(
        0, gsid, pattern="Выдано: ФАМИЛИЯ.", token="ФАМИЛИЯ", mode="fragment")])]
    rep_fg = generate(frag2, TEMPLATE_PPTX, frag_g, {"ФИО": 0}, [["Иванов И.И."]],
                      os.path.join(tmp, "fg"), MODE_SEPARATE, 0, set())
    check("fragfmt: multi-format flagged",
          any("первого run" in n for _r, notes in rep_fg["review"] for n in notes),
          str(rep_fg["review"]))
    prs_gout = Presentation(os.path.join(tmp, "fg", rep_fg["created"][0][1]))
    gshape = next(s for s in prs_gout.slides[0].shapes if s.has_text_frame)
    g_runs = [(r.text, r.font.size.pt if r.font.size else None)
              for para in gshape.text_frame.paragraphs for r in para.runs]
    check("fragfmt: multi-run value uses first format",
          any(r[0] == "Иванов И.И." and r[1] == 36 for r in g_runs), str(g_runs))
    check("fragfmt: tail kept", any(r[0] == "." and r[1] == 18 for r in g_runs))

    # ===================== fifth-iteration cases ==========================

    # ---- P0: PDF fill — coloured plate under a too-wide rect ---------------
    plate_pdf = os.path.join(tmp, "plate.pdf")
    font_path = find_font_file("Arial") or find_font_file("DejaVu Sans")
    d = fitz.open()
    pg = d.new_page(width=842, height=595)
    GOLD = (0.8, 0.65, 0.1)
    pg.draw_rect(fitz.Rect(200, 100, 650, 180), color=None, fill=GOLD)
    pg.insert_text(fitz.Point(240, 150), "FAMILIYA", fontsize=28,
                   color=(1, 1, 1))
    d.save(plate_pdf)
    d.close()
    d = fitz.open(plate_pdf)
    trect = d[0].search_for("FAMILIYA")[0]
    d.close()
    # rect spills past the plate (plate y 100..180) -> edge samples hit white
    wide = (trect.x0 - 30, 80.0, trect.x1 + 30, 195.0)

    def white_hole(pdf_path):
        dd = fitz.open(pdf_path)
        pix = dd[0].get_pixmap(dpi=72)
        whites = 0
        # strip inside the plate (y=105) and inside the redacted rect,
        # left of where the text sits
        for x in range(210, 260):
            r, g, b = pix.pixel(x, 105)[:3]
            if r > 240 and g > 240 and b > 240:
                whites += 1
        dd.close()
        return whites

    pl_auto = [Field("ФИО", [PdfPlacement(0, wide, font_path or "", 28,
                                          "#FFFFFF", "left")])]
    rep_pa = generate(plate_pdf, TEMPLATE_PDF, pl_auto, {"ФИО": 0},
                      [["IVANOV"]], os.path.join(tmp, "pa"),
                      MODE_SEPARATE, 0, set())
    check("fillrisk: auto warns about wide rect",
          any("края рамки светлее" in n
              for _r, notes in rep_pa["review"] for n in notes),
          str(rep_pa["review"]))
    auto_out = os.path.join(tmp, "pa", rep_pa["created"][0][1])
    check("fillrisk: auto leaves white hole (defect documented)",
          white_hole(auto_out) > 10, str(white_hole(auto_out)))

    pl_man = [Field("ФИО", [PdfPlacement(0, wide, font_path or "", 28,
                                         "#FFFFFF", "left",
                                         fill_color="#CCA61A")])]
    rep_pm = generate(plate_pdf, TEMPLATE_PDF, pl_man, {"ФИО": 0},
                      [["IVANOV"]], os.path.join(tmp, "pm"),
                      MODE_SEPARATE, 0, set())
    man_out = os.path.join(tmp, "pm", rep_pm["created"][0][1])
    check("fillrisk: manual colour leaves no white hole",
          white_hole(man_out) == 0, str(white_hole(man_out)))
    check("fillrisk: no warn when colour is manual",
          not any("края рамки" in n
                  for _r, notes in rep_pm["review"] for n in notes))
    # fill_color survives mapping roundtrip
    mapping_io.save_mapping(plate_pdf, TEMPLATE_PDF, pl_man)
    _tt, ld_plate, _e = mapping_io.load_mapping(plate_pdf)
    check("fillrisk: fill_color persisted",
          ld_plate[0].placements[0].fill_color == "#CCA61A")
    # old mapping without fill_color -> auto
    pl_nofill = PdfPlacement.from_dict(
        {"type": "pdf_rect", "page": 0, "rect": [0, 0, 10, 10]})
    check("fillrisk: missing fill_color -> auto", pl_nofill.fill_color == "")

    # ---- P2: detect_token only sees @@...@@ --------------------------------
    check("token: <...> not detected",
          pptx_engine.detect_token("Диплом № <номер> выдан") == "")
    check("token: [...] not detected",
          pptx_engine.detect_token("Секция [название] победитель") == "")
    check("token: _..._ not detected",
          pptx_engine.detect_token("100%_результат_") == "")
    check("token: @@...@@ detected",
          pptx_engine.detect_token("Выдан: @@FIO@@") == "@@FIO@@")

    # ---- P2: token twice in one block does not skew autofit -----------------
    # Deterministic width stub makes the check platform-independent. Literal
    # runs "№ " + " + " = 5 chars -> 90 pt at 36 pt; box 14 cm minus margins ->
    # ~382.4 pt. The second occurrence must be measured at the candidate size
    # like the first (90 + 2*5.5*s <= 382.4 -> s = 26), not at the literal
    # 36 pt (90 + 5.5*s + 198 -> s = 21, i.e. the old skewed behaviour).
    from unittest.mock import patch
    dup_pptx = os.path.join(tmp, "dup_tok.pptx")
    prs_d = Presentation()
    prs_d.slide_width = Cm(29.7)
    prs_d.slide_height = Cm(21)
    sld = prs_d.slides.add_slide(prs_d.slide_layouts[6])
    tbA = sld.shapes.add_textbox(Cm(2), Cm(5), Cm(14), Cm(3))
    ra = tbA.text_frame.paragraphs[0].add_run()
    ra.text = "№ @@FIO@@"
    ra.font.size = Pt(36)
    tbB = sld.shapes.add_textbox(Cm(2), Cm(10), Cm(14), Cm(3))
    rb = tbB.text_frame.paragraphs[0].add_run()
    rb.text = "№ @@FIO@@ + @@FIO@@"
    rb.font.size = Pt(36)
    tbC = sld.shapes.add_textbox(Cm(2), Cm(15), Cm(14), Cm(3))
    rc = tbC.text_frame.paragraphs[0].add_run()
    rc.text = "№ @@FIO@@ + "   # same literal text, token only once
    rc.font.size = Pt(36)
    prs_d.save(dup_pptx)

    sidB, sidC = tbB.shape_id, tbC.shape_id
    val = "Иванов И.И."
    fldB = [Field("ФИО", [PptxPlacement(0, sidB, pattern="№ @@FIO@@ + @@FIO@@",
                                      token="@@FIO@@", mode="fragment")])]
    fldC = [Field("ФИО", [PptxPlacement(0, sidC, pattern="№ @@FIO@@ + ",
                                      token="@@FIO@@", mode="fragment")])]
    with patch("app.core.fonts.measure_text_width",
               new=lambda text, fp, size: len(text) * size * 0.5):
        rep_db = generate(dup_pptx, TEMPLATE_PPTX, fldB, {"ФИО": 0}, [[val]],
                          os.path.join(tmp, "db"), MODE_SEPARATE, 0, set())
        rep_dc = generate(dup_pptx, TEMPLATE_PPTX, fldC, {"ФИО": 0}, [[val]],
                          os.path.join(tmp, "dc"), MODE_SEPARATE, 0, set())

    def value_sizes(path):
        prs = Presentation(path)
        out = []
        for s in prs.slides[0].shapes:
            if s.has_text_frame:
                for pa in s.text_frame.paragraphs:
                    for r in pa.runs:
                        if r.text == val:
                            out.append(r.font.size.pt)
        return out

    szB = value_sizes(os.path.join(tmp, "db", rep_db["created"][0][1]))
    szC = value_sizes(os.path.join(tmp, "dc", rep_dc["created"][0][1]))
    check("duptok: both occurrences replaced", len(szB) == 2, str(szB))
    check("duptok: repeated token does not skew autofit",
          szB == [26.0, 26.0] and szC == [36.0],
          f"double={szB} single={szC}")

    from app.core.generator import MAX_PATH_LEN, plan_batch_names, row_values
    from pptx.dml.color import RGBColor
    from unittest.mock import patch

    names = plan_batch_names(["Иванов", "Иванов", "Иванов_2", "A", "a", "CON", "con.",
                              "a/b", "a:b"], tmp, ".pptx", EXISTS_OVERWRITE)
    actual = [n for n, _ in names]
    check("iter6: unique casefold names and reserved devices",
          len(actual) == len(set(n.casefold() for n in actual)) == 9
          and actual[0:3] == ["Иванов.pptx", "Иванов_2.pptx", "Иванов_2_2.pptx"]
          and all(os.path.splitext(n)[0].casefold() not in ("con", "nul", "prn")
                  for n in actual), str(actual))
    collision_rep = generate(T_PPTX, TEMPLATE_PPTX, fields, cm,
                             [["Иванов", "X"], ["Иванов", "Y"], ["Иванов_2", "Z"]],
                             os.path.join(tmp, "batch_collision"), MODE_SEPARATE, 0)
    check("iter6: batch collision saves three real files", len(collision_rep["created"]) == 3
          and len({n.casefold() for _, n in collision_rep["created"]}) == 3
          and all(len(Presentation(os.path.join(collision_rep["out_dir"], n)).slides) == 1
                  for _, n in collision_rep["created"]))
    deep = os.path.join(tmp, "д" * max(1, MAX_PATH_LEN - len(tmp) - 85))
    long_names = [n for n, _ in plan_batch_names(["Ж" * 80] * 3, deep, ".pptx", EXISTS_OVERWRITE)]
    check("iter6: suffixes survive truncation", len(set(long_names)) == 3
          and long_names[1].endswith("_2.pptx") and long_names[2].endswith("_3.pptx")
          and all(len(os.path.join(deep, n)) <= MAX_PATH_LEN for n in long_names), str(long_names))
    impossible = plan_batch_names(["file"], "/" + "a" * MAX_PATH_LEN, ".pptx", EXISTS_SUFFIX)
    check("iter6: impossible path diagnosed", impossible[0][0] is None and bool(impossible[0][1]))
    tight_dir = "/" + "a" * (MAX_PATH_LEN - 1 - len(".pptx") - 3 - 1)
    tight_name = plan_batch_names(["CON_"], tight_dir, ".pptx", EXISTS_OVERWRITE)[0][0]
    check("iter6: truncation cannot revive Windows device name",
          tight_name is not None and os.path.splitext(tight_name)[0].casefold() != "con", str(tight_name))

    decorated = os.path.join(tmp, "decorated.pptx")
    deck = Presentation()
    deck.slide_width = Cm(28)
    deck.slide_height = Cm(20)
    d0 = deck.slides.add_slide(deck.slide_layouts[6])
    d0.background.fill.solid()
    d0.background.fill.fore_color.rgb = RGBColor(12, 34, 56)
    n0 = d0.shapes.add_textbox(Cm(2), Cm(4), Cm(10), Cm(2))
    n0.text = "@@FIO@@"
    d1 = deck.slides.add_slide(deck.slide_layouts[6])
    d1.background.fill.solid()
    d1.background.fill.fore_color.rgb = RGBColor(48, 65, 82)
    n1 = d1.shapes.add_textbox(Cm(2), Cm(4), Cm(10), Cm(2))
    n1.text = "@@FIO@@"
    deck.save(decorated)
    both = [Field("ФИО", [PptxPlacement(i, b.shape_id, token="@@FIO@@", mode="token")
                          for i, b in enumerate([n0, n1])])]
    dr = generate(decorated, TEMPLATE_PPTX, both, {"ФИО": 0}, [["Первый"], ["Второй"]],
                  os.path.join(tmp, "deck"), MODE_COMBINED, 0)
    result = Presentation(os.path.join(tmp, "deck", dr["created"][0][1]))
    names_order = [next(s.text_frame.text for s in sl.shapes if s.has_text_frame)
                   for sl in result.slides]
    colors = [str(sl.background.fill.fore_color.rgb) for sl in result.slides]
    check("iter6: deck grouped by row", names_order == ["Первый", "Первый", "Второй", "Второй"], str(names_order))
    check("iter6: solid slide-level backgrounds copied",
          colors == ["0C2238", "304152", "0C2238", "304152"], str(colors))
    one = generate(decorated, TEMPLATE_PPTX, both, {"ФИО": 0}, [["Первый"], ["Второй"]],
                   os.path.join(tmp, "deck_one"), MODE_COMBINED, 0, slide_sel=[1])
    selected = Presentation(os.path.join(tmp, "deck_one", one["created"][0][1]))
    check("iter6: selected slide only", len(selected.slides) == 2 and
          [str(sl.background.fill.fore_color.rgb) for sl in selected.slides] == ["304152"] * 2)
    from pptx.oxml.xmlchemy import OxmlElement
    from pptx.oxml.ns import qn
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    styled = Presentation(decorated)
    p_bg = styled.slides[0].element.find(qn("p:cSld")).find(qn("p:bg"))
    p_bg.remove(p_bg[0])
    grad = OxmlElement("p:bgPr")
    grad_fill = OxmlElement("a:gradFill")
    grad_stops = OxmlElement("a:gsLst")
    for pos, hexcolor in (("0", "0C2238"), ("100000", "FFFFFF")):
        stop = OxmlElement("a:gs")
        stop.set("pos", pos)
        color = OxmlElement("a:srgbClr")
        color.set("val", hexcolor)
        stop.append(color)
        grad_stops.append(stop)
    grad_fill.append(grad_stops)
    grad.append(grad_fill)
    p_bg.append(grad)
    pic = next(s for s in Presentation(T_PPTX).slides[0].shapes
               if s.shape_type == MSO_SHAPE_TYPE.PICTURE)
    pic_part = pic.part.rels[pic._element.blip_rId].target_part
    image_slide = styled.slides[1]
    image_bg = image_slide.element.find(qn("p:cSld")).find(qn("p:bg"))
    image_bg.remove(image_bg[0])
    image_pr = OxmlElement("p:bgPr")
    blip_fill = OxmlElement("a:blipFill")
    blip = OxmlElement("a:blip")
    blip.set(qn("r:embed"), image_slide.part.rels.get_or_add(RT.IMAGE, pic_part))
    blip_fill.append(blip)
    image_pr.append(blip_fill)
    image_bg.append(image_pr)
    styled_file = os.path.join(tmp, "backgrounds.pptx")
    styled.save(styled_file)
    styled_rep = generate(styled_file, TEMPLATE_PPTX, both, {"ФИО": 0},
                          [["Первый"], ["Второй"]], os.path.join(tmp, "backgrounds"), MODE_COMBINED, 0)
    reopened = Presentation(os.path.join(tmp, "backgrounds", styled_rep["created"][0][1]))
    bg_first = reopened.slides[0].element.find(qn("p:cSld")).find(qn("p:bg"))
    bg_copy = reopened.slides[2].element.find(qn("p:cSld")).find(qn("p:bg"))
    check("iter6: gradient background survives duplicate", bg_first.xml == bg_copy.xml)
    bg_pic = reopened.slides[3].element.find(qn("p:cSld")).find(qn("p:bg"))
    copied_rid = bg_pic.find(".//" + qn("a:blip")).get(qn("r:embed"))
    check("iter6: picture background relation survives duplicate",
          reopened.slides[3].part.rels[copied_rid].target_part.blob == pic_part.blob)

    cancelled = generate(decorated, TEMPLATE_PPTX, both, {"ФИО": 0}, [["x"], ["y"]],
                         os.path.join(tmp, "cancel_pptx"), MODE_COMBINED, 0,
                         cancel_cb=lambda: True)
    check("iter6: combined pptx cancel has no output", cancelled["cancelled"] and
          not cancelled["created"] and not any(n.endswith(".pptx") for n in os.listdir(cancelled["out_dir"])))
    calls = [0]
    def abort_second():
        calls[0] += 1
        return calls[0] >= 2
    cancelled_pdf = generate(T_PDF, TEMPLATE_PDF, pfields, cm, rows[:2],
                             os.path.join(tmp, "cancel_pdf"), MODE_COMBINED, 0,
                             cancel_cb=abort_second)
    check("iter6: combined pdf cancel has no output", cancelled_pdf["cancelled"] and
          not cancelled_pdf["created"] and not any(n.endswith(".pdf") for n in os.listdir(cancelled_pdf["out_dir"])))
    atom_dir = os.path.join(tmp, "atomic")
    initial = generate(decorated, TEMPLATE_PPTX, both, {"ФИО": 0}, [["Исходный"]],
                       atom_dir, MODE_COMBINED, 0)
    target = os.path.join(atom_dir, initial["created"][0][1])
    before_hash = _sha(target)
    with patch("app.core.generator.os.replace", side_effect=OSError("no write")):
        failed = generate(decorated, TEMPLATE_PPTX, both, {"ФИО": 0}, [["Новый"]],
                          atom_dir, MODE_COMBINED, 0)
    check("iter6: failed save keeps existing output", not failed["created"] and
          _sha(target) == before_hash and any("no write" in r for _, r in failed["skipped"]))
    pdf_atom_dir = os.path.join(tmp, "atomic_pdf")
    pdf_initial = generate(T_PDF, TEMPLATE_PDF, pfields, cm, rows[:1],
                           pdf_atom_dir, MODE_COMBINED, 0)
    pdf_target = os.path.join(pdf_atom_dir, pdf_initial["created"][0][1])
    pdf_hash = _sha(pdf_target)
    with patch("app.core.generator.os.replace", side_effect=OSError("no write")):
        pdf_failed = generate(T_PDF, TEMPLATE_PDF, pfields, cm, rows[1:2],
                              pdf_atom_dir, MODE_COMBINED, 0)
    check("iter6: failed combined pdf save keeps existing output",
          not pdf_failed["created"] and _sha(pdf_target) == pdf_hash
          and any("no write" in reason for _, reason in pdf_failed["skipped"]))

    tree = Presentation()
    sl = tree.slides.add_slide(tree.slide_layouts[6])
    group = sl.shapes.add_group_shape()
    inner = group.shapes.add_group_shape()
    box = inner.shapes.add_textbox(Cm(2), Cm(2), Cm(8), Cm(2))
    box.text = "Группа @@FIO@@"
    table = sl.shapes.add_table(2, 2, Cm(2), Cm(7), Cm(12), Cm(4)).table
    table.cell(1, 0).text = "@@FIO@@"
    tree_file = os.path.join(tmp, "tree.pptx")
    tree.save(tree_file)
    tb = pptx_engine.list_text_shapes(Presentation(tree_file))
    grp_block = next(b for b in tb if "Группа" in b["text"])
    cell_block = next(b for b in tb if b["cell"] == (1, 0))
    tree_fields = [Field("ФИО", [PptxPlacement(0, grp_block["shape_id"], token="@@FIO@@", mode="token"),
                                  PptxPlacement(0, cell_block["shape_id"], token="@@FIO@@",
                                                mode="token", cell=cell_block["cell"])])]
    tree_rep = generate(tree_file, TEMPLATE_PPTX, tree_fields, {"ФИО": 0}, [["Участник"]],
                        os.path.join(tmp, "tree_out"), MODE_SEPARATE, 0)
    new_tree = pptx_engine.list_text_shapes(Presentation(os.path.join(tmp, "tree_out", tree_rep["created"][0][1])))
    check("iter6: nested groups and table cells", any(b["text"] == "Группа Участник" for b in new_tree)
          and any(b["cell"] == (1, 0) and b["text"] == "Участник" for b in new_tree))
    check("iter6: table cell mapping roundtrip", PptxPlacement.from_dict(
          tree_fields[0].placements[1].to_dict()).cell == (1, 0))
    mismatch = pptx_engine.apply_field_to_slide(sl, PptxPlacement(0, box.shape_id,
                                                token="НЕСОВПАДЕНИЕ", mode="fragment"), "X")
    check("iter6: mismatch preserves text", box.text == "Группа @@FIO@@" and mismatch)
    keep_field = Field("ФИО", [PptxPlacement(0, box.shape_id)], source="keep")
    check("iter6: optional keep does not clear template", "ФИО" not in row_values([keep_field], {}, []))
    req_field = Field("ФИО", [PptxPlacement(0, n0.shape_id)], required=True)
    invalid = generate(decorated, TEMPLATE_PPTX, [req_field], {}, [["Тест"]],
                       os.path.join(tmp, "invalid"), MODE_COMBINED, 0)
    check("iter6: missing required column blocks generation", not invalid["created"]
          and any(r == 0 for r, _ in invalid["skipped"]))
    empty_case = generate(decorated, TEMPLATE_PPTX, [req_field], {"ФИО": 0},
                          [["Готов"], [""], ["Далее"]], os.path.join(tmp, "папка результата"),
                          MODE_COMBINED, 0)
    empty_slides = Presentation(os.path.join(empty_case["out_dir"], empty_case["created"][0][1]))
    check("iter6: blank row omitted with source row preserved", len(empty_slides.slides) == 4
          and any(r == 3 and "пуст" in why for r, why in empty_case["skipped"]))
    fifty = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, rows[:5] * 10,
                     os.path.join(tmp, "50 дипломов"), MODE_COMBINED, 0)
    fifty_file = os.path.join(fifty["out_dir"], fifty["created"][0][1])
    check("iter6: 50 rows combined pptx reopens", len(fifty["created"]) == 50
          and len(Presentation(fifty_file).slides) == 50)

    guarded = Field("ФИО", [PptxPlacement(0, n0.shape_id, token="@@FIO@@", mode="token")],
                    required=True)
    guarded_dir = os.path.join(tmp, "guarded")
    valid_guard = generate(decorated, TEMPLATE_PPTX,
                           [Field("ФИО", [PptxPlacement(1, n1.shape_id, token="@@FIO@@",
                                                      mode="token")], required=True)],
                           {"ФИО": 0}, [["Участник А"], ["Участник Б"]],
                           guarded_dir, MODE_COMBINED, 0, slide_sel=[1])
    guarded_file = os.path.join(guarded_dir, valid_guard["created"][0][1])
    original_guard_hash = _sha(guarded_file)
    blocked_combined = generate(decorated, TEMPLATE_PPTX, [guarded], {"ФИО": 0},
                                [["Участник А"], ["Участник Б"]], guarded_dir,
                                MODE_COMBINED, 0, slide_sel=[1])
    blocked_separate = generate(decorated, TEMPLATE_PPTX, [guarded], {"ФИО": 0},
                                [["Участник А"], ["Участник Б"]],
                                os.path.join(tmp, "guarded_separate"), MODE_SEPARATE,
                                0, slide_sel=[1])
    check("review6: excluded required placement blocks both modes",
          not blocked_combined["created"] and not blocked_separate["created"]
          and any("ФИО" in reason for _, reason in blocked_combined["skipped"])
          and _sha(guarded_file) == original_guard_hash)
    check("review6: valid selected slide uses each row",
          [next(s.text_frame.text for s in sl.shapes if s.has_text_frame)
           for sl in Presentation(guarded_file).slides] == ["Участник А", "Участник Б"])
    wrong_token = Field("ФИО", [PptxPlacement(1, n1.shape_id, token="@@MISSING@@",
                                             mode="token")], required=True)
    wrong_shape = Field("ФИО", [PptxPlacement(1, 999999)], required=True)
    for name, bad_field in (("token", wrong_token), ("shape", wrong_shape)):
        blocked = generate(decorated, TEMPLATE_PPTX, [bad_field], {"ФИО": 0},
                           [["Участник А"]], guarded_dir, MODE_COMBINED, 0,
                           slide_sel=[1])
        check(f"review6: invalid required {name} preserves existing output",
              not blocked["created"] and _sha(guarded_file) == original_guard_hash
              and any("ФИО" in reason for _, reason in blocked["skipped"]))
    optional_excluded = Field("Номинация", [PptxPlacement(0, n0.shape_id)])
    optional_rep = generate(decorated, TEMPLATE_PPTX,
                            [Field("ФИО", [PptxPlacement(1, n1.shape_id, token="@@FIO@@",
                                                       mode="token")], required=True),
                             optional_excluded], {"ФИО": 0}, [["Участник А"]],
                            os.path.join(tmp, "optional_excluded"), MODE_COMBINED, 0,
                            slide_sel=[1])
    check("review6: optional excluded placement does not block", len(optional_rep["created"]) == 1)

    inherited = Presentation()
    inherited_slide = inherited.slides.add_slide(inherited.slide_layouts[6])
    inherited_box = inherited_slide.shapes.add_textbox(Pt(5), Pt(5), Pt(300), Pt(100))
    inherited_box.text_frame.word_wrap = False
    inherited_box.text_frame.margin_left = 0
    inherited_box.text_frame.margin_right = 0
    inherited_box.text_frame.margin_top = 0
    inherited_box.text_frame.margin_bottom = 0
    ip = inherited_box.text_frame.paragraphs[0]
    ip.font.name = "Arial"
    ip.font.size = Pt(44)
    ip.add_run().text = "@@FIO@@"
    inherited_note = pptx_engine.apply_field_to_slide(
        inherited_slide, PptxPlacement(0, inherited_box.shape_id,
                                       token="@@FIO@@", mode="fragment"), "Тестовый Участник А")
    inherited_run = next(r for r in ip.runs if "Тестовый" in r.text)
    check("review6: paragraph font inherited for fragment",
          inherited_run.font.size is not None and inherited_run.font.size.pt < 44
          and any("уменьшен" in note or "не поместился" in note for note in inherited_note),
          str(inherited_note))

    inherited_master = Presentation()
    master_style = inherited_master.slide_masters[0].element.find(qn("p:txStyles"))
    # text boxes are styled by p:otherStyle (bodyStyle is for placeholders)
    level = master_style.find(qn("p:otherStyle")).find(qn("a:lvl1pPr"))
    master_rpr = level.find(qn("a:defRPr"))
    master_rpr.set("sz", "4400")
    master_rpr.find(qn("a:latin")).set("typeface", "Arial")
    master_slide = inherited_master.slides.add_slide(inherited_master.slide_layouts[6])
    master_box = master_slide.shapes.add_textbox(Pt(5), Pt(5), Pt(300), Pt(100))
    master_box.text_frame.word_wrap = False
    master_box.text_frame.margin_left = 0
    master_box.text_frame.margin_right = 0
    master_box.text_frame.paragraphs[0].add_run().text = "@@FIO@@"
    master_note = pptx_engine.apply_field_to_slide(
        master_slide, PptxPlacement(0, master_box.shape_id,
                                    token="@@FIO@@", mode="fragment"), "Тестовый Участник А")
    master_run = next(r for r in master_box.text_frame.paragraphs[0].runs if "Тестовый" in r.text)
    check("review6: master font inherited for fragment",
          master_run.font.size is not None and master_run.font.size.pt < 44
          and any("уменьшен" in note or "не поместился" in note for note in master_note),
          str(master_note))
    inherited_layout = Presentation()
    layout_shape = next(x for x in inherited_layout.slide_layouts[1].placeholders
                        if x.placeholder_format.idx == 1)
    layout_shape.text_frame.paragraphs[0].font.size = Pt(44)
    layout_shape.text_frame.paragraphs[0].font.name = "Arial"
    layout_slide = inherited_layout.slides.add_slide(inherited_layout.slide_layouts[1])
    layout_box = next(x for x in layout_slide.placeholders if x.placeholder_format.idx == 1)
    layout_box.width = Pt(300)
    layout_box.height = Pt(100)
    layout_box.text_frame.word_wrap = False
    layout_box.text_frame.margin_left = 0
    layout_box.text_frame.margin_right = 0
    layout_box.text_frame.paragraphs[0].add_run().text = "@@FIO@@"
    layout_file = os.path.join(tmp, "layout_font.pptx")
    inherited_layout.save(layout_file)
    loaded_layout = Presentation(layout_file).slides[0]
    loaded_box = next(x for x in loaded_layout.placeholders if x.placeholder_format.idx == 1)
    layout_note = pptx_engine.apply_field_to_slide(
        loaded_layout, PptxPlacement(0, loaded_box.shape_id,
                                     token="@@FIO@@", mode="fragment"), "Тестовый Участник А")
    layout_run = next(r for r in loaded_box.text_frame.paragraphs[0].runs if "Тестовый" in r.text)
    check("review6: layout font inherited for fragment",
          layout_run.font.size is not None and layout_run.font.size.pt < 44
          and any("уменьшен" in note or "не поместился" in note for note in layout_note),
          str(layout_note))

    from lxml import etree
    styled_slide = Presentation(decorated)
    styled_slide.slides[0].element.set("showMasterSp", "0")
    styled_slide.slides[0].element.set("showMasterPhAnim", "0")
    styled_slide.slides[0].element.set("show", "0")
    source_color = styled_slide.slides[0].element.find(qn("p:clrMapOvr"))
    source_color.remove(source_color[0])
    override = OxmlElement("a:overrideClrMapping")
    for key, value in {"accent1": "accent2", "accent2": "accent1", "accent3": "accent3",
                       "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
                       "bg1": "lt1", "bg2": "lt2", "tx1": "dk1", "tx2": "dk2",
                       "hlink": "hlink", "folHlink": "folHlink"}.items():
        override.set(key, value)
    source_color.append(override)
    styled_file = os.path.join(tmp, "master_hidden.pptx")
    styled_slide.save(styled_file)
    hidden_rep = generate(styled_file, TEMPLATE_PPTX, both, {"ФИО": 0},
                          [["Участник А"], ["Участник Б"]],
                          os.path.join(tmp, "hidden_out"), MODE_COMBINED, 0)
    hidden_result = Presentation(os.path.join(tmp, "hidden_out", hidden_rep["created"][0][1]))
    hidden_source, hidden_copy = hidden_result.slides[0], hidden_result.slides[2]
    check("review6: slide visibility flags and color mapping copied",
          all(hidden_source.element.get(attr) == hidden_copy.element.get(attr) == "0"
              for attr in ("showMasterSp", "showMasterPhAnim", "show"))
          and etree.tostring(hidden_copy.element.find(qn("p:clrMapOvr"))) ==
              etree.tostring(hidden_source.element.find(qn("p:clrMapOvr"))))

    os.environ["QT_QPA_PLATFORM"] = "offscreen"
    from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox
    from app.gui.state import AppState
    from app.gui.step_excel import StepExcel
    gui_app = QApplication.instance() or QApplication([])
    sheets_path = os.path.join(tmp, "two_sheets.xlsx")
    sheets_book = openpyxl.Workbook()
    filled_sheet = sheets_book.active
    filled_sheet.title = "Filled"
    filled_sheet.append(["ФИО"])
    filled_sheet.append(["Участник А"])
    empty_sheet = sheets_book.create_sheet("Empty")
    empty_sheet.append(["ФИО"])
    sheets_book.save(sheets_path)
    sheets_book.close()
    gui_state = AppState(fields=[Field("ФИО", required=True)])
    navigations = []
    step = StepExcel(gui_state, lambda: navigations.append(True), lambda: None)
    with patch.object(QFileDialog, "getOpenFileName", return_value=(sheets_path, "")), \
         patch.object(QMessageBox, "warning", return_value=QMessageBox.Ok):
        step._open_excel()
        step.cmb_sheet.setCurrentIndex(1)
        step._next()
        check("review6: empty worksheet clears GUI data and blocks next",
              not gui_state.rows and gui_state.sheet_name == "Empty"
              and step.table.rowCount() == 0 and not navigations)
        step.cmb_sheet.setCurrentIndex(0)
        check("review6: returning to filled worksheet restores mapping",
              gui_state.rows == [["Участник А"]] and gui_state.colmap == {"ФИО": 0})
    with patch("app.gui.step_excel.excel_reader.read_xlsx", side_effect=OSError("read failed")), \
         patch.object(QMessageBox, "critical", return_value=QMessageBox.Ok):
        step.cmb_sheet.setCurrentIndex(1)
    check("review6: failed worksheet read clears active participants",
          gui_state.sheet_name == "Empty" and not gui_state.rows
          and not step.btn_next.isEnabled())
    step.close()
    from app.gui.step_template import StepTemplate
    from app.gui.step_generate import StepGenerate
    invalid_state = AppState(template_path=decorated, template_type=TEMPLATE_PPTX,
                             fields=[guarded], slide_sel=[1], rows=[["Участник А"]],
                             colmap={"ФИО": 0}, out_dir=tmp)
    blocked_navigation = []
    template_step = StepTemplate(invalid_state, lambda: blocked_navigation.append(True))
    template_step._prs = Presentation(decorated)
    generator_step = StepGenerate(invalid_state, lambda: None, lambda: None)
    with patch.object(QMessageBox, "warning", return_value=QMessageBox.Ok):
        template_step._next()
        generator_step._start()
    check("review6: GUI blocks invalid required slide binding",
          not blocked_navigation and generator_step._worker is None)
    template_step.close()
    generator_step.close()

    # ---- review7: empty sheet with a different schema keeps the mapping ---
    swap_path = os.path.join(tmp, "swap_sheets.xlsx")
    swap_book = openpyxl.Workbook()
    ws_filled = swap_book.active
    ws_filled.title = "Filled"
    ws_filled.append(["Номинация", "ФИО", "Регион"])
    ws_filled.append(["Тестовая номинация", "Участник А", "Тестовый регион"])
    ws_empty = swap_book.create_sheet("Empty")
    ws_empty.append(["ФИО", "Номинация"])   # same titles, reversed order, no rows
    swap_book.create_sheet("Broken").append(["ФИО"])
    swap_book.save(swap_path)
    swap_book.close()
    swap_state = AppState(fields=[Field("ФИО", required=True), Field("Номинация"),
                                  Field("Регион"), Field("Город")])
    step2 = StepExcel(swap_state, lambda: None, lambda: None)
    expected_map = {"ФИО": 1, "Номинация": 0, "Регион": 2, "Город": 2}
    with patch.object(QFileDialog, "getOpenFileName", return_value=(swap_path, "")), \
         patch.object(QMessageBox, "warning", return_value=QMessageBox.Ok):
        step2._open_excel()
        # manual mapping of a non-standard field name to the "Регион" column
        step2._combos["Город"].setCurrentIndex(4)
        check("review7: initial map set", swap_state.colmap == expected_map,
              str(swap_state.colmap))
        step2.cmb_sheet.setCurrentIndex(1)   # Empty with reversed headers
        step2.cmb_sheet.setCurrentIndex(0)   # back to Filled
        check("review7: swapped empty sheet keeps column mapping",
              swap_state.colmap == expected_map, str(swap_state.colmap))
    real_read = excel_reader.read_xlsx
    def flaky_read(path, sheet=None):
        if sheet == "Broken":
            raise OSError("read failed")
        return real_read(path, sheet)
    with patch("app.gui.step_excel.excel_reader.read_xlsx", side_effect=flaky_read), \
         patch.object(QMessageBox, "critical", return_value=QMessageBox.Ok):
        step2.cmb_sheet.setCurrentIndex(2)   # Broken: read error
    with patch.object(QMessageBox, "warning", return_value=QMessageBox.Ok):
        step2.cmb_sheet.setCurrentIndex(0)
    check("review7: read failure keeps column mapping",
          swap_state.colmap == expected_map, str(swap_state.colmap))
    # end-to-end: generated PPTX must put each value in its own block
    mini = Presentation()
    mini_slide = mini.slides.add_slide(mini.slide_layouts[6])
    m_fio = mini_slide.shapes.add_textbox(Cm(1), Cm(2), Cm(10), Cm(2))
    m_fio.text_frame.text = "@@FIO@@"
    m_nom = mini_slide.shapes.add_textbox(Cm(1), Cm(6), Cm(14), Cm(2))
    m_nom.text_frame.text = "Номинация: «Тестовая номинация»"
    mini_file = os.path.join(tmp, "mini.pptx")
    mini.save(mini_file)
    mini_fields = [
        Field("ФИО", [PptxPlacement(0, m_fio.shape_id, "@@FIO@@", "@@FIO@@",
                                    "token")], required=True),
        Field("Номинация", [PptxPlacement(0, m_nom.shape_id, m_nom.text_frame.text,
                                          "Тестовая номинация", "fragment")],
              required=True),
    ]
    mini_rep = generate(mini_file, TEMPLATE_PPTX, mini_fields,
                        {k: v for k, v in swap_state.colmap.items() if k != "Город"},
                        swap_state.rows, os.path.join(tmp, "mini_out"),
                        MODE_SEPARATE, swap_state.colmap["ФИО"])
    mini_texts = [s.text_frame.text for s in Presentation(
        os.path.join(tmp, "mini_out", mini_rep["created"][0][1])).slides[0].shapes
        if s.has_text_frame]
    check("review7: generated file puts values in the right blocks",
          "Участник А" in mini_texts
          and "Номинация: «Тестовая номинация»" in mini_texts, str(mini_texts))
    step2.close()
    gui_app.quit()

    # ---- review7: group scale does not fool autofit -----------------------
    def scaled_group_deck(scale_x, scale_y, nested=False):
        deck = Presentation()
        s = deck.slides.add_slide(deck.slide_layouts[6])
        grp = s.shapes.add_group_shape()
        parent = grp
        if nested:
            inner = grp.shapes.add_group_shape()
            parent = inner
        box = parent.shapes.add_textbox(Pt(5), Pt(5), Pt(300), Pt(100))
        tf = box.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.word_wrap = False
        par = tf.paragraphs[0]
        par.font.name = "Arial"
        par.font.size = Pt(44)
        par.add_run().text = "@@FIO@@"
        if nested:
            inner.width = int(inner.width * 2)
            inner.height = int(inner.height * 2)
        grp.width = int(grp.width * scale_x)
        grp.height = int(grp.height * scale_y)
        return s, box

    for tag, sx, sy, nested in (("x2", 2, 2, False), ("x0.5", 0.5, 0.5, False),
                                ("nested", 2, 2, True), ("aniso", 2, 1, False)):
        gs, gbox = scaled_group_deck(sx, sy, nested)
        gnote = pptx_engine.apply_field_to_slide(
            gs, PptxPlacement(0, gbox.shape_id, token="@@FIO@@", mode="fragment"),
            "Тестовый Участник А")
        grun = next(r for r in gbox.text_frame.paragraphs[0].runs if "Тестовый" in r.text)
        check(f"review7: group {tag} autofit uses local coords",
              grun.font.size is not None and grun.font.size.pt < 44
              and any("уменьшен" in n or "не поместился" in n for n in gnote),
              f"notes={gnote} size={grun.font.size}")

    # ---- review7: shape text resolves p:otherStyle ------------------------
    from pptx.enum.shapes import MSO_SHAPE
    from app.core.fonts import measure_text_width
    style_deck = Presentation()
    tx = style_deck.slide_masters[0].element.find(qn("p:txStyles"))
    body_rpr = tx.find(qn("p:bodyStyle")).find(qn("a:lvl1pPr")).find(qn("a:defRPr"))
    body_rpr.set("sz", "3200")
    body_rpr.set("b", "1")
    body_rpr.find(qn("a:latin")).set("typeface", "Times New Roman")
    other_rpr = tx.find(qn("p:otherStyle")).find(qn("a:lvl1pPr")).find(qn("a:defRPr"))
    other_rpr.set("sz", "4400")
    other_rpr.find(qn("a:latin")).set("typeface", "Arial")
    style_slide = style_deck.slides.add_slide(style_deck.slide_layouts[6])
    rect = style_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(5), Pt(5), Pt(300), Pt(100))
    rect.text_frame.margin_left = rect.text_frame.margin_right = 0
    rect.text_frame.word_wrap = False
    rect.text_frame.text = "@@FIO@@"
    eff = pptx_engine._effective_run_style(
        style_slide, rect, rect.text_frame, rect.text_frame.paragraphs[0],
        rect.text_frame.paragraphs[0].runs[0]._r)
    check("review7: autoshape resolves otherStyle (size/family/bold)",
          eff == (44.0, "Arial", False), str(eff))
    arial = find_font_file("Arial")
    aval = "Участник Тестовый"
    overflow_at_44 = measure_text_width(aval, arial, 44) > 300
    s_note = pptx_engine.apply_field_to_slide(
        style_slide, PptxPlacement(0, rect.shape_id, token="@@FIO@@",
                                   mode="fragment"), aval)
    s_run = next(r for r in rect.text_frame.paragraphs[0].runs if "Участник" in r.text)
    check("review7: autoshape fragment uses otherStyle size",
          (not overflow_at_44) or
          (any("уменьшен" in n or "не поместился" in n for n in s_note)
           and s_run.font.size is not None and s_run.font.size.pt < 44),
          f"notes={s_note} w44={measure_text_width(aval, arial, 44):.0f}")
    rect2 = style_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(5), Pt(200), Pt(300), Pt(100))
    rect2.text_frame.margin_left = rect2.text_frame.margin_right = 0
    rect2.text_frame.word_wrap = False
    rect2.text_frame.text = "Шаблонный текст"
    w_note = pptx_engine.apply_field_to_slide(
        style_slide, PptxPlacement(0, rect2.shape_id, "Шаблонный текст", "", "whole"),
        aval)
    w_run = next((r for r in rect2.text_frame.paragraphs[0].runs if "Участник" in r.text),
                 None)
    check("review7: autoshape whole uses otherStyle size",
          (not overflow_at_44) or
          (any("уменьшен" in n or "не поместился" in n for n in w_note)
           and w_run is not None and w_run.font.size is not None
           and w_run.font.size.pt < 44),
          f"notes={w_note}")
    style_file = os.path.join(tmp, "styled_shape.pptx")
    style_deck.save(style_file)
    reloaded = Presentation(style_file).slides[0]
    reloaded_run = next(r for s in reloaded.shapes if s.has_text_frame
                        for r in s.text_frame.paragraphs[0].runs
                        if "Участник" in r.text)
    check("review7: autoshape size survives save/reopen",
          reloaded_run.font.size is not None
          and reloaded_run.font.size.pt < 44 or not overflow_at_44,
          str(reloaded_run.font.size))

    # ---- anonymized end-to-end fixture: 6 rows -> combined deck ---------
    fixture_pptx = os.path.join(tmp, "fixture_template.pptx")
    fixture_xlsx = os.path.join(tmp, "fixture_list.xlsx")
    fdeck = Presentation()
    fdeck.slide_width = Cm(28)
    fdeck.slide_height = Cm(20)
    fslide = fdeck.slides.add_slide(fdeck.slide_layouts[6])
    fio_box = fslide.shapes.add_textbox(Cm(1), Cm(2), Cm(9), Cm(1.5))
    fio_box.text_frame.word_wrap = False
    fio_par = fio_box.text_frame.paragraphs[0]
    fio_par.font.size = Pt(40)
    fio_par.font.name = "Arial"
    fio_par.add_run().text = "@@FIO@@"
    nom_box = fslide.shapes.add_textbox(Cm(1), Cm(6), Cm(18), Cm(2))
    nom_box.text_frame.text = "Номинация: «Тестовая номинация»"
    dep_box = fslide.shapes.add_textbox(Cm(1), Cm(10), Cm(18), Cm(2))
    dep_box.text_frame.text = "Тестовое управление"
    fdeck.save(fixture_pptx)
    fbook = openpyxl.Workbook()
    fsheet = fbook.active
    fsheet.append(["ФИО", "Номинация", "Управление"])
    anon_rows = [[f"Участник {ch}", f"Тестовая номинация {i}", f"Тестовое управление {i}"]
                 for i, ch in enumerate("АБВГД", 1)]
    anon_rows.append(["Участник Е С Очень Длинной Фамилией Для Проверки Автоподбора Шрифта",
                      "Тестовая номинация 6", "Тестовое управление 6"])
    for row in anon_rows:
        fsheet.append(row)
    fbook.save(fixture_xlsx)
    fbook.close()

    th, xh = _sha(fixture_pptx), _sha(fixture_xlsx)
    rh, rr = excel_reader.read_xlsx(fixture_xlsx)
    real_blocks = {b["shape_id"]: b for b in pptx_engine.list_text_shapes(Presentation(fixture_pptx))}
    nom_id = next(sid for sid, b in real_blocks.items() if "Номинация:" in b["text"])
    dep_id = next(sid for sid, b in real_blocks.items() if "Тестовое управление" in b["text"])
    fio_id = next(sid for sid, b in real_blocks.items() if "@@FIO@@" in b["text"])
    field_map = [
        Field("ФИО", [PptxPlacement(0, fio_id, real_blocks[fio_id]["text"], "@@FIO@@",
                                   "token")], required=True),
        Field("Номинация", [PptxPlacement(0, nom_id, real_blocks[nom_id]["text"],
                                         "Тестовая номинация", "fragment")], required=True),
        Field("Управление", [PptxPlacement(0, dep_id, real_blocks[dep_id]["text"], "",
                                          "whole")], required=True),
    ]
    real_out = os.path.join(tmp, "real")
    real_rep = generate(fixture_pptx, TEMPLATE_PPTX, field_map,
                        {name: rh.index(name) for name in ("ФИО", "Номинация", "Управление")},
                        rr, real_out, MODE_COMBINED, rh.index("ФИО"))
    real_result = Presentation(os.path.join(real_out, real_rep["created"][0][1]))
    text_rows = [[s.text_frame.text for s in slide.shapes if s.has_text_frame]
                 for slide in real_result.slides]
    check("iter6: workbook six slides in row order", len(text_rows) == 6 and all(
          all(rr[i][rh.index(name)] in " ".join(text_rows[i])
              for name in ("ФИО", "Номинация", "Управление"))
          for i in range(6)))
    check("iter6: nomination surrounds preserved", all(
          f"Номинация: «{rr[i][rh.index('Номинация')]}»" in text_rows[i]
          for i in range(6)))
    check("iter6: long name explicitly reviewed", any(
          row == 7 and any("шрифт" in note or "не поместился" in note for note in notes)
          for row, notes in real_rep["review"]))
    check("iter6: original template and Excel unchanged",
          th == _sha(fixture_pptx) and xh == _sha(fixture_xlsx))

    # ---- iter7: facsimile signature ------------------------------------------
    from app.core import signature as sigmod
    from app.core.models import SignatureSpec
    from PySide6.QtCore import Qt, QPointF
    from PySide6.QtGui import QImage, QPixmap
    import struct, zlib

    def _png_rgba(path, w, h):
        """Deterministic RGBA PNG: dark-blue diagonal strokes on alpha 0."""
        def chunk(tag, data):
            return (struct.pack(">I", len(data)) + tag + data
                    + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))
        raw = b"".join(
            b"\x00" + b"".join(
                bytes((11, 61, 145, 255)) if (x + y) % 9 < 4
                else bytes((0, 0, 0, 0)) for x in range(w))
            for y in range(h))
        with open(path, "wb") as fh:
            fh.write(b"\x89PNG\r\n\x1a\n"
                     + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0))
                     + chunk(b"IDAT", zlib.compress(raw))
                     + chunk(b"IEND", b""))

    sig_png = os.path.join(tmp, "sig.png")
    _png_rgba(sig_png, 200, 80)
    sig_jpg = os.path.join(tmp, "sig.jpg")
    QImage(200, 80, QImage.Format_RGB32).save(sig_jpg, "JPG")
    doc_w_pt = Cm(28) / 12700.0
    doc_h_pt = Cm(20) / 12700.0
    sig_rect = (doc_w_pt - 240, doc_h_pt - 140, doc_w_pt - 40, doc_h_pt - 40)

    check("iter7: png/jpg accepted, bad files rejected",
          sigmod.image_error(sig_png) is None
          and sigmod.image_error(sig_jpg) is None
          and sigmod.image_error("") is not None
          and sigmod.image_error(os.path.join(tmp, "nope.png")) is not None
          and sigmod.image_error(fixture_xlsx) is not None)

    check("iter7: spec validation covers missing pieces",
          sigmod.issues(SignatureSpec(True, sig_png, sig_rect), doc_w_pt, doc_h_pt) == []
          and sigmod.issues(SignatureSpec(True, sig_png, None), doc_w_pt, doc_h_pt)
          and sigmod.issues(SignatureSpec(True, "", sig_rect), doc_w_pt, doc_h_pt)
          and sigmod.issues(SignatureSpec(True, sig_png, (5, 5, 10, 10)),
                            doc_w_pt, doc_h_pt)
          and sigmod.issues(SignatureSpec(True, sig_png,
                                          (doc_w_pt + 10, 0, doc_w_pt + 40, 30)),
                            doc_w_pt, doc_h_pt)
          and not sigmod.issues(SignatureSpec(False, "", None),
                                doc_w_pt, doc_h_pt))

    sig_spec = SignatureSpec(True, sig_png, sig_rect)
    sig_rep = generate(fixture_pptx, TEMPLATE_PPTX, field_map,
                       {n: rh.index(n) for n in ("ФИО", "Номинация", "Управление")},
                       rr, os.path.join(tmp, "sig_out"), MODE_COMBINED,
                       rh.index("ФИО"), signature=sig_spec)
    sig_deck = Presentation(os.path.join(tmp, "sig_out", sig_rep["created"][0][1]))
    pics = [[s for s in sl.shapes
             if s.name == pptx_engine.SIGNATURE_PICTURE_NAME]
            for sl in sig_deck.slides]
    check("iter7: combined pptx — one signature picture on every slide",
          len(sig_deck.slides) == 6 and all(len(p) == 1 for p in pics))
    boxes = [(p[0].left, p[0].top, p[0].width, p[0].height) for p in pics]
    check("iter7: identical signature placement on all slides",
          len(set(boxes)) == 1)
    l0, t0, w0, h0 = boxes[0]
    # 200x80 px in a 200x100 pt rect -> scale 1.0, vertically centered
    exp_w, exp_h = int(200 * 12700), int(80 * 12700)
    check("iter7: aspect preserved, fit inside rect",
          abs(w0 - exp_w) < 3 * 12700 and abs(h0 - exp_h) < 3 * 12700
          and l0 >= int(sig_rect[0] * 12700) - 12700
          and t0 >= int(sig_rect[1] * 12700) - 12700
          and l0 + w0 <= int(sig_rect[2] * 12700) + 12700
          and t0 + h0 <= int(sig_rect[3] * 12700) + 12700,
          f"{l0},{t0} {w0}x{h0}")
    check("iter7: png transparency kept (RGBA image part)",
          pics[0][0].image.blob[25] in (4, 6))

    # idempotency: stamping the same slide twice must not stack copies
    idem = Presentation(fixture_pptx)
    pptx_engine.add_signature_to_slide(idem.slides[0], sig_spec,
                                       doc_w_pt, doc_h_pt)
    pptx_engine.add_signature_to_slide(idem.slides[0], sig_spec,
                                       doc_w_pt, doc_h_pt)
    check("iter7: re-stamping replaces, no duplicates",
          sum(1 for s in idem.slides[0].shapes
              if s.name == pptx_engine.SIGNATURE_PICTURE_NAME) == 1)

    # separate mode gets the same picture at the same coordinates
    sep_rep = generate(fixture_pptx, TEMPLATE_PPTX, field_map,
                       {n: rh.index(n) for n in ("ФИО", "Номинация", "Управление")},
                       rr[:1], os.path.join(tmp, "sig_sep"), MODE_SEPARATE,
                       rh.index("ФИО"), signature=sig_spec)
    sep_pic = next(s for s in Presentation(
        os.path.join(tmp, "sig_sep", sep_rep["created"][0][1])).slides[0].shapes
        if s.name == pptx_engine.SIGNATURE_PICTURE_NAME)
    check("iter7: separate pptx placement matches combined",
          (sep_pic.left, sep_pic.top, sep_pic.width, sep_pic.height) == boxes[0])

    # pdf mode stamps the image on every page too
    pdf_sig_rep = generate(T_PDF, TEMPLATE_PDF, pfields, cm, rows[:2],
                           os.path.join(tmp, "sig_pdf"), MODE_COMBINED, 0,
                           signature=SignatureSpec(True, sig_png, sig_rect))
    pdf_sig_doc = fitz.open(os.path.join(tmp, "sig_pdf",
                                         pdf_sig_rep["created"][0][1]))
    check("iter7: pdf pages carry the signature image",
          pdf_sig_doc.page_count == 2
          and all(pdf_sig_doc[i].get_images() for i in range(2)))
    pdf_sig_doc.close()

    # invalid signature blocks generation before touching the output dir
    bad_out = os.path.join(tmp, "sig_bad")
    bad_rep = generate(fixture_pptx, TEMPLATE_PPTX, field_map,
                       {n: rh.index(n) for n in ("ФИО", "Номинация", "Управление")},
                       rr[:1], bad_out, MODE_SEPARATE, rh.index("ФИО"),
                       signature=SignatureSpec(True, "", None))
    check("iter7: broken signature stops generation early",
          not bad_rep["created"]
          and any("подпись" in rsn for _r, rsn in bad_rep["skipped"]))

    # ---- iter7: canvas signature editing ---------------------------------
    from PySide6.QtTest import QTest
    from app.gui.canvas import MODE_SIGNATURE, PreviewCanvas
    gui_app2 = QApplication.instance() or QApplication([])
    canvas = PreviewCanvas()
    canvas.resize(800, 567)   # ~28x20 cm slide at ~1 px/pt
    canvas.set_image(QPixmap(793, 567), 793.0, 567.0)
    canvas.set_signature(sig_rect, QPixmap(sig_png))
    seen = []
    canvas.sigSignatureChanged.connect(lambda *r: seen.append(r))

    def draw(c, a, b):
        QTest.mousePress(c, Qt.LeftButton, Qt.NoModifier, a.toPoint())
        QTest.mouseMove(c, b.toPoint())
        QTest.mouseRelease(c, Qt.LeftButton, Qt.NoModifier, b.toPoint())

    def view_pt(rect, cw, ch):
        """doc-pt rect -> widget px endpoints at canvas size cw x ch."""
        s = min(cw / 793.0, ch / 567.0)
        mx, my = (cw - 793 * s) / 2, (ch - 567 * s) / 2
        return (QPointF(mx + rect[0] * s, my + rect[1] * s),
                QPointF(mx + rect[2] * s, my + rect[3] * s), s, mx, my)

    canvas.set_mode(MODE_SIGNATURE)
    canvas.show()
    a1, b1, _s1, _mx1, _my1 = view_pt(sig_rect, 800, 567)
    draw(canvas, a1, b1)
    check("iter7: drawn signature rect in slide coords",
          bool(seen) and abs(seen[-1][0] - sig_rect[0]) < 3
          and abs(seen[-1][2] - sig_rect[2]) < 3
          and abs(seen[-1][3] - sig_rect[3]) < 3, str(seen[-1] if seen else None))

    # same gesture at another widget size -> same document coordinates
    canvas.resize(1200, 850)
    a2, b2, _s2, _mx2, _my2 = view_pt(sig_rect, 1200, 850)
    draw(canvas, a2, b2)
    check("iter7: preview zoom does not change doc coordinates",
          bool(seen) and abs(seen[-1][0] - sig_rect[0]) < 3
          and abs(seen[-1][2] - sig_rect[2]) < 3,
          str(seen[-1] if seen else None))

    # move: press inside the rect and drag 30 px right / 15 px down
    canvas.resize(800, 567)
    _a3, _b3, s2, mx, my = view_pt(sig_rect, 800, 567)
    ctr = QPointF(mx + (sig_rect[0] + 100) * s2, my + (sig_rect[1] + 50) * s2)
    draw(canvas, ctr, QPointF(ctr.x() + 30, ctr.y() + 15))
    moved = seen[-1]
    check("iter7: move keeps size, shifts origin",
          abs(moved[0] - sig_rect[0] - 30 / s2) < 3
          and abs(moved[1] - sig_rect[1] - 15 / s2) < 3
          and abs((moved[2] - moved[0]) - (sig_rect[2] - sig_rect[0])) < 3,
          str(moved))

    # resize via the bottom-right handle (press lands 4px inside the handle,
    # so the corner moves by dragdelta - 4)
    hr = QPointF(mx + moved[2] * s2 - 4, my + moved[3] * s2 - 4)
    draw(canvas, hr, QPointF(hr.x() + 40, hr.y() + 20))
    grown = seen[-1]
    check("iter7: corner handle resizes, origin stays",
          abs(grown[0] - moved[0]) < 3
          and abs(grown[2] - moved[2] - 36 / s2) < 3
          and abs(grown[3] - moved[3] - 16 / s2) < 3, str(grown))
    canvas.close()

    # template/slide change drops the stored area and asks for a new one
    sig_state = AppState(template_path=fixture_pptx,
                         template_type=TEMPLATE_PPTX)
    sig_state.signature = SignatureSpec(True, sig_png, sig_rect)
    sig_step = StepTemplate(sig_state, lambda: None)
    sig_step._prs = Presentation(fixture_pptx)
    sig_step._doc_w, sig_step._doc_h = doc_w_pt, doc_h_pt
    sig_step._sig_toggled(True)
    sig_step._reset_signature_area()
    check("iter7: template/slide change clears the stored area",
          sig_state.signature.rect is None and sig_state.signature.enabled)
    sig_step.close()
    gui_app2.quit()

    # ---- perf: 150 rows -------------------------------------------------------
    big_rows = [[f"{n} Тестович{i}", REGIONS[i % len(REGIONS)]]
                for i, n in enumerate(NAMES * 15)]
    assert len(big_rows) == 150
    import tracemalloc, time
    tracemalloc.start()
    t0 = time.time()
    rep_big = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, big_rows,
                       os.path.join(tmp, "big"), MODE_SEPARATE, cm["ФИО"], set())
    t_pptx = time.time() - t0
    _c, peak1 = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    tracemalloc.start()
    t0 = time.time()
    rep_big2 = generate(T_PDF, TEMPLATE_PDF, pfields, cm, big_rows,
                        os.path.join(tmp, "bigpdf"), MODE_SEPARATE, cm["ФИО"], set())
    t_pdf = time.time() - t0
    _c, peak2 = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    mem_mb = max(peak1, peak2) / 1024 / 1024
    check("perf: 150 pptx", len(rep_big["created"]) == 150,
          f"{t_pptx:.1f}s, peak {peak1/1024/1024:.0f} MB")
    check("perf: 150 pdf", len(rep_big2["created"]) == 150,
          f"{t_pdf:.1f}s, peak {peak2/1024/1024:.0f} MB")
    PERF["pptx_150_s"] = round(t_pptx, 1)
    PERF["pdf_150_s"] = round(t_pdf, 1)
    PERF["peak_mem_mb"] = round(mem_mb)

    # ---- summary ------------------------------------------------------------
    ok = sum(1 for _n, c, _e in results if c)
    print(f"\n{ok}/{len(results)} checks passed")
    print("PERF:", PERF)
    if ok != len(results):
        sys.exit(1)


if __name__ == "__main__":
    main()
