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
    dup_pptx = os.path.join(tmp, "dup_tok.pptx")
    prs_d = Presentation()
    prs_d.slide_width = Cm(29.7)
    prs_d.slide_height = Cm(21)
    sld = prs_d.slides.add_slide(prs_d.slide_layouts[6])
    tbA = sld.shapes.add_textbox(Cm(2), Cm(5), Cm(8), Cm(3))
    ra = tbA.text_frame.paragraphs[0].add_run()
    ra.text = "№ @@FIO@@"
    ra.font.size = Pt(36)
    tbB = sld.shapes.add_textbox(Cm(2), Cm(10), Cm(8), Cm(3))
    rb = tbB.text_frame.paragraphs[0].add_run()
    rb.text = "№ @@FIO@@ + @@FIO@@"
    rb.font.size = Pt(36)
    tbC = sld.shapes.add_textbox(Cm(2), Cm(15), Cm(8), Cm(3))
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
          szB and szC and all(abs(s - szC[0]) < 0.01 for s in szB),
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

    real_template = "/Users/denis/Downloads/Диплом ЛВТ.pptx"
    real_excel = "/Users/denis/Downloads/test-participants.xlsx"
    if os.path.isfile(real_template) and os.path.isfile(real_excel):
        th, xh = _sha(real_template), _sha(real_excel)
        rh, rr = excel_reader.read_xlsx(real_excel)
        real_blocks = pptx_engine.list_text_shapes(Presentation(real_template))
        signatures = {"ФИО": "ИВАНОВ ИВАН ИВАНОВИЧ", "Номинация": "Номинация: «",
                      "Управление": "УФНС России по субъекту РФ"}
        field_map = []
        for field_name, signature in signatures.items():
            b = next(b for b in real_blocks if signature in b["text"])
            token = b["text"].split("«", 1)[1].split("»", 1)[0] if field_name == "Номинация" else ""
            field_map.append(Field(field_name, [PptxPlacement(0, b["shape_id"], b["text"],
                                token, "fragment" if token else "whole")], required=True))
        real_out = os.path.join(tmp, "real")
        real_rep = generate(real_template, TEMPLATE_PPTX, field_map,
                            {name: rh.index(name) for name in signatures}, rr,
                            real_out, MODE_COMBINED, rh.index("ФИО"))
        real_result = Presentation(os.path.join(real_out, real_rep["created"][0][1]))
        text_rows = [[s.text_frame.text for s in slide.shapes if s.has_text_frame]
                     for slide in real_result.slides]
        check("iter6: real workbook six slides in row order", len(text_rows) == 6 and all(
              all(rr[i][rh.index(name)] in " ".join(text_rows[i]) for name in signatures)
              for i in range(6)))
        check("iter6: nomination surrounds preserved", all(
              f"Номинация: «{rr[i][rh.index('Номинация')]}»" in text_rows[i]
              for i in range(6)))
        check("iter6: long name explicitly reviewed", any(
              row == 7 and any("шрифт" in note for note in notes)
              for row, notes in real_rep["review"]))
        check("iter6: original template and Excel unchanged", th == _sha(real_template)
              and xh == _sha(real_excel))

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
