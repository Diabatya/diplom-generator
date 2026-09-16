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
    path = mapping_io.save_mapping(T_PPTX, TEMPLATE_PPTX, fields)
    check("c7: mapping file written", os.path.isfile(path), path)
    ttype, loaded = mapping_io.load_mapping(T_PPTX)
    check("c7: mapping roundtrip",
          ttype == "pptx" and len(loaded) == 2
          and loaded[0].name == "ФИО" and len(loaded[0].placements) == 2)
    # same for pdf
    mapping_io.save_mapping(T_PDF, TEMPLATE_PDF, pfields)
    tt2, ld2 = mapping_io.load_mapping(T_PDF)
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
    check("exists: suffix _new", all("_new" in n for n in names_suf), str(names_suf))
    rep_suf2 = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, rows[:3], out_e,
                        MODE_SEPARATE, cm["ФИО"], set(), exists_mode=EXISTS_SUFFIX)
    names_suf2 = sorted(fn for _r, fn in rep_suf2["created"])
    check("exists: suffix _new2", all("_new2" in n for n in names_suf2), str(names_suf2))
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

    # ---- P1.1: fragment mode ------------------------------------------------
    reg_sid = fields[1].placements[0].shape_id
    frag_fields = [
        fields[0],
        Field("Регион", [PptxPlacement(
            0, reg_sid, pattern="Награждается житель региона РЕГИОН за победу",
            token="РЕГИОН", mode="fragment")]),
    ]
    out_f = os.path.join(tmp, "frag")
    rep_f = generate(T_PPTX, TEMPLATE_PPTX, frag_fields, cm, [rows[0]], out_f,
                     MODE_SEPARATE, cm["ФИО"], set())
    texts_f = pptx_texts(os.path.join(out_f, rep_f["created"][0][1]))
    check("fragment: surrounding text kept",
          any("Награждается житель региона Москва за победу" in t for t in texts_f),
          str(texts_f))
    # back-compat: placement dict without "mode" behaves as whole
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

    # ---- perf: 150 rows -------------------------------------------------------
    big_rows = [[f"{n} Тестович{i}", REGIONS[i % len(REGIONS)]]
                for i, n in enumerate(NAMES * 15)]
    assert len(big_rows) == 150
    import resource, time
    t0 = time.time()
    mem0 = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    rep_big = generate(T_PPTX, TEMPLATE_PPTX, fields, cm, big_rows,
                       os.path.join(tmp, "big"), MODE_SEPARATE, cm["ФИО"], set())
    t_pptx = time.time() - t0
    mem1 = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    t0 = time.time()
    rep_big2 = generate(T_PDF, TEMPLATE_PDF, pfields, cm, big_rows,
                        os.path.join(tmp, "bigpdf"), MODE_SEPARATE, cm["ФИО"], set())
    t_pdf = time.time() - t0
    mem2 = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    check("perf: 150 pptx", len(rep_big["created"]) == 150,
          f"{t_pptx:.1f}s, maxrss {mem1/1024/1024:.0f} MB")
    check("perf: 150 pdf", len(rep_big2["created"]) == 150,
          f"{t_pdf:.1f}s, maxrss {mem2/1024/1024:.0f} MB")
    PERF["pptx_150_s"] = round(t_pptx, 1)
    PERF["pdf_150_s"] = round(t_pdf, 1)
    PERF["maxrss_mb"] = round(mem2 / 1024 / 1024)

    # ---- summary ------------------------------------------------------------
    ok = sum(1 for _n, c, _e in results if c)
    print(f"\n{ok}/{len(results)} checks passed")
    print("PERF:", PERF)
    if ok != len(results):
        sys.exit(1)


if __name__ == "__main__":
    main()
