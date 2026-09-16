"""Batch generation engine (Qt-free, usable from a worker thread)."""

from __future__ import annotations

import os
import re

try:
    import pymupdf as fitz
except ImportError:
    import fitz  # type: ignore

from . import pdf_engine, pptx_engine
from .models import Field, TEMPLATE_PDF, TEMPLATE_PPTX

MODE_SEPARATE = "separate"
MODE_COMBINED = "combined"

EXISTS_OVERWRITE = "overwrite"   # перезаписать существующие
EXISTS_SKIP = "skip"             # пропустить существующие
EXISTS_SUFFIX = "suffix"         # добавить суффикс _new / _new2 / ...

_INVALID = re.compile(r'[\\/:*?"<>|]')
MAX_NAME_LEN = 80
MAX_PATH_LEN = 240  # safe Windows path limit (260 minus slack)


class Cancelled(Exception):
    pass


def sanitize_filename(name: str) -> str:
    name = _INVALID.sub("_", name).strip().strip(".")
    name = re.sub(r"\s+", " ", name)
    return name[:MAX_NAME_LEN].rstrip(" .") or "file"


def choose_filename_field(fields: list[Field]) -> int:
    """Index of the field used for file names ("Фамилия" / ФИО / first)."""
    for i, f in enumerate(fields):
        if "фамил" in f.name.lower():
            return i
    for i, f in enumerate(fields):
        if "фио" in f.name.lower():
            return i
    return 0


def row_values(fields: list[Field], colmap: dict[str, int], row: list[str]) -> dict[str, str]:
    out = {}
    for f in fields:
        ci = colmap.get(f.name, -1)
        out[f.name] = row[ci] if 0 <= ci < len(row) else ""
    return out


def _ext(template_type: str) -> str:
    return ".pptx" if template_type == TEMPLATE_PPTX else ".pdf"


def _combined_name(template_path: str, template_type: str) -> str:
    stem = os.path.splitext(os.path.basename(template_path))[0]
    return sanitize_filename(stem) + "_combined" + _ext(template_type)


def _base_name(row: list[str], filename_col: int, row_no: int) -> str:
    if 0 <= filename_col < len(row) and row[filename_col].strip():
        return row[filename_col]
    return f"row_{row_no}"


def planned_filenames(
    template_path: str,
    template_type: str,
    rows: list[list[str]],
    mode: str,
    filename_col: int,
    skip_rows: set[int] | None = None,
) -> list[str]:
    """Basenames that a run will try to create — used by the UI to detect
    collisions with existing files before generation starts."""
    skip_rows = skip_rows or set()
    if mode == MODE_COMBINED:
        return [_combined_name(template_path, template_type)] if any(
            i not in skip_rows for i in range(len(rows))) else []
    counters: dict = {}
    out = []
    for i, row in enumerate(rows):
        if i in skip_rows:
            continue
        out.append(_next_name(counters, _base_name(row, filename_col, i + 2)) + _ext(template_type))
    return out


def _fit_path(out_dir: str, name: str, ext: str) -> str:
    """Truncate ``name`` so the full path stays under the Windows limit."""
    room = MAX_PATH_LEN - len(out_dir) - len(ext) - 5
    if len(name) > room:
        name = name[: max(10, room)].rstrip(" .")
    return name + ext


def _resolve_collision(out_dir: str, name: str, ext: str, exists_mode: str):
    """Return (final_name, skip_flag) honoring the existing-files policy."""
    path = os.path.join(out_dir, _fit_path(out_dir, name, ext))
    if not os.path.exists(path):
        return _fit_path(out_dir, name, ext), False
    if exists_mode == EXISTS_OVERWRITE:
        return _fit_path(out_dir, name, ext), False
    if exists_mode == EXISTS_SKIP:
        return None, True
    # suffix: name_new.ext, name_new2.ext, ...
    n = 0
    while True:
        cand = f"{name}_new" if n == 0 else f"{name}_new{n + 1}"
        fname = _fit_path(out_dir, cand, ext)
        if not os.path.exists(os.path.join(out_dir, fname)):
            return fname, False
        n += 1


def generate(
    template_path: str,
    template_type: str,
    fields: list[Field],
    colmap: dict[str, int],
    rows: list[list[str]],
    out_dir: str,
    mode: str,
    filename_col: int,
    skip_rows: set[int] | None = None,
    exists_mode: str = EXISTS_OVERWRITE,
    progress_cb=None,
    cancel_cb=None,
) -> dict:
    """Run generation. Returns a report dict.

    ``rows`` are data rows (without header); indexes in ``skip_rows`` are
    0-based row indexes. progress_cb(done, total, label) is called per row.
    """
    skip_rows = skip_rows or set()
    os.makedirs(out_dir, exist_ok=True)

    total = len(rows)
    report = {
        "created": [],          # [(row_no, filename)]
        "skipped": [],          # [(row_no, reason)]
        "review": [],           # [(row_no, [notes])]
        "cancelled": False,
        "out_dir": out_dir,
        "report_path": None,
    }

    def progress(i, label):
        if progress_cb:
            progress_cb(i, total, label)

    def check_cancel():
        if cancel_cb and cancel_cb():
            raise Cancelled()

    try:
        if mode == MODE_COMBINED:
            _generate_combined(
                template_path, template_type, fields, colmap, rows,
                out_dir, skip_rows, exists_mode, progress, check_cancel, report,
            )
        else:
            _generate_separate(
                template_path, template_type, fields, colmap, rows,
                out_dir, filename_col, skip_rows, exists_mode,
                progress, check_cancel, report,
            )
    except Cancelled:
        report["cancelled"] = True

    report["report_path"] = _write_report(report, out_dir)
    return report


def _next_name(counters: dict, base: str) -> str:
    base = sanitize_filename(base)
    n = counters.get(base, 0) + 1
    counters[base] = n
    return base if n == 1 else f"{base}_{n}"


def _generate_separate(
    template_path, template_type, fields, colmap, rows, out_dir,
    filename_col, skip_rows, exists_mode, progress, check_cancel, report,
):
    counters: dict = {}
    ext = _ext(template_type)
    for i, row in enumerate(rows):
        check_cancel()
        row_no = i + 2
        progress(i, f"строка {row_no}")
        if i in skip_rows:
            report["skipped"].append((row_no, "пропущена по решению пользователя"))
            continue
        name = _next_name(counters, _base_name(row, filename_col, row_no))
        fname, skip = _resolve_collision(out_dir, name, ext, exists_mode)
        if skip:
            report["skipped"].append((row_no, "файл с таким именем уже существует — пропущен"))
            continue
        out_path = os.path.join(out_dir, fname)
        values = row_values(fields, colmap, row)
        try:
            if template_type == TEMPLATE_PPTX:
                notes = pptx_engine.render_single(template_path, fields, values, out_path)
            else:
                notes = pdf_engine.render_single(template_path, fields, values, out_path)
            report["created"].append((row_no, os.path.basename(out_path)))
            if notes:
                report["review"].append((row_no, notes))
        except Exception as exc:  # per-row fault isolation
            report["skipped"].append((row_no, f"ошибка: {exc}"))
    progress(len(rows), "готово")


def _generate_combined(
    template_path, template_type, fields, colmap, rows, out_dir,
    skip_rows, exists_mode, progress, check_cancel, report,
):
    ext = _ext(template_type)
    stem = _combined_name(template_path, template_type)[: -len(ext)]

    used = [(i, r) for i, r in enumerate(rows) if i not in skip_rows]
    for i, _row in enumerate(rows):
        if i in skip_rows:
            report["skipped"].append((i + 2, "пропущена по решению пользователя"))

    if not used:
        return

    fname, skip = _resolve_collision(out_dir, stem, ext, exists_mode)
    if skip:
        report["skipped"].append((0, "целевой файл уже существует — генерация пропущена"))
        return
    out_path = os.path.join(out_dir, fname)

    if template_type == TEMPLATE_PPTX:
        values_list = [row_values(fields, colmap, r) for _i, r in used]
        try:
            notes_per_row = pptx_engine.render_combined(template_path, fields, values_list, out_path)
            for (i, _r), notes in zip(used, notes_per_row):
                report["created"].append((i + 2, os.path.basename(out_path)))
                if notes:
                    report["review"].append((i + 2, notes))
        except Exception as exc:
            for i, _r in used:
                report["skipped"].append((i + 2, f"ошибка: {exc}"))
        progress(len(rows), "готово")
        return

    out_doc = fitz.open()
    try:
        for done, (i, row) in enumerate(used):
            check_cancel()
            progress(done, f"строка {i + 2}")
            values = row_values(fields, colmap, row)
            try:
                src, notes = pdf_engine.render_pdf_doc(template_path, fields, values)
                out_doc.insert_pdf(src)
                src.close()
                report["created"].append((i + 2, os.path.basename(out_path)))
                if notes:
                    report["review"].append((i + 2, notes))
            except Exception as exc:
                report["skipped"].append((i + 2, f"ошибка: {exc}"))
        if out_doc.page_count:
            out_doc.save(out_path, garbage=4, deflate=True)
    finally:
        out_doc.close()
    progress(len(rows), "готово")


def _write_report(report: dict, out_dir: str) -> str:
    path = os.path.join(out_dir, "generation_report.txt")
    lines = [
        "Отчёт о генерации",
        "=" * 40,
        f"Создано записей: {len(report['created'])}",
        f"Пропущено строк: {len(report['skipped'])}",
        f"Требуют ручной проверки: {len(report['review'])}",
        f"Отменено пользователем: {'да' if report['cancelled'] else 'нет'}",
        "",
    ]
    if report["skipped"]:
        lines.append("Пропущенные строки:")
        for row_no, reason in report["skipped"]:
            lines.append(f"  строка {row_no}: {reason}")
        lines.append("")
    if report["review"]:
        lines.append("Требуют ручной проверки (автоподбор шрифта и т.п.):")
        for row_no, notes in report["review"]:
            for n in notes:
                lines.append(f"  строка {row_no}: {n}")
        lines.append("")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))
    return path
