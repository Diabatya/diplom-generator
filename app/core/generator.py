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
EXISTS_SUFFIX = "suffix"         # нумеровать свободные имена _2, _3, ...

_INVALID = re.compile(r'[\\/:*?"<>|]')
MAX_NAME_LEN = 80
MAX_PATH_LEN = 240  # safe Windows path limit (260 minus slack)

# Windows device names — illegal as filenames in ANY case and with any ext.
_RESERVED_NAMES = (
    {"con", "prn", "aux", "nul"}
    | {f"com{i}" for i in range(1, 10)}
    | {f"lpt{i}" for i in range(1, 10)}
)
_MAX_COLLISION_TRIES = 9999  # hard bound for suffix search


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
        if f.source == "keep":
            continue
        if f.source == "constant":
            out[f.name] = f.constant
            continue
        ci = colmap.get(f.name, -1)
        if 0 <= ci < len(row) and (row[ci].strip() or f.required):
            out[f.name] = row[ci]
    return out


def _ext(template_type: str) -> str:
    return ".pptx" if template_type == TEMPLATE_PPTX else ".pdf"


def _combined_name(template_path: str, template_type: str) -> str:
    stem = os.path.splitext(os.path.basename(template_path))[0]
    return sanitize_filename(stem) + "_combined" + _ext(template_type)


def _base_name(row: list[str], filename_col: int, row_no: int,
               first_word: bool = False) -> str:
    if 0 <= filename_col < len(row):
        value = row[filename_col].strip()
        if first_word and value:
            value = value.split()[0]
        if value:
            return value
    return f"row_{row_no}"


def planned_filenames(
    template_path: str,
    template_type: str,
    rows: list[list[str]],
    mode: str,
    filename_col: int,
    skip_rows: set[int] | None = None,
    first_word: bool = False,
    out_dir: str = ".",
) -> list[str]:
    """Basenames that a run will try to create — used by the UI to detect
    collisions with existing files before generation starts. Pass the real
    ``out_dir`` so truncation matches the generator exactly."""
    skip_rows = skip_rows or set()
    if mode == MODE_COMBINED:
        if not any(i not in skip_rows for i in range(len(rows))):
            return []
        ext = _ext(template_type)
        stem = _combined_name(template_path, template_type)[:-len(ext)]
        fname = _fit_path(out_dir, stem, ext)
        return [fname] if fname else []
    ext = _ext(template_type)
    bases = [
        _base_name(row, filename_col, i + 2, first_word)
        for i, row in enumerate(rows) if i not in skip_rows
    ]
    # same planning the generator uses (pre-policy), so the UI checks the
    # exact final names
    return [fn for fn, _r in plan_batch_names(bases, out_dir, ext,
                                              EXISTS_OVERWRITE)
            if fn is not None]


def _fit_path(out_dir: str, name: str, ext: str, suffix: str = "") -> str | None:
    """Truncate ``name`` so that ``dir/name<suffix><ext>`` fits the Windows
    path limit. ``suffix`` (``_2`` etc.) is reserved space, never truncated.
    Returns None when the target directory alone is already too long."""
    room = (MAX_PATH_LEN - len(out_dir) - len(os.sep) - len(ext)
            - len(suffix))
    if room < 1:
        return None
    if len(name) > room:
        name = name[:room].rstrip(" .")
    if not name:
        return None
    if not suffix and name.casefold() in _RESERVED_NAMES:
        if room < 2:
            return None
        name = name[:room - 1].rstrip(" .") + "_"
    return name + suffix + ext


def _unreserved_stem(stem: str) -> str:
    """Guard against Windows device names (CON, NUL, COM1, ...) in any case."""
    return stem + "_" if stem.lower() in _RESERVED_NAMES else stem


def plan_batch_names(
    bases: list[str],
    out_dir: str,
    ext: str,
    exists_mode: str,
) -> list[tuple[str | None, str | None]]:
    """Reserve a unique final filename for every entry of the batch up-front.

    Returns a per-index list of ``(filename, skip_reason)`` — exactly one of
    the two is set. Guarantees: no two entries share a name (case-insensitive,
    as on Windows), suffixes always fit the path limit, and the search
    terminates. ``exists_mode`` governs only clashes with files that were in
    the folder *before* the run — rows never overwrite each other.
    """
    existing = (
        {n.lower() for n in os.listdir(out_dir)}
        if os.path.isdir(out_dir) else set()
    )
    reserved: set[str] = set()   # lowercase names assigned in this batch
    out: list[tuple[str | None, str | None]] = []
    for raw in bases:
        stem = _unreserved_stem(sanitize_filename(raw))
        fname: str | None = None
        reason: str | None = None
        for n in range(1, _MAX_COLLISION_TRIES + 1):
            suffix = "" if n == 1 else f"_{n}"
            cand = _fit_path(out_dir, stem, ext, suffix)
            if cand is None:
                reason = ("слишком длинная папка результатов: не хватает "
                          "места даже для минимального имени файла")
                break
            key = cand.lower()
            in_batch = key in reserved
            on_disk = key in existing
            if not in_batch and not on_disk:
                fname = cand
            elif in_batch:
                continue            # another row of this batch owns it
            elif exists_mode == EXISTS_OVERWRITE:
                fname = cand        # may overwrite a pre-existing file
            elif exists_mode == EXISTS_SKIP:
                reason = "файл с таким именем уже существует — пропущен"
                break
            else:                   # EXISTS_SUFFIX
                continue
            break
        else:
            reason = "не удалось подобрать свободное имя файла"
        if fname is not None:
            reserved.add(fname.lower())
            existing.add(fname.lower())   # later rows see it too
        out.append((fname, reason))
    return out


def _resolve_collision(out_dir: str, name: str, ext: str, exists_mode: str):
    """Return (final_name, skip_flag) honoring the existing-files policy."""
    fname, reason = plan_batch_names([name], out_dir, ext, exists_mode)[0]
    return fname, fname is None


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
    first_word: bool = False,
    slide_sel: list[int] | None = None,
    progress_cb=None,
    cancel_cb=None,
) -> dict:
    """Run generation. Returns a report dict.

    ``rows`` are data rows (without header); indexes in ``skip_rows`` are
    0-based row indexes. progress_cb(done, total, label) is called per row.
    ``slide_sel`` (PPTX only): original indexes of the template slides that
    make up one output document; None uses the whole deck.
    """
    skip_rows = set(skip_rows or ())
    empty_rows = {i for i, row in enumerate(rows) if not any(str(v).strip() for v in row)}
    skip_rows.update(empty_rows)
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
    missing = [f.name for f in fields if f.required and
               (f.source == "column" and colmap.get(f.name, -1) < 0 or
                f.source == "keep" or f.source == "constant" and not f.constant.strip())]
    if missing:
        report["skipped"].append((0, "не задано обязательное поле: " + ", ".join(missing)))
        report["report_path"] = _write_report(report, out_dir)
        return report
    if template_type == TEMPLATE_PPTX:
        try:
            issues = pptx_engine.required_placement_issues(
                pptx_engine.open_template(template_path), fields, slide_sel)
        except Exception as exc:
            issues = [f"не удалось проверить шаблон PPTX: {exc}"]
        if issues:
            report["skipped"].extend((0, issue) for issue in issues)
            report["report_path"] = _write_report(report, out_dir)
            return report
    invalid_rows = {}
    for i, row in enumerate(rows):
        for f in fields:
            if f.required and f.source == "column":
                ci = colmap[f.name]
                if ci >= len(row) or not row[ci].strip():
                    invalid_rows[i] = f"пустое обязательное поле «{f.name}»"
                    break
    skip_rows.update(invalid_rows)

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
                out_dir, skip_rows, empty_rows, invalid_rows, exists_mode, slide_sel,
                progress, check_cancel, report,
            )
        else:
            _generate_separate(
                template_path, template_type, fields, colmap, rows,
                out_dir, filename_col, skip_rows, empty_rows, invalid_rows, exists_mode, first_word,
                slide_sel, progress, check_cancel, report,
            )
    except Cancelled:
        report["cancelled"] = True

    report["report_path"] = _write_report(report, out_dir)
    return report


def _verify_output(path: str, template_type: str):
    if not os.path.isfile(path) or not os.path.getsize(path):
        raise OSError("файл результата не создан")
    if template_type == TEMPLATE_PPTX:
        prs = pptx_engine.open_template(path)
        if not len(prs.slides):
            raise ValueError("сохранённая презентация не содержит слайдов")
    else:
        with fitz.open(path, filetype="pdf") as doc:
            if not doc.page_count:
                raise ValueError("сохранённый PDF не содержит страниц")


def _atomic_save(render_fn, out_path: str, template_type: str):
    """Render, reopen, then atomically replace the target."""
    tmp_path = f"{out_path}.tmp{os.getpid()}"
    try:
        notes = render_fn(tmp_path)
        _verify_output(tmp_path, template_type)
        os.replace(tmp_path, out_path)
        return notes
    except Exception:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise


def _generate_separate(
    template_path, template_type, fields, colmap, rows, out_dir,
    filename_col, skip_rows, empty_rows, invalid_rows, exists_mode, first_word, slide_sel,
    progress, check_cancel, report,
):
    ext = _ext(template_type)
    used = [(i, r) for i, r in enumerate(rows) if i not in skip_rows]
    bases = [_base_name(r, filename_col, i + 2, first_word) for i, r in used]
    plan = plan_batch_names(bases, out_dir, ext, exists_mode)
    for i, _row in enumerate(rows):
        if i in skip_rows:
            report["skipped"].append((i + 2, invalid_rows.get(i,
                                     "пустая строка Excel" if i in empty_rows
                                     else "пропущена по решению пользователя")))
    for done, ((i, row), (fname, reason)) in enumerate(zip(used, plan)):
        check_cancel()
        row_no = i + 2
        progress(done, f"строка {row_no}")
        if fname is None:
            report["skipped"].append((row_no, reason))
            continue
        out_path = os.path.join(out_dir, fname)
        values = row_values(fields, colmap, row)
        engine = (pptx_engine if template_type == TEMPLATE_PPTX else pdf_engine)
        try:
            if template_type == TEMPLATE_PPTX:
                notes = _atomic_save(
                    lambda p: engine.render_single(
                        template_path, fields, values, p, slide_sel=slide_sel),
                    out_path, template_type)
            else:
                notes = _atomic_save(
                    lambda p: engine.render_single(template_path, fields, values, p),
                    out_path, template_type)
            # counted only after the file is really written
            report["created"].append((row_no, fname))
            if notes:
                report["review"].append((row_no, notes))
        except Exception as exc:  # per-row fault isolation
            report["skipped"].append((row_no, f"ошибка: {exc}"))
    progress(len(rows), "готово")


def _generate_combined(
    template_path, template_type, fields, colmap, rows, out_dir,
    skip_rows, empty_rows, invalid_rows, exists_mode, slide_sel, progress, check_cancel, report,
):
    ext = _ext(template_type)
    stem = _combined_name(template_path, template_type)[: -len(ext)]

    used = [(i, r) for i, r in enumerate(rows) if i not in skip_rows]
    for i, _row in enumerate(rows):
        if i in skip_rows:
            report["skipped"].append((i + 2, invalid_rows.get(i,
                                     "пустая строка Excel" if i in empty_rows
                                     else "пропущена по решению пользователя")))

    if not used:
        return

    fname, skip = _resolve_collision(out_dir, stem, ext, exists_mode)
    if skip or fname is None:
        report["skipped"].append((0, "целевой файл недоступен — генерация пропущена"))
        return
    out_path = os.path.join(out_dir, fname)

    if template_type == TEMPLATE_PPTX:
        values_list = [row_values(fields, colmap, r) for _i, r in used]
        pending_created: list = []
        pending_review: list = []

        def render(tmp_path):
            return pptx_engine.render_combined(
                template_path, fields, values_list, tmp_path,
                slide_sel=slide_sel,
                progress_cb=lambda d: progress(d, f"слайд {d + 1}"),
                cancel_cb=check_cancel)
        try:
            notes_per_row = _atomic_save(render, out_path, template_type)
            for (i, _r), notes in zip(used, notes_per_row):
                pending_created.append((i + 2, fname))
                if notes:
                    pending_review.append((i + 2, notes))
        except Cancelled:
            raise
        except Exception as exc:
            for i, _r in used:
                report["skipped"].append((i + 2, f"ошибка: {exc}"))
        # file is counted only after it is really saved
        report["created"].extend(pending_created)
        report["review"].extend(pending_review)
        progress(len(rows), "готово")
        return

    out_doc = fitz.open()
    pending_created = []
    pending_review = []
    try:
        for done, (i, row) in enumerate(used):
            check_cancel()
            progress(done, f"строка {i + 2}")
            values = row_values(fields, colmap, row)
            try:
                src, notes = pdf_engine.render_pdf_doc(template_path, fields, values)
                out_doc.insert_pdf(src)
                src.close()
                pending_created.append((i + 2, fname))
                if notes:
                    pending_review.append((i + 2, notes))
            except Exception as exc:
                report["skipped"].append((i + 2, f"ошибка: {exc}"))
        if out_doc.page_count:
            tmp_path = f"{out_path}.tmp{os.getpid()}"
            try:
                check_cancel()
                out_doc.save(tmp_path, garbage=4, deflate=True)
                _verify_output(tmp_path, TEMPLATE_PDF)
                check_cancel()
                os.replace(tmp_path, out_path)
            except Cancelled:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
                raise
            except Exception as exc:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
                for i, _r in used:
                    if (i + 2, fname) in pending_created:
                        report["skipped"].append((i + 2, f"ошибка сохранения: {exc}"))
                pending_created = []
            else:
                report["created"].extend(pending_created)
                report["review"].extend(pending_review)
        else:
            report["skipped"].append((0, "нечего сохранять — файл не создан"))
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
