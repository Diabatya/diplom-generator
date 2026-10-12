"""Shared application state passed between wizard steps."""

from __future__ import annotations

from dataclasses import dataclass, field

from app.core.models import Field


@dataclass
class AppState:
    template_path: str | None = None
    template_type: str | None = None  # "pptx" | "pdf"
    fields: list[Field] = field(default_factory=list)
    slide_sel: list[int] | None = None

    excel_path: str | None = None
    sheet_name: str | None = None
    headers: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)
    colmap: dict[str, int] = field(default_factory=dict)  # field name -> column index

    # ("field", field_name) or ("col", column_index) — what names output files
    filename_source: tuple = ("auto", None)
    filename_first_word: bool = False
    out_dir: str | None = None
    filename_col: int = -1
