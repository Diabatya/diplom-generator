"""Data model for template markup (fields and placements)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

TEMPLATE_PPTX = "pptx"
TEMPLATE_PDF = "pdf"

ALIGN_LEFT = "left"
ALIGN_CENTER = "center"
ALIGN_RIGHT = "right"


@dataclass
class PdfPlacement:
    """A rectangle on a PDF page where a field value is drawn."""

    page: int
    rect: tuple[float, float, float, float]  # x0, y0, x1, y1 in PDF points
    font_path: str = ""
    font_size: float = 24.0
    color: str = "#000000"
    align: str = ALIGN_CENTER
    fill_color: str = ""  # "" = auto (edge sampling); "#rrggbb" = manual

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": "pdf_rect",
            "page": self.page,
            "rect": list(self.rect),
            "font_path": self.font_path,
            "font_size": self.font_size,
            "color": self.color,
            "align": self.align,
            "fill_color": self.fill_color,
        }

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "PdfPlacement":
        return cls(
            page=int(d["page"]),
            rect=tuple(float(v) for v in d["rect"]),
            font_path=str(d.get("font_path", "")),
            font_size=float(d.get("font_size", 24.0)),
            color=str(d.get("color", "#000000")),
            align=str(d.get("align", ALIGN_CENTER)),
            fill_color=str(d.get("fill_color", "")),
        )


@dataclass
class PptxPlacement:
    """A text shape on a PPTX slide whose text is replaced by a field value.

    ``pattern`` is the shape's original text at bind time, ``token`` is the
    substring to replace inside it. ``mode`` selects the replacement:

    - ``whole``    — replace the entire shape text (default; also used for
                     old mappings saved before ``mode`` existed);
    - ``fragment`` — replace only ``token`` (a user-chosen substring) inside
                     ``pattern``;
    - ``token``    — same as fragment, but ``token`` was auto-detected
                     (``@@FIO@@``, ``{...}`` etc.).
    """

    slide: int
    shape_id: int
    pattern: str = ""
    token: str = ""
    mode: str = "whole"
    cell: tuple[int, int] | None = None  # (row, col) inside a table shape

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": "pptx_shape",
            "slide": self.slide,
            "shape_id": self.shape_id,
            "pattern": self.pattern,
            "token": self.token,
            "mode": self.mode,
            "cell": list(self.cell) if self.cell else None,
        }

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "PptxPlacement":
        mode = str(d.get("mode", "whole"))
        if mode not in ("whole", "fragment", "token"):
            mode = "whole"
        cell = d.get("cell")
        return cls(
            slide=int(d["slide"]),
            shape_id=int(d["shape_id"]),
            pattern=str(d.get("pattern", "")),
            token=str(d.get("token", "")),
            mode=mode,
            cell=(int(cell[0]), int(cell[1])) if cell else None,
        )


Placement = PdfPlacement | PptxPlacement


@dataclass
class SignatureSpec:
    """Optional facsimile signature stamped on every output slide/page.

    ``rect`` is (x0, y0, x1, y1) in document points (PPTX: EMU/12700,
    PDF: native page points) measured on the template the user previewed.
    The image is aspect-fit inside the rect, never stretched or cropped.
    """

    enabled: bool = False
    image_path: str = ""
    rect: tuple[float, float, float, float] | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "enabled": self.enabled,
            "image_path": self.image_path,
            "rect": list(self.rect) if self.rect else None,
        }

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "SignatureSpec":
        rect = d.get("rect")
        return cls(
            enabled=bool(d.get("enabled", False)),
            image_path=str(d.get("image_path", "")),
            rect=(tuple(float(v) for v in rect)
                  if isinstance(rect, (list, tuple)) and len(rect) == 4
                  else None),
        )


@dataclass
class Field:
    """A named placeholder bound to one or more placements on the template."""

    name: str
    placements: list[Placement] = field(default_factory=list)
    source: str = "column"
    constant: str = ""
    required: bool = False

    def to_dict(self) -> dict[str, Any]:
        return {"name": self.name, "placements": [p.to_dict() for p in self.placements],
                "source": self.source, "constant": self.constant, "required": self.required}

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "Field":
        f = cls(name=str(d["name"]), source=str(d.get("source", "column")),
                constant=str(d.get("constant", "")), required=bool(d.get("required", False)))
        for pd in d.get("placements", []):
            if pd.get("type") == "pdf_rect":
                f.placements.append(PdfPlacement.from_dict(pd))
            elif pd.get("type") == "pptx_shape":
                f.placements.append(PptxPlacement.from_dict(pd))
        return f


def placement_from_dict(d: dict[str, Any]) -> Placement | None:
    t = d.get("type")
    if t == "pdf_rect":
        return PdfPlacement.from_dict(d)
    if t == "pptx_shape":
        return PptxPlacement.from_dict(d)
    return None
