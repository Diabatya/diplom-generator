"""Save / load template markup to ``<template>.mapping.json`` files."""

from __future__ import annotations

import json
import os

from .models import Field, TEMPLATE_PDF, TEMPLATE_PPTX

MAPPING_VERSION = 1


def mapping_path_for(template_path: str) -> str:
    base, _ = os.path.splitext(template_path)
    return base + ".mapping.json"


def has_saved_mapping(template_path: str) -> bool:
    return os.path.isfile(mapping_path_for(template_path))


def save_mapping(template_path: str, template_type: str, fields: list[Field]) -> str:
    path = mapping_path_for(template_path)
    data = {
        "version": MAPPING_VERSION,
        "template_type": template_type,
        "template_file": os.path.basename(template_path),
        "fields": [f.to_dict() for f in fields],
    }
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    return path


def load_mapping(template_path: str) -> tuple[str, list[Field]]:
    """Return (template_type, fields) stored next to the template."""
    path = mapping_path_for(template_path)
    with open(path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    ttype = data.get("template_type", TEMPLATE_PDF)
    if ttype not in (TEMPLATE_PDF, TEMPLATE_PPTX):
        ttype = TEMPLATE_PDF
    fields = [Field.from_dict(fd) for fd in data.get("fields", [])]
    return ttype, fields
