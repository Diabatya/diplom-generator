"""Access to bundled resource files (works both from sources and inside
the PyInstaller onefile exe)."""

from __future__ import annotations

import os
import sys


def resource_path(*parts: str) -> str:
    """Absolute path to a file inside ``app/resources`` — in the frozen exe
    resources live under ``sys._MEIPASS``."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return os.path.join(base, "app", "resources", *parts)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "resources", *parts)


def emblem_path() -> str:
    """FNS emblem PNG used in the application header."""
    return resource_path("fns_emblem.png")
