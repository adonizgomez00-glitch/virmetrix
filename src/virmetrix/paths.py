"""Rutas de recursos para modo desarrollo y empaquetado."""

from __future__ import annotations

import sys
from pathlib import Path


def resource_root() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    # src/virmetrix/paths.py -> project root
    return Path(__file__).resolve().parents[2]