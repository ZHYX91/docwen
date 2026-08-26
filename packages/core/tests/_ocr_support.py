"""Tests for the shared OCR wrapper."""

from __future__ import annotations

import sys
import types
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def _first_existing_font(candidates: tuple[str, ...]) -> Path | None:
    for candidate in candidates:
        path = Path(candidate)
        if path.is_file():
            return path
    return None


def _write_ocr_input(tmp_path: Path, name: str = "sample.png") -> Path:
    image_path = tmp_path / name
    image_path.write_bytes(b"ocr input")
    return image_path


__all__ = (
    "Path",
    "_first_existing_font",
    "_write_ocr_input",
    "pytest",
    "pytestmark",
    "sys",
    "types",
)
