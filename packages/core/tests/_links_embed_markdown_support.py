"""Tests for docwen_core.links — Markdown file embedding, single-embed
dispatch, and batch embed resolution.

Covers F-H2-009 (``process_embedded_md_file``) and
F-H2-026 (``process_single_embed``).
"""

from __future__ import annotations

import base64
from io import BytesIO
from pathlib import Path

import pytest
from PIL import Image

from docwen_core.links import (
    EmbeddedMdMode,
    NotFoundAction,
    process_embedded_md_file,
    process_single_embed,
    resolve_embedded_links,
)

pytestmark = pytest.mark.unit


def _write(path: Path, text: str) -> str:
    """Write *text* to *path* and return the absolute path string."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return str(path)


def _sample_image_bytes(image_format: str) -> bytes:
    with BytesIO() as buffer:
        Image.new("RGB", (2, 2), (40, 80, 120)).save(buffer, format=image_format)
        return buffer.getvalue()


_SAMPLE_PNG_BYTES = _sample_image_bytes("PNG")


def _data_uri() -> str:
    payload = base64.b64encode(_SAMPLE_PNG_BYTES).decode("ascii")
    return f"data:image/png;base64,{payload}"


__all__ = (
    "_SAMPLE_PNG_BYTES",
    "EmbeddedMdMode",
    "NotFoundAction",
    "Path",
    "_data_uri",
    "_sample_image_bytes",
    "_write",
    "process_embedded_md_file",
    "process_single_embed",
    "pytest",
    "pytestmark",
    "resolve_embedded_links",
)
