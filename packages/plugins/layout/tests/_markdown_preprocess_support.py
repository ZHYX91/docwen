"""Tests for Markdown embed preprocessing in the layout-to-markdown pipeline.

Verifies that ``![[...]]`` wiki-style embed links in extracted Markdown are
resolved via the shared ``docwen_core.links`` dispatch (F-H2-026),
with recursive expansion for nested embeds (F-H2-009).

Also covers the ``process_markdown_links`` orchestrator (F-H2-019) as the
single entry point for link processing, and verifies non-embed wiki link
and Markdown link handling (F-H2-020) in the converter integration path.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image

from docwen_core.links import resolve_embedded_links

pytestmark = pytest.mark.unit


def _write(path: Path, text: str) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return str(path)


def _write_image(path: Path, image_format: str, *, color: tuple[int, int, int] = (1, 2, 3)) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with Image.new("RGB", (2, 2), color) as image:
        image.save(path, format=image_format)
    return path


__all__ = (
    "Path",
    "_write",
    "_write_image",
    "pytest",
    "pytestmark",
    "resolve_embedded_links",
)
