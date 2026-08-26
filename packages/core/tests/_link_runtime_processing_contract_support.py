"""RED contract tests for consuming ``LinkRuntimeConfig`` in link processing.

These tests describe the public, target-aware boundary expected by converter
consumers.  They intentionally avoid runtime-global configuration so a caller
can project its request snapshot into one immutable policy object.
"""

from __future__ import annotations

import inspect
from dataclasses import replace
from io import BytesIO
from pathlib import Path

import pytest
from PIL import Image

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import process_markdown_links, resolve_embedded_links
from docwen_core.links._markdown_inline import escape_markdown_source_literal

pytestmark = pytest.mark.contract


def _write(path: Path, content: str = "") -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    return str(path)


def _write_png(path: Path) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    with BytesIO() as buffer:
        Image.new("RGB", (2, 2), (40, 80, 120)).save(buffer, format="PNG")
        path.write_bytes(buffer.getvalue())
    return str(path)


__all__ = (
    "LinkRuntimeConfig",
    "Path",
    "_write",
    "_write_png",
    "escape_markdown_source_literal",
    "inspect",
    "process_markdown_links",
    "pytest",
    "pytestmark",
    "replace",
    "resolve_embedded_links",
)
