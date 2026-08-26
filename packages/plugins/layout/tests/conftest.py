"""Shared fixtures for Layout plugin tests."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[4]

LOCAL_SRC_PATHS = [
    PROJECT_ROOT,
    PROJECT_ROOT / "packages" / "core" / "src",
    PROJECT_ROOT / "packages" / "runtime" / "src",
    PROJECT_ROOT / "packages" / "plugins" / "layout" / "src",
]

for path in reversed(LOCAL_SRC_PATHS):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))


def _create_test_pdf(path: Path, pages: int = 1, text: str = "Hello PDF") -> None:
    """Create a minimal test PDF with *pages* pages, each containing *text*."""
    import fitz  # PyMuPDF

    doc = fitz.open()
    for i in range(pages):
        page = doc.new_page(width=595, height=842)  # A4
        page.insert_text((72, 72), f"{text} — page {i + 1}", fontsize=12)
    doc.save(str(path))
    doc.close()


@pytest.fixture
def sample_pdf_path(tmp_path: Path) -> Path:
    """A single-page PDF with simple text."""
    path = tmp_path / "sample.pdf"
    _create_test_pdf(path, pages=1, text="Test document")
    return path


@pytest.fixture
def sample_multi_page_pdf_path(tmp_path: Path) -> Path:
    """A 3-page PDF."""
    path = tmp_path / "multi.pdf"
    _create_test_pdf(path, pages=3, text="Page content")
    return path
