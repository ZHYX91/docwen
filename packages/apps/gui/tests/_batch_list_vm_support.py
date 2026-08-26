"""Model-state tests for BatchListViewModel.

Tests verify the ViewModel's observable state transitions, filter/sort
operations, file entry management, and signal emissions. No QApplication
required — pure state tests.
"""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from docwen_core.detection import OOXML_SIGNATURE_VALIDATION_UNAVAILABLE, SUPPORTED_EXTENSION_FORMATS
from docwen_core.formats.categories import get_category
from docwen_core.models import AdmissionDecision
from docwen_gui.view_models.batch_list_vm import (
    CATEGORY_ORDER,
    BatchFileEntry,
    BatchListViewModel,
    _sort_value,
    format_size,
    should_pulse_processing_transition,
)

pytestmark = pytest.mark.unit


@pytest.fixture
def vm() -> BatchListViewModel:
    return BatchListViewModel()


def _synthetic_file_resolver(path: str) -> dict[str, str] | None:
    """Resolve deliberate synthetic paths from Core's public registry."""

    fmt = SUPPORTED_EXTENSION_FORMATS.get(Path(path).suffix.lower())
    if fmt is None:
        return None
    category = "markdown" if fmt in {"markdown", "txt"} else get_category(fmt)
    return {"detected_format": fmt, "workflow_category": category}


def _add_synthetic(
    vm: BatchListViewModel,
    paths: list[str],
) -> tuple[list[str], list[tuple[str, str]]]:
    return vm.add_files(paths, file_resolver=_synthetic_file_resolver)


@pytest.fixture
def sample_entries():
    """Create sample BatchFileEntry objects for testing."""
    return [
        BatchFileEntry(
            file_path="/test/doc1.docx",
            file_name="doc1.docx",
            detected_format="docx",
            workflow_category="document",
            size_bytes=1024,
            status="pending",
        ),
        BatchFileEntry(
            file_path="/test/sheet1.xlsx",
            file_name="sheet1.xlsx",
            detected_format="xlsx",
            workflow_category="spreadsheet",
            size_bytes=2048,
            status="pending",
        ),
        BatchFileEntry(
            file_path="/test/img1.png",
            file_name="img1.png",
            detected_format="png",
            workflow_category="image",
            size_bytes=512,
            status="pending",
        ),
        BatchFileEntry(
            file_path="/test/layout.pdf",
            file_name="layout.pdf",
            detected_format="pdf",
            workflow_category="layout",
            size_bytes=4096,
            status="pending",
        ),
        BatchFileEntry(
            file_path="/test/readme.md",
            file_name="readme.md",
            detected_format="markdown",
            workflow_category="markdown",
            size_bytes=256,
            status="pending",
        ),
        BatchFileEntry(
            file_path="/test/book.epub",
            file_name="book.epub",
            detected_format="epub",
            workflow_category="other",
            size_bytes=8192,
            status="pending",
        ),
    ]


__all__ = (
    "CATEGORY_ORDER",
    "OOXML_SIGNATURE_VALIDATION_UNAVAILABLE",
    "AdmissionDecision",
    "BatchFileEntry",
    "BatchListViewModel",
    "SimpleNamespace",
    "_add_synthetic",
    "_sort_value",
    "format_size",
    "pytest",
    "pytestmark",
    "sample_entries",
    "should_pulse_processing_transition",
    "vm",
)
