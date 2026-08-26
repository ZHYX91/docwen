"""Smoke tests for BatchList widget.

Tests validate widget construction, 6 category tabs, entry card structure,
filter/sort popup menus, Ctrl+Up/Down reorder, context menu, and compact
mode behavior.  Require a QApplication instance.

These tests focus on user-visible behavior, NOT private widget attributes.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Iterator
from pathlib import Path
from typing import Protocol, cast

import pytest
from PySide6.QtCore import QEvent, Qt
from PySide6.QtGui import QKeyEvent
from PySide6.QtWidgets import QApplication, QBoxLayout, QListWidget, QListWidgetItem, QPushButton

from docwen_core.detection import SUPPORTED_EXTENSION_FORMATS
from docwen_core.formats.categories import get_category
from docwen_gui.i18n import t as _t
from docwen_gui.styles._hex_helper import _hex_to_rgba
from docwen_gui.styles.batch_list import build_batch_list_stylesheet
from docwen_gui.styles.theme_manager import ThemeManager
from docwen_gui.styles.theme_semantics import COLOR_SECONDARY, get_status_color
from docwen_gui.view_models.batch_list_vm import (
    CATEGORY_ORDER,
    FILTER_OPTIONS,
    BatchFileEntry,
    BatchListViewModel,
)
from docwen_gui.view_models.input_area_vm import _BATCH_SCAN_LIMIT
from docwen_gui.widgets.batch_list import (
    BatchEntryItemWidget,
    BatchList,
    ReorderableListWidget,
    WrapRowLayout,
    _filter_option_label,
    _format_size,
    _load_status_icon,
    _source_path_text,
)

pytestmark = pytest.mark.gui


def _dominant_opaque_pixmap_color(label) -> str | None:
    pixmap = label.pixmap()
    image = pixmap.toImage()
    colors = Counter(
        image.pixelColor(x, y).name().upper()
        for y in range(image.height())
        for x in range(image.width())
        if image.pixelColor(x, y).alpha() > 0
    )
    return colors.most_common(1)[0][0] if colors else None


class _PivotText(Protocol):
    def text(self) -> str: ...


def _synthetic_file_resolver(path: str) -> dict[str, str] | None:
    fmt = SUPPORTED_EXTENSION_FORMATS.get(Path(path).suffix.lower())
    if fmt is None:
        return None
    category = "markdown" if fmt in {"markdown", "txt"} else get_category(fmt)
    return {"detected_format": fmt, "workflow_category": category}


def _add_synthetic(vm: BatchListViewModel, paths: list[str]) -> tuple[list[str], list[tuple[str, str]]]:
    return vm.add_files(paths, file_resolver=_synthetic_file_resolver)


@pytest.fixture
def vm() -> BatchListViewModel:
    return BatchListViewModel()


@pytest.fixture
def widget(qapp: QApplication, vm: BatchListViewModel) -> Iterator[BatchList]:
    w = BatchList(view_model=vm)
    yield w
    w.deleteLater()


@pytest.fixture
def populated_widget(qapp: QApplication, vm: BatchListViewModel) -> Iterator[BatchList]:
    w = BatchList(view_model=vm)
    _add_synthetic(
        vm,
        [
            "/test/doc.docx",
            "/test/sheet.xlsx",
            "/test/img.png",
            "/test/layout.pdf",
            "/test/readme.md",
            "/test/book.epub",
        ],
    )
    yield w
    w.deleteLater()


__all__ = (
    "CATEGORY_ORDER",
    "COLOR_SECONDARY",
    "FILTER_OPTIONS",
    "_BATCH_SCAN_LIMIT",
    "BatchEntryItemWidget",
    "BatchFileEntry",
    "BatchList",
    "BatchListViewModel",
    "Path",
    "QApplication",
    "QBoxLayout",
    "QEvent",
    "QKeyEvent",
    "QListWidget",
    "QListWidgetItem",
    "QPushButton",
    "Qt",
    "ReorderableListWidget",
    "ThemeManager",
    "WrapRowLayout",
    "_PivotText",
    "_add_synthetic",
    "_dominant_opaque_pixmap_color",
    "_filter_option_label",
    "_format_size",
    "_hex_to_rgba",
    "_load_status_icon",
    "_source_path_text",
    "_t",
    "build_batch_list_stylesheet",
    "cast",
    "get_status_color",
    "populated_widget",
    "pytest",
    "pytestmark",
    "vm",
    "widget",
)
