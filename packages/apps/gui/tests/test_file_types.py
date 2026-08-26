from __future__ import annotations

from itertools import chain

import pytest

from docwen_core.detection import SUPPORTED_EXTENSION_FORMATS
from docwen_gui.file_types import FILE_CATEGORY_ORDER, FILE_EXTENSIONS_BY_CATEGORY, is_supported_file

pytestmark = pytest.mark.contract


def test_gui_groups_are_a_lossless_projection_of_core_registry() -> None:
    grouped_extensions = list(chain.from_iterable(FILE_EXTENSIONS_BY_CATEGORY.values()))

    assert tuple(FILE_EXTENSIONS_BY_CATEGORY) == FILE_CATEGORY_ORDER
    assert len(grouped_extensions) == len(set(grouped_extensions))
    assert set(grouped_extensions) == set(SUPPORTED_EXTENSION_FORMATS)


def test_markdown_alias_is_present_in_text_picker_group() -> None:
    assert ".markdown" in FILE_EXTENSIONS_BY_CATEGORY["text"]
    assert is_supported_file("README.markdown") is True
