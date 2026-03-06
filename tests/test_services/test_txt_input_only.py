"""services 单元测试。"""

from __future__ import annotations

import pytest

from docwen.errors import StrategyNotFoundError
from docwen.services.strategies import get_strategy
from docwen.services.strategies.markdown import MdToDocxStrategy, MdToXlsxStrategy

pytestmark = pytest.mark.unit


def test_txt_is_markdown_source_for_document_and_spreadsheet() -> None:
    assert get_strategy(source_format="txt", target_format="docx") is MdToDocxStrategy
    assert get_strategy(source_format="txt", target_format="xlsx") is MdToXlsxStrategy


@pytest.mark.parametrize(
    ("source_format", "target_format"),
    [
        ("docx", "txt"),
        ("xlsx", "txt"),
        ("pdf", "txt"),
        ("png", "txt"),
    ],
)
def test_no_strategy_outputs_txt(source_format: str, target_format: str) -> None:
    with pytest.raises(StrategyNotFoundError):
        get_strategy(source_format=source_format, target_format=target_format)
