"""GUI 逻辑单元测试。"""

from __future__ import annotations

import pytest

from docwen.errors import StrategyNotFoundError
from docwen.gui.core.logic import MainWindowLogic
from docwen.services.result import ConversionResult

pytestmark = [pytest.mark.unit, pytest.mark.windows_only]


class _MainWindow:
    pass


def test_gui_format_result_message_prefers_docwenerror_over_generic_message() -> None:
    instance = MainWindowLogic(_MainWindow())
    e = StrategyNotFoundError(action_type="convert", source_format="docx", target_format="pdf")
    result = ConversionResult(success=False, message="generic", error=e, error_code=None)
    assert instance._format_result_message(result, "default") == str(e)


def test_gui_format_result_message_falls_back_to_details_and_default() -> None:
    instance = MainWindowLogic(_MainWindow())
    result = ConversionResult(success=False, message=None, details="d", error_code=None)
    assert instance._format_result_message(result, "default") == "d"

    result2 = ConversionResult(success=False, message=None, details=None, error_code=None)
    assert instance._format_result_message(result2, "default") == "default"
