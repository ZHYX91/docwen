"""User-facing OCR state comes from diagnostic codes, never translated prose."""

from __future__ import annotations

import pytest

from docwen_core.models.result import ConversionDiagnostic, ConversionResult
from docwen_gui.execution_presenter import _result_warning_messages
from docwen_gui.i18n import get_locale, set_locale

pytestmark = pytest.mark.unit


def test_warning_localization_does_not_parse_english_message():
    previous = get_locale()
    try:
        set_locale("zh_CN")
        result = ConversionResult(
            task_id="t",
            success=True,
            diagnostics=[
                ConversionDiagnostic("warning", "没有英文状态句子", code="OCR-BEST-EFFORT.no_text", location="page-2"),
                ConversionDiagnostic("info", "untranslated", code="OCR-QUALITY-NOTICE"),
            ],
        )
        messages = _result_warning_messages(result)
        assert len(messages) == 1
        assert "page-2" in messages[0]
        assert "没有英文状态句子" not in messages[0]
        assert "OCR" in messages[0]
    finally:
        set_locale(previous)
