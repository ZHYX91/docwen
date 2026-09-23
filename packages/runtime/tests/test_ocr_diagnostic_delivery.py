"""Typed OCR diagnostic delivery keeps quality notices informational and bounded."""

from __future__ import annotations

import pytest

from docwen_core.text.ocr import OcrStatus, report_ocr_outcome
from docwen_runtime._execution_context import _RuntimeProgressSink

pytestmark = pytest.mark.unit


def test_repeated_success_is_one_notice_but_failures_keep_locations():
    sink = _RuntimeProgressSink("task")
    for index in range(3):
        report_ocr_outcome(sink, OcrStatus.SUCCESS, location=f"page-{index}")
    report_ocr_outcome(sink, OcrStatus.NO_TEXT, location="page-4")
    report_ocr_outcome(sink, OcrStatus.MODEL_MISSING, location="page-5")
    assert [(d.level, d.code, d.location) for d in sink.diagnostics] == [
        ("info", "OCR-QUALITY-NOTICE", "page-0"),
        ("warning", "OCR-BEST-EFFORT.no_text", "page-4"),
        ("warning", "OCR-BEST-EFFORT.model_missing", "page-5"),
    ]
    assert len(sink.events) == 3
