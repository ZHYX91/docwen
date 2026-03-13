"""services.result_presentation 的单元测试。"""

from __future__ import annotations

import pytest

from docwen.errors import DocWenError
from docwen.services.error_codes import ERROR_CODE_UNKNOWN_ERROR
from docwen.services.result import ConversionResult
from docwen.services.result_presentation import (
    format_exception_message,
    format_result_message,
    normalize_result_error_fields,
)

pytestmark = pytest.mark.unit


def test_normalize_result_error_fields_success_is_noop() -> None:
    ok = ConversionResult.ok(message="ok")
    assert normalize_result_error_fields(ok) is ok


def test_normalize_result_error_fields_fills_from_docwenerror() -> None:
    e = DocWenError("E_TEST", "msg", details="d")
    result = ConversionResult(success=False, message=None, error=e, error_code=None, details=None)
    out = normalize_result_error_fields(result)
    assert out is not result
    assert out.error_code == "E_TEST"
    assert out.message == "msg"
    assert out.details == "d"


def test_normalize_result_error_fields_keeps_existing_fields_when_already_present() -> None:
    e = DocWenError("E_TEST", "msg", details="d")
    result = ConversionResult(success=False, message="m2", error=e, error_code="E2", details="d2")
    out = normalize_result_error_fields(result)
    assert out is result
    assert out.error_code == "E2"
    assert out.message == "m2"
    assert out.details == "d2"


def test_normalize_result_error_fields_adds_unknown_error_code_and_details_from_exception() -> None:
    e = ValueError("boom")
    result = ConversionResult(success=False, message=None, error=e, error_code=None, details=None)
    out = normalize_result_error_fields(result)
    assert out.error_code == ERROR_CODE_UNKNOWN_ERROR
    assert out.details == "boom"


def test_format_result_message_includes_details_when_present() -> None:
    result = ConversionResult(success=False, message="m", error=None, error_code=ERROR_CODE_UNKNOWN_ERROR, details="d")
    assert format_result_message(result, "default") == "m (d)"


def test_format_result_message_falls_back_to_default_message() -> None:
    result = ConversionResult(
        success=False, message=None, error=None, error_code=ERROR_CODE_UNKNOWN_ERROR, details=None
    )
    assert format_result_message(result, "default") == "default"


def test_format_exception_message_for_docwenerror_uses_user_message_and_details() -> None:
    assert format_exception_message(DocWenError("E", "m", details="d"), "default") == "m (d)"
    assert format_exception_message(DocWenError("E", "m"), "default") == "m"


def test_format_exception_message_falls_back_to_default_when_str_is_empty() -> None:
    class _EmptyStrError(Exception):
        def __str__(self) -> str:
            return ""

    assert format_exception_message(_EmptyStrError(), "default") == "default"
