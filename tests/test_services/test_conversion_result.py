"""services 单元测试。"""

from __future__ import annotations

import pytest

from docwen.services.error_codes import ERROR_CODE_UNKNOWN_ERROR
from docwen.services.result import ConversionResult
from docwen.services.result_presentation import normalize_result_error_fields

pytestmark = pytest.mark.unit


def test_conversion_result_fail_rejects_non_exception() -> None:
    with pytest.raises(TypeError):
        ConversionResult.fail("x", error="boom", error_code=ERROR_CODE_UNKNOWN_ERROR)  # type: ignore[arg-type]


def test_conversion_result_fail_requires_error_code() -> None:
    with pytest.raises(ValueError):
        ConversionResult.fail("x", error_code="")


def test_conversion_result_fail_accepts_exception() -> None:
    e = RuntimeError("x")
    r = ConversionResult.fail("x", error=e, error_code=ERROR_CODE_UNKNOWN_ERROR, details="d")
    assert r.success is False
    assert r.error is e
    assert r.error_code == ERROR_CODE_UNKNOWN_ERROR
    assert r.details == "d"


def test_normalize_result_error_fields_adds_unknown_error_code_when_missing() -> None:
    r = ConversionResult(success=False, message="x")
    normalized = normalize_result_error_fields(r)
    assert normalized.error_code == ERROR_CODE_UNKNOWN_ERROR
