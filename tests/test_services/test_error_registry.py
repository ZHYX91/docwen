"""services 单元测试。"""

from __future__ import annotations

import pytest

from docwen.errors import ExitCode, exit_code_from_error_code
from docwen.services.error_codes import (
    ERROR_CODE_DEPENDENCY_MISSING,
    ERROR_CODE_INVALID_INPUT,
    ERROR_CODE_STRATEGY_NOT_FOUND,
    KNOWN_ERROR_CODES,
)
from docwen.services.error_registry import ERROR_REGISTRY

pytestmark = pytest.mark.unit


def test_error_registry_covers_known_error_codes() -> None:
    missing = sorted(KNOWN_ERROR_CODES - set(ERROR_REGISTRY.keys()))
    assert missing == []


def test_exit_code_from_error_code_uses_registry_mapping() -> None:
    assert exit_code_from_error_code(ERROR_CODE_INVALID_INPUT) == ExitCode.INVALID_INPUT
    assert exit_code_from_error_code(ERROR_CODE_STRATEGY_NOT_FOUND) == ExitCode.STRATEGY_NOT_FOUND
    assert exit_code_from_error_code(ERROR_CODE_DEPENDENCY_MISSING) == ExitCode.DEPENDENCY_MISSING

