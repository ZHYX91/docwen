"""docwen.errors 的单元测试。"""

from __future__ import annotations

import pytest

from docwen.errors import (
    DependencyMissingError,
    ExitCode,
    StrategyNotFoundError,
    exit_code_from_error_code,
)

pytestmark = pytest.mark.unit


def test_dependency_missing_error_string_includes_details() -> None:
    assert str(DependencyMissingError("m")) == "m"
    assert str(DependencyMissingError("m", details="d")) == "m (d)"


def test_strategy_not_found_error_details_variants() -> None:
    e1 = StrategyNotFoundError()
    assert e1.details is None

    e2 = StrategyNotFoundError(action_type="convert")
    assert e2.details == "action='convert'"

    e3 = StrategyNotFoundError(action_type="convert", source_format="a", target_format="b")
    assert e3.details == "action='convert' conversion='a->b'"


def test_exit_code_from_error_code_returns_unknown_when_registry_exit_code_is_invalid(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class _Def:
        exit_code = "not-an-int"

    monkeypatch.setattr("docwen.services.error_registry.get_error_definition", lambda _c: _Def(), raising=True)
    assert exit_code_from_error_code("E_ANY") == ExitCode.UNKNOWN_ERROR
