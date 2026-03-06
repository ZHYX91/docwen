"""security 单元测试。"""

from __future__ import annotations

import types

import pytest

from docwen.bootstrap import initialize_app
from docwen.errors import SecurityCheckFailedError

pytestmark = pytest.mark.unit


def test_initialize_app_raises_in_strict_mode_on_security_failure(monkeypatch: pytest.MonkeyPatch) -> None:
    m = types.ModuleType("docwen.security.protection_utils")

    def _fail():
        raise RuntimeError("boom")

    m.run_all_protections = _fail
    monkeypatch.setitem(__import__("sys").modules, "docwen.security.protection_utils", m)

    with pytest.raises(SecurityCheckFailedError):
        initialize_app(strict_security=True, return_status=True)


def test_initialize_app_degrades_in_non_strict_mode_on_security_failure(monkeypatch: pytest.MonkeyPatch) -> None:
    m = types.ModuleType("docwen.security.protection_utils")

    def _fail():
        raise RuntimeError("boom")

    m.run_all_protections = _fail
    monkeypatch.setitem(__import__("sys").modules, "docwen.security.protection_utils", m)

    ok, msg = initialize_app(strict_security=False, return_status=True)
    assert ok is False
    assert msg is not None
    assert "降级" in msg
