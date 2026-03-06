"""GUI 逻辑单元测试。"""

from __future__ import annotations

import types

import pytest

pytest.importorskip("ttkbootstrap")

from docwen.gui.settings.dialog import SettingsDialog

pytestmark = [pytest.mark.unit, pytest.mark.windows_only]


class _Tab:
    def __init__(self, *, ok: bool = True) -> None:
        self.ok = ok
        self.called = 0

    def apply_settings(self) -> bool:
        self.called += 1
        return self.ok


def test_apply_all_settings_includes_export_tab() -> None:
    export_tab = _Tab(ok=True)
    fake = types.SimpleNamespace(tabs={"export": export_tab})

    success = SettingsDialog._apply_all_settings(fake)  # type: ignore[misc]

    assert success is True
    assert export_tab.called == 1


def test_apply_all_settings_fails_when_export_tab_fails() -> None:
    export_tab = _Tab(ok=False)
    fake = types.SimpleNamespace(tabs={"export": export_tab})

    success = SettingsDialog._apply_all_settings(fake)  # type: ignore[misc]

    assert success is False
    assert export_tab.called == 1

