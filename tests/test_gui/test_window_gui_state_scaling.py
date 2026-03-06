"""GUI 逻辑单元测试。"""

from __future__ import annotations

import pytest

pytest.importorskip("ttkbootstrap")

from docwen.gui.core.window import MainWindow
from docwen.utils import dpi_utils

pytestmark = [pytest.mark.unit, pytest.mark.windows_only]


class _ConfigManager:
    def __init__(self) -> None:
        self.calls: list[tuple[str, str, str, str, int]] = []

    def update_config_value(self, block: str, section: str, key: str, value: int) -> None:
        self.calls.append(("update_config_value", block, section, key, value))


class _Root:
    def __init__(self, x: int, y: int, height: int) -> None:
        self._x = x
        self._y = y
        self._height = height

    def winfo_x(self) -> int:
        return self._x

    def winfo_y(self) -> int:
        return self._y

    def winfo_height(self) -> int:
        return self._height


class _FakeMainWindow:
    def __init__(self, *, factor: float) -> None:
        self.config_manager = _ConfigManager()
        self.root = _Root(x=0, y=0, height=0)
        self.remember_gui_state = True
        self.batch_panel_visible = False
        self.left_width = 0
        self.margin = 0
        self.center_panel_screen_x = 0
        self.current_y = 0

        dpi_utils.dpi_manager.scaling_factor = factor

    def unscale(self, value: int) -> int:
        return dpi_utils.unscale(value)


def _get_saved_value(cfg: _ConfigManager, key: str) -> int:
    for name, block, section, saved_key, value in cfg.calls:
        if name == "update_config_value" and block == "gui_config" and section == "window" and saved_key == key:
            return value
    raise AssertionError(f"missing saved key: {key}")


def test_save_gui_state_writes_unscaled_logical_values(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(dpi_utils.DPIManager, "ENABLE_DPI_SCALING", True)

    fake = _FakeMainWindow(factor=1.5)
    fake.center_panel_screen_x = 1500
    fake.current_y = 300
    fake.root = _Root(x=10, y=20, height=900)

    MainWindow._save_gui_state(fake)  # type: ignore[misc]

    assert _get_saved_value(fake.config_manager, "center_panel_screen_x") == 1000
    assert _get_saved_value(fake.config_manager, "window_y") == 200
    assert _get_saved_value(fake.config_manager, "default_height") == 600


def test_save_gui_state_noop_when_remember_gui_state_disabled(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(dpi_utils.DPIManager, "ENABLE_DPI_SCALING", True)

    fake = _FakeMainWindow(factor=2.0)
    fake.remember_gui_state = False
    fake.center_panel_screen_x = 800
    fake.current_y = 200
    fake.root = _Root(x=0, y=0, height=1000)

    MainWindow._save_gui_state(fake)  # type: ignore[misc]

    assert fake.config_manager.calls == []
