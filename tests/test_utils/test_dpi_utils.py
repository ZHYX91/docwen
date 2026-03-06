"""utils 单元测试。"""

from __future__ import annotations

import pytest

from docwen.utils import dpi_utils

pytestmark = pytest.mark.unit


def test_scale_unscale_roundtrip(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(dpi_utils.DPIManager, "ENABLE_DPI_SCALING", True)
    dpi_utils.dpi_manager.scaling_factor = 1.5

    for value in [0, 1, 2, 3, 4, 5, 10, 101, 999]:
        assert dpi_utils.unscale(dpi_utils.scale(value)) == value


def test_unscale_noop_when_scaling_disabled(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(dpi_utils.DPIManager, "ENABLE_DPI_SCALING", False)
    dpi_utils.dpi_manager.scaling_factor = 2.0

    assert dpi_utils.scale(123) == 123
    assert dpi_utils.unscale(123) == 123


def test_initialize_dpi_manager_forced_scale_overrides_detected_factor(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(dpi_utils.DPIManager, "ENABLE_DPI_SCALING", True)

    def fake_detect(self: dpi_utils.DPIManager) -> None:
        self.scaling_factor = 2.0

    monkeypatch.setattr(dpi_utils.DPIManager, "_detect_scaling_factor", fake_detect)

    dpi_utils.initialize_dpi_manager(None, forced_scale=1.25)
    assert dpi_utils.get_scaling_factor() == pytest.approx(1.25)


def test_initialize_dpi_manager_forced_scale_ignored_when_disabled(monkeypatch: pytest.MonkeyPatch) -> None:
    dpi_utils.initialize_dpi_manager(None, scaling_enabled=False, forced_scale=1.5)
    assert dpi_utils.get_scaling_factor() == pytest.approx(1.0)
