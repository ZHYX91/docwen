"""config 单元测试。"""

from __future__ import annotations

import pytest

from docwen.config.schemas.gui import GUIConfigMixin

pytestmark = pytest.mark.unit


class _Dummy(GUIConfigMixin):
    def __init__(self, configs: dict) -> None:
        self._configs = configs


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        (0, 0.0),
        (1.25, 1.25),
        ("1.5", 1.5),
        ("150%", 1.5),
        ("", 0.0),
        (None, 0.0),
    ],
)
def test_get_ui_scale_parses_values(raw: object, expected: float) -> None:
    cfg = _Dummy({"gui_config": {"dpi": {"ui_scale": raw}}})
    assert cfg.get_ui_scale() == pytest.approx(expected)


def test_is_dpi_scaling_enabled_defaults_true() -> None:
    cfg = _Dummy({"gui_config": {}})
    assert cfg.is_dpi_scaling_enabled() is True


def test_is_dpi_scaling_enabled_reads_value() -> None:
    cfg = _Dummy({"gui_config": {"dpi": {"enable_dpi_scaling": False}}})
    assert cfg.is_dpi_scaling_enabled() is False
