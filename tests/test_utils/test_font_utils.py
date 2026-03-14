"""utils 单元测试。"""

from __future__ import annotations

import pytest

import docwen.utils.font_utils as font_utils

pytestmark = pytest.mark.unit


@pytest.mark.unit
def test_get_available_font_exact_match(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(font_utils, "get_system_fonts", lambda: ["Arial", "SimSun"], raising=True)
    assert font_utils.get_available_font(["SimSun", "Missing"]) == "SimSun"


@pytest.mark.unit
def test_get_available_font_alias_match(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(font_utils, "get_system_fonts", lambda: ["微软雅黑"], raising=True)
    assert font_utils.get_available_font(["Microsoft YaHei"]) == "微软雅黑"


@pytest.mark.unit
def test_get_available_font_fuzzy_match(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(font_utils, "get_system_fonts", lambda: ["Helvetica Neue"], raising=True)
    assert font_utils.get_available_font(["Helvetica"]) == "Helvetica Neue"


@pytest.mark.unit
def test_get_default_font_falls_back_to_tk_default(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(font_utils, "get_available_font", lambda families: None, raising=True)
    name, size = font_utils.get_default_font()
    assert name == "TkDefaultFont"
    assert size == font_utils.DEFAULT_FONT_SIZE
