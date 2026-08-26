"""GUI font selection tests."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.gui


def test_get_system_fonts_returns_cached_result_without_query(monkeypatch: pytest.MonkeyPatch) -> None:
    from PySide6.QtGui import QFontDatabase

    from docwen_gui import font_utils

    monkeypatch.setattr(font_utils, "_system_fonts_cache", ["Cached Family"])
    monkeypatch.setattr(
        QFontDatabase,
        "families",
        staticmethod(lambda: (_ for _ in ()).throw(AssertionError("font database must not be queried"))),
    )

    assert font_utils.get_system_fonts() == ["Cached Family"]


def test_get_system_fonts_queries_database_and_caches_available_family(monkeypatch: pytest.MonkeyPatch) -> None:
    from PySide6.QtGui import QFontDatabase

    from docwen_gui import font_utils

    calls = 0

    def families() -> list[str]:
        nonlocal calls
        calls += 1
        return ["Arial", "Microsoft YaHei"]

    monkeypatch.setattr(font_utils, "_system_fonts_cache", None)
    monkeypatch.setattr(font_utils, "_load_known_cjk_font_files", lambda: pytest.fail("unexpected font load"))
    monkeypatch.setattr(QFontDatabase, "families", staticmethod(families))

    assert font_utils.get_system_fonts() == ["Arial", "Microsoft YaHei"]
    assert font_utils.get_system_fonts() == ["Arial", "Microsoft YaHei"]
    assert calls == 1


def test_get_system_fonts_loads_known_cjk_fonts_then_requeries(monkeypatch: pytest.MonkeyPatch) -> None:
    from PySide6.QtGui import QFontDatabase

    from docwen_gui import font_utils

    query_results = iter((["Comic Sans MS"], ["Comic Sans MS", "Noto Sans CJK SC"]))
    load_calls: list[bool] = []
    monkeypatch.setattr(font_utils, "_system_fonts_cache", None)
    monkeypatch.setattr(font_utils, "_load_known_cjk_font_files", lambda: load_calls.append(True))
    monkeypatch.setattr(QFontDatabase, "families", staticmethod(lambda: next(query_results)))

    assert font_utils.get_system_fonts() == ["Comic Sans MS", "Noto Sans CJK SC"]
    assert load_calls == [True]


def test_get_system_fonts_fails_soft_and_caches_empty_result(monkeypatch: pytest.MonkeyPatch) -> None:
    from PySide6.QtGui import QFontDatabase

    from docwen_gui import font_utils

    calls = 0

    def fail_query() -> list[str]:
        nonlocal calls
        calls += 1
        raise RuntimeError("font database unavailable")

    monkeypatch.setattr(font_utils, "_system_fonts_cache", None)
    monkeypatch.setattr(QFontDatabase, "families", staticmethod(fail_query))

    assert font_utils.get_system_fonts() == []
    assert font_utils.get_system_fonts() == []
    assert calls == 1


def test_get_available_font_uses_old_project_cjk_priority(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_gui import font_utils

    monkeypatch.setattr(
        font_utils,
        "get_system_fonts",
        lambda: ["Arial", "Segoe UI", "Microsoft YaHei"],
    )

    assert font_utils.get_available_font() == "Microsoft YaHei"


def test_get_available_font_accepts_alias(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_gui import font_utils

    monkeypatch.setattr(font_utils, "get_system_fonts", lambda: ["微软雅黑", "Arial"])

    assert font_utils.get_available_font(["Microsoft YaHei", "Arial"]) == "微软雅黑"


def test_apply_application_font_preserves_cjk_family_in_offscreen_qapp(qapp, monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_gui import font_utils

    monkeypatch.setattr(font_utils, "get_available_font", lambda font_families=None: "Arial")

    font_utils.apply_application_font(qapp, font_size_preset="large")

    assert qapp.font().family() == "Arial"
    assert qapp.font().pointSize() == 13


def test_font_size_preset_contract_matches_old_projects() -> None:
    from docwen_gui import font_utils

    assert font_utils.FONT_SIZE_PRESETS == {
        "small": 11,
        "default": 12,
        "large": 13,
        "xlarge": 15,
    }
    assert font_utils.resolve_font_size_preset("extra-large") == 15
    assert font_utils.resolve_font_size_preset("unknown") == 12
