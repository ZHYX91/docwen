"""ThemeManager behavior tests."""

from __future__ import annotations

import pytest

from docwen_gui.styles.theme_manager import ThemeManager

pytestmark = pytest.mark.unit


class _FakeStyle:
    def __init__(self) -> None:
        self.unpolished = 0
        self.polished = 0

    def unpolish(self, widget: object) -> None:
        self.unpolished += 1

    def polish(self, widget: object) -> None:
        self.polished += 1


class _FakeWidget:
    def __init__(self, *, visible: bool, children: list[_FakeWidget] | None = None) -> None:
        self._visible = visible
        self._style = _FakeStyle()
        self._children = children or []
        self.updated = 0

    def isVisible(self) -> bool:
        return self._visible

    def style(self) -> _FakeStyle:
        return self._style

    def findChildren(self, widget_type: object) -> list[_FakeWidget]:
        return self._children

    def update(self) -> None:
        self.updated += 1


class _FakeApp:
    def __init__(self, widgets: list[_FakeWidget]) -> None:
        self._widgets = widgets
        self.palette_set = False
        self.stylesheet = ""

    def setPalette(self, palette: object) -> None:
        self.palette_set = True

    def setStyleSheet(self, stylesheet: str) -> None:
        self.stylesheet = stylesheet

    def topLevelWidgets(self) -> list[_FakeWidget]:
        return self._widgets


def test_apply_theme_repolishes_visible_top_level_widgets_only(monkeypatch: pytest.MonkeyPatch) -> None:
    visible_child = _FakeWidget(visible=True)
    visible = _FakeWidget(visible=True, children=[visible_child])
    hidden = _FakeWidget(visible=False)
    app = _FakeApp([visible, hidden])
    manager = ThemeManager()
    monkeypatch.setattr(manager, "_app", app)

    monkeypatch.setattr("qfluentwidgets.setTheme", lambda theme: None)

    manager.apply_theme("dark")

    assert app.palette_set is True
    assert app.stylesheet
    assert visible.style().unpolished == 1
    assert visible.style().polished == 1
    assert visible.updated == 1
    assert visible_child.updated == 1
    assert hidden.style().unpolished == 0
    assert hidden.style().polished == 0
    assert hidden.updated == 0


def test_font_size_getter_defaults_and_tracks_normalized_writes(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_gui.styles import theme_manager

    class _NoApplication:
        @staticmethod
        def instance() -> None:
            return None

    monkeypatch.setattr(theme_manager, "QApplication", _NoApplication)
    manager = ThemeManager()

    assert manager.get_font_size_preset() == "default"
    for raw, expected in (
        (None, "default"),
        ("", "default"),
        (" unsupported ", "default"),
        (" LARGE ", "large"),
        ("xlarge", "xlarge"),
    ):
        assert manager.apply_font_size_preset(raw) == expected
        assert manager.get_font_size_preset() == expected
