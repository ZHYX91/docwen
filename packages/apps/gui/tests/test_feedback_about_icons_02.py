"""Focused tests split from test_feedback_about_icons.py."""

from __future__ import annotations

import pytest

from ._feedback_about_icons_support import (
    Qt,
    _write_minimal_base_config_tree,
)

pytestmark = pytest.mark.gui


class TestWindowOpacity:
    """Tests for window opacity application."""

    def test_apply_window_opacity_method_exists(self, main_window) -> None:
        """MainWindow has _apply_window_opacity method."""
        assert hasattr(main_window, "_apply_window_opacity")
        assert callable(main_window._apply_window_opacity)

    def test_apply_window_opacity_no_controller_no_op(self, main_window) -> None:
        """_apply_window_opacity is a no-op when controller is None."""
        # main_window fixture has controller=None, so this should not raise
        main_window._apply_window_opacity()
        # Window should still be at default opacity (1.0)
        assert main_window.windowOpacity() >= 0.99

    def test_apply_window_opacity_with_config(self, qapp, tmp_path) -> None:
        """_apply_window_opacity reads transparency from config."""
        from PySide6.QtWidgets import QApplication

        from docwen_application.controller import ApplicationController
        from docwen_bundle.config_port import ConfigPortAdapter
        from docwen_bundle.runtime_factory import create_runtime_port
        from docwen_gui.app import create_main_window
        from docwen_gui.qt_bridge.task_event_bridge import TaskEventBridge

        bridge = TaskEventBridge()

        config_dir = tmp_path / "configs"
        config_dir.mkdir(parents=True)
        _write_minimal_base_config_tree(config_dir)

        def _event_callback(event) -> None:
            payload = {"task_id": event.task_id, **dict(event.payload)}
            bridge.enqueue(event.event_type, payload)

        runtime_port = create_runtime_port(event_callback=_event_callback)
        config_port = ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir)
        # Pre-set transparency config
        config_port.set("gui.transparency.enabled", True)
        config_port.set("gui.transparency.default_value", 0.85)

        controller = ApplicationController(
            runtime_port=runtime_port,
            config_port=config_port,
        )
        controller.start()

        window = create_main_window(controller=controller, task_event_bridge=bridge)
        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        # After setup_ui + _load_initial_preferences, opacity should be applied
        assert abs(window.windowOpacity() - 0.85) < 0.01

        # Cleanup
        from tests.support.gui import shutdown_main_window

        shutdown_main_window(window)


class TestAppIcon:
    """Tests for application icon initialization."""

    def test_initialize_app_icon_method_exists(self, main_window) -> None:
        """MainWindow has _initialize_app_icon method."""
        assert hasattr(main_window, "_initialize_app_icon")
        assert callable(main_window._initialize_app_icon)

    def test_initialize_app_icon_no_error(self, main_window) -> None:
        """_initialize_app_icon does not raise."""
        main_window._initialize_app_icon()


class TestIconResources:
    """Tests for icon resource helpers."""

    def test_get_icon_path_returns_none_when_no_assets(self, tmp_path, monkeypatch) -> None:
        """get_icon_path returns None when assets dir has no icon files."""
        from docwen_runtime.resources.registry import ResourceRegistry

        root = tmp_path / "bundle"
        (root / "assets").mkdir(parents=True)
        monkeypatch.setattr(
            ResourceRegistry,
            "default",
            classmethod(lambda cls: ResourceRegistry(root)),
        )

        from docwen_gui.resources import get_icon_path

        result = get_icon_path()
        assert result is None

    def test_get_icon_path_with_png(self, tmp_path, monkeypatch) -> None:
        """get_icon_path returns icon.png path when present."""
        from docwen_runtime.resources.registry import ResourceRegistry

        root = tmp_path / "bundle"
        (root / "assets").mkdir(parents=True)
        (root / "assets" / "icon.png").write_text("")
        monkeypatch.setattr(
            ResourceRegistry,
            "default",
            classmethod(lambda cls: ResourceRegistry(root)),
        )

        from docwen_gui.resources import get_icon_path

        result = get_icon_path()
        assert result is not None
        assert result.endswith("icon.png")

    def test_load_svg_icon_not_found(self, tmp_path, monkeypatch) -> None:
        """load_svg_icon returns None when the file does not exist."""
        from docwen_runtime.resources.registry import ResourceRegistry

        root = tmp_path / "bundle"
        (root / "assets" / "icons").mkdir(parents=True)
        monkeypatch.setattr(
            ResourceRegistry,
            "default",
            classmethod(lambda cls: ResourceRegistry(root)),
        )

        from docwen_gui.resources import load_svg_icon

        result = load_svg_icon("nonexistent.svg")
        assert result is None

    def test_load_image_icon_not_found(self, tmp_path, monkeypatch) -> None:
        """load_image_icon returns None when the file does not exist."""
        from docwen_runtime.resources.registry import ResourceRegistry

        root = tmp_path / "bundle"
        (root / "assets").mkdir(parents=True)
        monkeypatch.setattr(
            ResourceRegistry,
            "default",
            classmethod(lambda cls: ResourceRegistry(root)),
        )

        from docwen_gui.resources import load_image_icon

        result = load_image_icon("nonexistent.png")
        assert result is None

    def test_asset_path_returns_expected(self, tmp_path, monkeypatch) -> None:
        """asset_path returns correct path relative to assets dir."""
        from pathlib import Path

        from docwen_runtime.resources.registry import ResourceRegistry

        root = tmp_path / "bundle"
        (root / "assets").mkdir(parents=True)
        monkeypatch.setattr(
            ResourceRegistry,
            "default",
            classmethod(lambda cls: ResourceRegistry(root)),
        )

        from docwen_gui.resources import asset_path

        result = asset_path("icons/test.svg")
        assert isinstance(result, Path)
        assert result == root / "assets" / "icons" / "test.svg"

    def test_initialize_app_icon_no_assets(self, tmp_path, monkeypatch, qapp) -> None:
        """initialize_app_icon returns False when no icon assets exist."""
        from unittest.mock import MagicMock

        from docwen_runtime.resources.registry import ResourceRegistry

        root = tmp_path / "bundle"
        (root / "assets").mkdir(parents=True)
        monkeypatch.setattr(
            ResourceRegistry,
            "default",
            classmethod(lambda cls: ResourceRegistry(root)),
        )

        # Reset the singleton state so previous test runs don't affect this one
        from docwen_gui import resources as gui_resources

        gui_resources._IconManager._initialized = False
        gui_resources._IconManager._icon_path = None

        mock_window = MagicMock()
        mock_window.setWindowIcon = MagicMock()
        result = gui_resources.initialize_app_icon(mock_window)
        assert result is False

    def test_icon_manager_initialize_success_and_cache(self, tmp_path, monkeypatch, qapp) -> None:
        """A valid icon initializes once and keeps the resolved path cached."""
        from unittest.mock import MagicMock

        from PySide6.QtGui import QImage

        from docwen_gui import resources as gui_resources
        from docwen_runtime.resources.registry import ResourceRegistry

        root = tmp_path / "bundle"
        assets = root / "assets"
        assets.mkdir(parents=True)
        icon_path = assets / "icon.png"
        image = QImage(12, 8, QImage.Format.Format_ARGB32)
        image.fill(Qt.GlobalColor.red)
        assert image.save(str(icon_path))

        monkeypatch.setattr(
            ResourceRegistry,
            "default",
            classmethod(lambda cls: ResourceRegistry(root)),
        )
        monkeypatch.setattr(gui_resources._IconManager, "_initialized", False)
        monkeypatch.setattr(gui_resources._IconManager, "_icon_path", None)

        mock_window = MagicMock()
        assert gui_resources._IconManager.initialize(mock_window) is True
        assert gui_resources._IconManager._initialized is True
        assert gui_resources._IconManager._icon_path == str(icon_path)
        assert mock_window.setWindowIcon.call_count == 1

        assert gui_resources._IconManager.initialize(mock_window) is True
        assert gui_resources._IconManager.get_icon_path() == str(icon_path)
        assert mock_window.setWindowIcon.call_count == 1

    def test_icon_manager_get_icon_path_caches_lookup(self, tmp_path, monkeypatch) -> None:
        """Successful lookups are cached while a missing icon remains retryable."""
        from docwen_gui import resources as gui_resources

        resolved: list[str | None] = [str(tmp_path / "icon.png")]
        lookups: list[None] = []

        def fake_get_icon_path() -> str | None:
            lookups.append(None)
            return resolved[0]

        monkeypatch.setattr(gui_resources, "get_icon_path", fake_get_icon_path)
        monkeypatch.setattr(gui_resources._IconManager, "_icon_path", None)

        assert gui_resources._IconManager.get_icon_path() == resolved[0]
        assert gui_resources._IconManager.get_icon_path() == resolved[0]
        assert len(lookups) == 1

        resolved[0] = None
        gui_resources._IconManager._icon_path = None
        assert gui_resources._IconManager.get_icon_path() is None
        assert gui_resources._IconManager.get_icon_path() is None
        assert len(lookups) == 3

    def test_load_image_icon_with_valid_png_and_scaling(self, tmp_path, monkeypatch, qapp) -> None:
        """A valid PNG yields a non-null icon and an exact requested-size pixmap."""
        from PySide6.QtCore import QSize
        from PySide6.QtGui import QIcon, QImage, QPixmap

        from docwen_gui import resources as gui_resources
        from docwen_runtime.resources.registry import ResourceRegistry

        root = tmp_path / "bundle"
        assets = root / "assets"
        assets.mkdir(parents=True)
        icon_path = assets / "valid.png"
        image = QImage(12, 8, QImage.Format.Format_ARGB32)
        image.fill(Qt.GlobalColor.blue)
        assert image.save(str(icon_path))

        monkeypatch.setattr(
            ResourceRegistry,
            "default",
            classmethod(lambda cls: ResourceRegistry(root)),
        )
        monkeypatch.setattr(gui_resources, "_icon_cache", {})

        icon = gui_resources.load_image_icon("valid.png")
        assert isinstance(icon, QIcon)
        assert not icon.isNull()

        scaled = gui_resources.load_image_icon("valid.png", size=(9, 7))
        assert isinstance(scaled, QPixmap)
        assert not scaled.isNull()
        assert scaled.size() == QSize(9, 7)
        assert scaled.devicePixelRatio() >= 1.0

        cached = gui_resources.load_image_icon("valid.png", size=(9, 7))
        assert isinstance(cached, QPixmap)
        assert cached.cacheKey() == scaled.cacheKey()


class TestDialogsPackage:
    """Tests for the dialogs package __init__ exports."""

    def test_dialogs_init_exports(self, qapp) -> None:
        """dialogs.__init__ exports expected symbols."""
        from docwen_gui import dialogs

        assert hasattr(dialogs, "AboutDialog")
        assert hasattr(dialogs, "error")
        assert hasattr(dialogs, "warn")
        assert hasattr(dialogs, "info")
        assert hasattr(dialogs, "confirm")
        assert hasattr(dialogs, "choose")
        assert hasattr(dialogs, "notify")
        assert hasattr(dialogs, "exception")
        assert hasattr(dialogs, "FeedbackChoice")
        assert hasattr(dialogs, "FeedbackLevel")

    def test_dialogs_all_is_list_of_str(self, qapp) -> None:
        """dialogs.__all__ is a list of strings."""
        from docwen_gui import dialogs

        assert isinstance(dialogs.__all__, list)
        assert all(isinstance(name, str) for name in dialogs.__all__)


class TestAboutDialogStyles:
    """Tests for the about dialog stylesheet."""

    def test_build_about_dialog_stylesheet_returns_str(self) -> None:
        """build_about_dialog_stylesheet returns a non-empty string."""
        from docwen_gui.styles.about_dialog import build_about_dialog_stylesheet

        css = build_about_dialog_stylesheet()
        assert isinstance(css, str)
        assert len(css) > 0

    def test_build_about_dialog_stylesheet_includes_selectors(self) -> None:
        """Stylesheet includes key selectors for the about dialog."""
        from docwen_gui.styles.about_dialog import build_about_dialog_stylesheet

        css = build_about_dialog_stylesheet()
        assert "aboutDialog" in css
        assert "aboutTitle" in css
        assert "aboutHeroCard" in css
        assert "aboutCloseButton" in css
        assert "color: palette(text);" in css
        assert "QToolButton#aboutToolInfoButton" in css


class TestGlobalAggregateAbout:
    """Tests that the global aggregate includes about dialog styles."""

    def test_global_aggregate_includes_about_styles(self) -> None:
        """build_global_stylesheet includes about dialog CSS."""
        from docwen_gui.styles.global_aggregate import build_global_stylesheet

        css = build_global_stylesheet("light")
        assert "aboutDialog" in css
        assert "aboutTitle" in css
        assert "aboutHeroCard" in css
