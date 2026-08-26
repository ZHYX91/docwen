"""Focused tests split from test_main_window_features.py."""

from __future__ import annotations

import pytest

from ._main_window_features_support import (
    cast,
)

pytestmark = pytest.mark.gui


class TestShortcutsRegistered:
    """Verify that all shortcut key sequences are registered."""

    # Shortcuts that remain as QShortcut (not yet migrated to QAction)
    REQUIRED_SHORTCUT_KEYS = {  # noqa: RUF012
        "Ctrl+Shift+T",
        "Ctrl+Shift+O",
        "Ctrl+L",
        "Del",
    }

    # Shortcuts now handled by centralized QAction (see main_window._create_actions)
    ACTION_SHORTCUT_KEYS = {  # noqa: RUF012
        "Ctrl+O",
        "Ctrl+,",
        "Esc",
    }

    # Enter / Return may map to the same string on some platforms
    ENTER_EQUIVALENTS = {"Return", "Enter"}  # noqa: RUF012

    def test_core_shortcuts_registered(self, main_window) -> None:
        """All expected shortcuts are registered (QShortcut or QAction)."""
        from PySide6.QtGui import QAction, QShortcut

        shortcuts: list[QShortcut] = main_window.findChildren(QShortcut)
        registered_keys: set[str] = set()
        for sc in shortcuts:
            key_seq = sc.key()
            if not key_seq.isEmpty():
                registered_keys.add(key_seq.toString())

        # Also collect shortcuts from QActions
        actions: list[QAction] = main_window.findChildren(QAction)
        for action in actions:
            shortcut = action.shortcut()
            if not shortcut.isEmpty():
                registered_keys.add(shortcut.toString())

        for expected in self.REQUIRED_SHORTCUT_KEYS:
            assert expected in registered_keys, f"Missing QShortcut: {expected}"

        for expected in self.ACTION_SHORTCUT_KEYS:
            assert expected in registered_keys, f"Missing QAction shortcut: {expected}. Got: {registered_keys}"

    def test_enter_or_return_registered(self, main_window) -> None:
        from PySide6.QtGui import QAction, QShortcut

        shortcuts: list[QShortcut] = main_window.findChildren(QShortcut)
        registered_keys: set[str] = set()
        for sc in shortcuts:
            key_seq = sc.key()
            if not key_seq.isEmpty():
                registered_keys.add(key_seq.toString())

        # Also collect from QActions (Ctrl+Return for convert action)
        actions: list[QAction] = main_window.findChildren(QAction)
        for action in actions:
            shortcut = action.shortcut()
            if not shortcut.isEmpty():
                registered_keys.add(shortcut.toString())

        # At least one of Enter/Return must be registered; both is fine too
        overlap = registered_keys & self.ENTER_EQUIVALENTS
        assert len(overlap) >= 1, f"Neither Enter nor Return registered. Got: {registered_keys}"

    def test_total_shortcut_count_at_least_8(self, main_window) -> None:
        """Total shortcuts (QShortcut + QAction) should be >= 8."""
        from PySide6.QtGui import QAction, QShortcut

        shortcuts: list[QShortcut] = main_window.findChildren(QShortcut)
        actions: list[QAction] = main_window.findChildren(QAction)
        # Count QActions that have a non-empty shortcut
        action_shortcut_count = sum(1 for a in actions if not a.shortcut().isEmpty())
        total = len(shortcuts) + action_shortcut_count
        assert total >= 8, (
            f"Expected >=8 total shortcuts (QShortcut + QAction), "
            f"got {len(shortcuts)} QShortcut + {action_shortcut_count} QAction = {total}"
        )

    def test_shortcuts_use_window_context(self, main_window) -> None:
        from PySide6.QtCore import Qt
        from PySide6.QtGui import QShortcut

        shortcuts: list[QShortcut] = main_window.findChildren(QShortcut)
        for sc in shortcuts:
            assert sc.context() == Qt.ShortcutContext.WindowShortcut, (
                f"Shortcut {sc.key().toString()} has wrong context: {sc.context()}"
            )

    def test_shortcuts_have_auto_repeat_disabled(self, main_window) -> None:
        from PySide6.QtGui import QShortcut

        shortcuts: list[QShortcut] = main_window.findChildren(QShortcut)
        for sc in shortcuts:
            assert sc.autoRepeat() is False, f"Shortcut {sc.key().toString()} has autoRepeat enabled"


class TestFocusProtection:
    """Verify the _has_editable_text_focus() static method.

    Uses a standalone visible window to ensure focus reliably works.
    """

    @staticmethod
    def _with_focused_widget(widget_class, *args, **kwargs):
        """Create a visible window, put ``widget_class`` inside, focus it,
        run ``_has_editable_text_focus()``, and clean up.

        Returns the boolean result from ``_has_editable_text_focus()``.
        """
        from PySide6.QtCore import Qt
        from PySide6.QtWidgets import QApplication, QVBoxLayout, QWidget

        from docwen_gui.main_window import MainWindow

        app = QApplication.instance()
        win = QWidget()
        win.setWindowFlags(win.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        win.move(-20000, -20000)
        layout = QVBoxLayout(win)
        widget = widget_class(*args, parent=win, **kwargs)
        layout.addWidget(widget)
        win.show()
        win.raise_()
        if app is not None:
            app.processEvents()
        widget.setFocus()
        if app is not None:
            app.processEvents()
        result = MainWindow._has_editable_text_focus()
        win.hide()
        win.close()
        return result

    def test_no_focus_returns_false(self, main_window) -> None:
        # Clear focus first — focus a non-editable widget on main_window
        from PySide6.QtWidgets import QApplication

        main_window.move(-20000, -20000)
        main_window.show()
        main_window.raise_()
        app = QApplication.instance()
        if app is not None:
            app.processEvents()
        main_window._font_size_btn.setFocus()
        if app is not None:
            app.processEvents()
        assert main_window._has_editable_text_focus() is False
        main_window.hide()

    def test_line_edit_focus_returns_true(self, main_window) -> None:
        from PySide6.QtWidgets import QLineEdit

        assert self._with_focused_widget(QLineEdit) is True

    def test_text_edit_focus_returns_true(self, main_window) -> None:
        from PySide6.QtWidgets import QTextEdit

        assert self._with_focused_widget(QTextEdit) is True

    def test_plain_text_edit_focus_returns_true(self, main_window) -> None:
        from PySide6.QtWidgets import QPlainTextEdit

        assert self._with_focused_widget(QPlainTextEdit) is True

    def test_spin_box_focus_returns_true(self, main_window) -> None:
        from PySide6.QtWidgets import QSpinBox

        assert self._with_focused_widget(QSpinBox) is True

    def test_non_editable_combobox_focus_returns_false(self, main_window) -> None:
        from PySide6.QtWidgets import QComboBox

        assert self._with_focused_widget(QComboBox) is False

    def test_editable_combobox_focus_returns_true(self, main_window) -> None:
        from PySide6.QtCore import Qt
        from PySide6.QtWidgets import QApplication, QComboBox, QVBoxLayout, QWidget

        from docwen_gui.main_window import MainWindow

        app = QApplication.instance()
        win = QWidget()
        win.setWindowFlags(win.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        win.move(-20000, -20000)
        layout = QVBoxLayout(win)
        widget = QComboBox(win)
        widget.setEditable(True)
        layout.addWidget(widget)
        win.show()
        win.raise_()
        if app is not None:
            app.processEvents()
        widget.setFocus()
        if app is not None:
            app.processEvents()
        result = MainWindow._has_editable_text_focus()
        win.hide()
        win.close()
        assert result is True

    def test_push_button_focus_returns_false(self, main_window) -> None:
        from PySide6.QtWidgets import QPushButton

        assert self._with_focused_widget(QPushButton, "test") is False

    def test_tool_button_focus_returns_false(self, main_window) -> None:
        from PySide6.QtWidgets import QApplication

        main_window.move(-20000, -20000)
        main_window.show()
        main_window.raise_()
        app = QApplication.instance()
        if app is not None:
            app.processEvents()
        main_window._font_size_btn.setFocus()
        if app is not None:
            app.processEvents()
        assert main_window._has_editable_text_focus() is False
        main_window.hide()


class TestShortcutHandlers:
    """Smoke tests that shortcut handlers exist and are callable."""

    def test_add_file_handler_callable(self, main_window) -> None:
        assert callable(main_window._on_add_file_shortcut)

    def test_add_folder_handler_callable(self, main_window) -> None:
        assert callable(main_window._on_add_folder_shortcut)

    def test_locate_output_handler_callable(self, main_window) -> None:
        assert callable(main_window._on_locate_output_shortcut)

    def test_remove_handler_callable(self, main_window) -> None:
        assert callable(main_window._on_remove_selected_shortcut)

    def test_esc_handler_callable(self, main_window) -> None:
        assert callable(main_window._on_esc_shortcut)

    def test_toggle_top_handler_callable(self, main_window) -> None:
        assert callable(main_window._on_toggle_always_on_top_shortcut)

    def test_primary_handler_callable(self, main_window) -> None:
        assert callable(main_window._on_trigger_primary_shortcut)

    def test_add_file_handler_checks_focus_guard(self, main_window) -> None:
        """Verify the handler method exists and checks _has_editable_text_focus.

        We do NOT call the handler (it opens a file dialog), but we verify
        the guard function is used in the handler source pattern.
        """
        import inspect

        source = inspect.getsource(main_window._on_add_file_shortcut)
        assert "_has_editable_text_focus" in source, "Handler must check _has_editable_text_focus before acting"


class TestFontSizePresets:
    """Verify font size preset definitions and application."""

    def test_presets_defined(self, main_window) -> None:
        assert isinstance(main_window._FONT_SIZE_PRESETS, dict)
        assert "small" in main_window._FONT_SIZE_PRESETS
        assert "default" in main_window._FONT_SIZE_PRESETS
        assert "large" in main_window._FONT_SIZE_PRESETS
        assert "xlarge" in main_window._FONT_SIZE_PRESETS

    def test_preset_sizes_are_valid(self, main_window) -> None:
        for _preset, size in main_window._FONT_SIZE_PRESETS.items():
            assert isinstance(size, int)
            assert size > 0

    def test_preset_sizes_match_declared_design_scale(self, main_window) -> None:
        """Font presets keep the declared 11/12/13/15 design scale."""
        assert main_window._FONT_SIZE_PRESETS["small"] == 11
        assert main_window._FONT_SIZE_PRESETS["default"] == 12
        assert main_window._FONT_SIZE_PRESETS["large"] == 13
        assert main_window._FONT_SIZE_PRESETS["xlarge"] == 15

    def test_preset_labels_defined(self, main_window) -> None:
        assert isinstance(main_window._FONT_PRESET_LABELS, dict)
        for preset in ("small", "default", "large", "xlarge"):
            assert preset in main_window._FONT_PRESET_LABELS

    def test_default_preset_is_default(self, main_window) -> None:
        assert main_window._font_size_preset == "default"

    def test_apply_small_preset(self, main_window) -> None:
        main_window._apply_font_size_preset("small")
        assert main_window._font_size_preset == "small"

    def test_apply_default_preset(self, main_window) -> None:
        main_window._apply_font_size_preset("small")
        main_window._apply_font_size_preset("default")
        assert main_window._font_size_preset == "default"

    def test_apply_large_preset(self, main_window) -> None:
        main_window._apply_font_size_preset("large")
        assert main_window._font_size_preset == "large"

    def test_apply_xlarge_preset(self, main_window) -> None:
        main_window._apply_font_size_preset("xlarge")
        assert main_window._font_size_preset == "xlarge"

    def test_invalid_preset_falls_back_to_default(self, main_window) -> None:
        main_window._apply_font_size_preset("invalid")
        assert main_window._font_size_preset == "default"

    def test_empty_preset_falls_back_to_default(self, main_window) -> None:
        main_window._apply_font_size_preset("")
        assert main_window._font_size_preset == "default"

    def test_none_preset_falls_back_to_default(self, main_window) -> None:
        main_window._apply_font_size_preset(None)  # type: ignore[arg-type]
        assert main_window._font_size_preset == "default"

    def test_font_size_button_connected(self, main_window) -> None:
        """Verify _show_font_size_menu is callable and button exists."""
        assert main_window._font_size_btn is not None
        assert main_window._font_size_btn.text() == ""
        assert not main_window._font_size_btn.icon().isNull()
        assert callable(main_window._show_font_size_menu)

    def test_bottom_bar_buttons_use_icons(self, main_window) -> None:
        """Bottom utility actions should not render as placeholder text."""
        buttons = [
            main_window._font_size_btn,
            main_window._about_btn,
            main_window._settings_btn,
        ]

        for button in buttons:
            assert button.text() == ""
            assert not button.icon().isNull()
            assert button.toolTip()
            assert button.accessibleName() == button.toolTip()

    def test_bottom_bar_uses_three_clear_utility_zones(self, main_window) -> None:
        from PySide6.QtWidgets import QGridLayout, QWidget

        bar = main_window.findChild(QWidget, "bottomBar")
        assert bar is not None
        layout = bar.layout()
        assert isinstance(layout, QGridLayout)

        left_actions = main_window.findChild(QWidget, "bottomBarLeftActions")
        right_actions = main_window.findChild(QWidget, "bottomBarRightActions")
        assert left_actions is not None
        assert right_actions is not None
        left_position = cast(tuple[int, int, int, int], layout.getItemPosition(layout.indexOf(left_actions)))
        center_position = cast(
            tuple[int, int, int, int], layout.getItemPosition(layout.indexOf(main_window._version_label))
        )
        right_position = cast(tuple[int, int, int, int], layout.getItemPosition(layout.indexOf(right_actions)))
        assert left_position[:2] == (0, 0)
        assert center_position[:2] == (0, 1)
        assert right_position[:2] == (0, 2)

    def test_bottom_bar_icon_buttons_are_transparent(self) -> None:
        from docwen_gui.styles.main_window import build_main_window_stylesheet

        stylesheet = build_main_window_stylesheet()

        assert "QToolButton#fontSizeButton" in stylesheet
        assert "QToolButton#aboutButton" in stylesheet
        assert "QToolButton#settingsButton" in stylesheet
        assert "background: transparent;" in stylesheet
        assert "QToolButton#fontSizeButton:hover" in stylesheet
        assert "QToolButton#fontSizeButton:focus" in stylesheet

    def test_show_font_size_menu_structure(self, main_window) -> None:
        """Verify _show_font_size_menu is callable and references
        the expected Qt types (structural check; exec is not called
        because QMenu.exec() blocks the test runner)."""
        import inspect

        assert callable(main_window._show_font_size_menu)
        source = inspect.getsource(main_window._show_font_size_menu)
        assert "QMenu" in source
        assert "QActionGroup" in source
        assert "_FONT_PRESET_LABELS" in source

    def test_initial_font_preset_reads_config_port(self, qapp) -> None:
        from docwen_gui.main_window import MainWindow
        from docwen_gui.view_models.main_window_vm import MainWindowViewModel

        class _Cfg:
            def __init__(self) -> None:
                self.calls: list[tuple[str, object]] = []

            def get(self, key: str, default: object = None) -> object:
                return "large" if key == "gui.font.size_preset" else default

            def set(self, key: str, value: object) -> bool:
                self.calls.append((key, value))
                return True

        class _Controller:
            def __init__(self) -> None:
                self.config_port = _Cfg()

            def stop(self) -> None:
                return None

        vm = MainWindowViewModel(controller=_Controller())  # type: ignore[arg-type]
        window = MainWindow(view_model=vm)
        window.setup_ui()
        try:
            assert window._font_size_preset == "large"
        finally:
            window.close()

    def test_font_preset_persists_via_config_port(self, qapp) -> None:
        from docwen_gui.main_window import MainWindow
        from docwen_gui.view_models.main_window_vm import MainWindowViewModel

        class _Cfg:
            def __init__(self) -> None:
                self.calls: list[tuple[str, object]] = []

            def get(self, key: str, default: object = None) -> object:
                return default

            def set(self, key: str, value: object) -> bool:
                self.calls.append((key, value))
                return True

        class _Controller:
            def __init__(self) -> None:
                self.config_port = _Cfg()

            def stop(self) -> None:
                return None

        ctrl = _Controller()
        vm = MainWindowViewModel(controller=ctrl)  # type: ignore[arg-type]
        window = MainWindow(view_model=vm)
        window.setup_ui()
        try:
            window._apply_font_size_preset("small")
            assert ("gui.font.size_preset", "small") in ctrl.config_port.calls
        finally:
            window.close()


class TestInputAreaPublicAPI:
    """Verify the new public methods expose file/folder dialogs without
    calling private _prefixed methods."""

    def test_public_open_file_dialog_exists(self, main_window) -> None:
        assert hasattr(main_window._input_area, "open_file_dialog")
        assert callable(main_window._input_area.open_file_dialog)

    def test_public_open_folder_dialog_exists(self, main_window) -> None:
        assert hasattr(main_window._input_area, "open_folder_dialog")
        assert callable(main_window._input_area.open_folder_dialog)

    def test_shortcut_handlers_call_public_api(self, main_window) -> None:
        import inspect

        source = inspect.getsource(main_window._on_add_file_shortcut)
        assert "open_file_dialog" in source
        assert "_open_file_dialog" not in source

        source = inspect.getsource(main_window._on_add_folder_shortcut)
        assert "open_folder_dialog" in source
        assert "_open_folder_dialog" not in source
        assert "force_batch_mode=True" in source

    def test_add_folder_shortcut_forces_batch_mode(self, main_window, monkeypatch) -> None:
        calls: list[bool] = []
        monkeypatch.setattr(main_window, "_has_editable_text_focus", lambda: False)
        monkeypatch.setattr(
            main_window._input_area,
            "open_folder_dialog",
            lambda *, force_batch_mode=False: calls.append(force_batch_mode),
        )

        main_window._on_add_folder_shortcut()

        assert calls == [True]
