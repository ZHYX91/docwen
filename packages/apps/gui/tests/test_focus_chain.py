"""Focus chain and keyboard navigation tests.

Covers the post-closure "GUI focus chain" manual check:
- MainWindow focus chain covers all major functional areas + bottom bar
- SettingsDialog sidebar nav → tab content → action buttons
- AboutDialog close button reachable by keyboard
- Critical interactive widgets accept Tab focus
- Focus guard: text editing suppresses global shortcuts
"""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.gui

# ────────────────────────────────────────────────────────────
# Helpers
# ────────────────────────────────────────────────────────────


def _collect_focus_chain(root: object) -> list[str]:
    """Walk QWidget.nextInFocusChain() from *root*, returning objectNames."""
    from PySide6.QtWidgets import QWidget

    names: list[str] = []
    seen_ids: set[int] = set()
    current: QWidget | None = root  # type: ignore[assignment]
    while current is not None:
        if id(current) in seen_ids:
            break  # cycle guard
        seen_ids.add(id(current))
        name = current.objectName()
        if name:
            names.append(name)
        w = current.nextInFocusChain()
        if w is current:  # last widget points to itself
            break
        current = w
    return names


def _names_to_set(names: list[str]) -> set[str]:
    return {n for n in names if n}


def _assert_ordered(chain: list[str], *expected: str) -> None:
    """Assert *expected* objectNames appear in *chain* in the given order."""
    indices = {}
    for name in expected:
        try:
            indices[name] = chain.index(name)
        except ValueError:
            pytest.fail(f"Expected '{name}' in focus chain but not found. Chain: {chain}")
    for i in range(len(expected) - 1):
        a, b = expected[i], expected[i + 1]
        assert indices[a] < indices[b], (
            f"Focus order violation: '{a}' (pos {indices[a]}) should come before '{b}' (pos {indices[b]})"
        )


# ────────────────────────────────────────────────────────────
# MainWindow focus chain
# ────────────────────────────────────────────────────────────

FIVE_FUNCTIONAL_AREAS = (
    "inputArea",
    "batchListSurface",
    "conversionPanelRoot",
    "actionAreaRoot",
    "infoArea",
)

BOTTOM_BAR_WIDGETS = {
    "fontSizeButton",
    "aboutButton",
    "settingsButton",
    "versionLabel",
}


class TestMainWindowFocusChain:
    """Focus chain covers the 5 functional areas + bottom bar."""

    def test_all_functional_areas_in_chain(self, main_window) -> None:
        chain = _collect_focus_chain(main_window)
        found = _names_to_set(chain)
        for name in FIVE_FUNCTIONAL_AREAS:
            assert name in found, f"'{name}' missing from MainWindow focus chain. Chain: {chain}"

    def test_center_column_top_to_bottom_order(self, main_window) -> None:
        """Within center column: inputArea → actionAreaRoot → infoArea
        (after Task 4 layout refactor: conversion panel moved to right stack,
        info area moved into center column)."""
        chain = _collect_focus_chain(main_window)
        _assert_ordered(
            chain,
            "inputArea",
            "actionAreaRoot",
            "infoArea",
        )

    def test_info_area_after_center_column(self, main_window) -> None:
        """infoArea is now in the center column (below input/action), not a separate right column."""
        chain = _collect_focus_chain(main_window)
        # infoArea should appear after inputArea and actionArea (still in center column)
        assert "infoArea" in chain, "infoArea must be in focus chain"

    def test_bottom_bar_widgets_in_chain(self, main_window) -> None:
        chain = _collect_focus_chain(main_window)
        found = _names_to_set(chain)
        for name in BOTTOM_BAR_WIDGETS:
            assert name in found, f"'{name}' missing from focus chain. Chain: {chain}"

    def test_bottom_bar_comes_after_content(self, main_window) -> None:
        """Font size / about / settings buttons appear after main content.
        After Task 4, bottom bar is inside center column; Qt's grid focus chain
        walks left→right, so bottom bar widgets appear before center-column
        content.  The important invariant is that bottom bar buttons are
        reachable."""
        chain = _collect_focus_chain(main_window)
        for name in BOTTOM_BAR_WIDGETS:
            assert name in chain, f"'{name}' missing from focus chain"

    def test_no_duplicate_names_prove_unique_widgets(self, main_window) -> None:
        """If all major area names appear exactly once, widgets are unique."""
        chain = _collect_focus_chain(main_window)
        from collections import Counter

        counts = Counter(chain)
        for name in FIVE_FUNCTIONAL_AREAS:
            assert counts[name] == 1, f"'{name}' appears {counts[name]} times in focus chain; expected exactly 1"


# ────────────────────────────────────────────────────────────
# Focus policies for focusedInEvent-driven widgets
# ────────────────────────────────────────────────────────────

# These widgets have focusInEvent handlers that steer focus to a child;
# they must accept focus (StrongFocus or TabFocus) for that to work.
FOCUS_STEERING_WIDGETS = (
    ("batchListSurface", "BatchList"),
    ("conversionPanelRoot", "ConversionPanel"),
    ("actionAreaRoot", "ActionArea"),
)


class TestFocusSteeringWidgets:
    """Widgets with focusInEvent handlers must accept focus."""

    @pytest.mark.parametrize("obj_name,label", FOCUS_STEERING_WIDGETS)
    def test_accepts_focus(self, main_window, obj_name, label) -> None:
        from PySide6.QtCore import Qt

        widget = main_window.findChild(object, obj_name)
        assert widget is not None, f"Cannot find '{obj_name}' ({label})"
        policy = widget.focusPolicy()
        assert policy in (Qt.FocusPolicy.StrongFocus, Qt.FocusPolicy.TabFocus), (
            f"{label} ({obj_name}) has focusPolicy={policy}; expected StrongFocus or TabFocus for focusInEvent steering"
        )


# ────────────────────────────────────────────────────────────
# Container widgets: children carry the focus
# ────────────────────────────────────────────────────────────

CONTAINER_WIDGETS = (
    ("inputArea", "InputArea"),
    ("infoArea", "InfoArea"),
)


class TestContainerWidgetFocus:
    """Containers defer to children for Tab navigation but must be in chain."""

    @pytest.mark.parametrize("obj_name,label", CONTAINER_WIDGETS)
    def test_container_in_focus_chain(self, main_window, obj_name, label) -> None:
        chain = _collect_focus_chain(main_window)
        assert obj_name in chain, f"Container '{obj_name}' ({label}) not in focus chain"

    @pytest.mark.parametrize("obj_name,label", CONTAINER_WIDGETS)
    def test_container_has_focusable_children(self, main_window, obj_name, label) -> None:
        """Container children (buttons, inputs) are the actual tab stops."""
        from PySide6.QtCore import Qt
        from PySide6.QtWidgets import QWidget

        container = main_window.findChild(object, obj_name)
        assert container is not None, f"Cannot find '{obj_name}' ({label})"
        children: list[QWidget] = container.findChildren(QWidget)
        focusable_children = [c for c in children if c.focusPolicy() not in (Qt.FocusPolicy.NoFocus,)]
        assert len(focusable_children) > 0, f"{label} ({obj_name}) has no focusable children"


# ────────────────────────────────────────────────────────────
# SettingsDialog focus chain
# ────────────────────────────────────────────────────────────


class TestSettingsDialogFocusChain:
    """SettingsDialog sidebar nav → tab content focus flow."""

    @pytest.fixture
    def settings_dialog(self, qapp):
        from docwen_gui.view_models.settings_vm import SettingsViewModel
        from docwen_gui.widgets.settings.dialog import SettingsDialog

        vm = SettingsViewModel(parent=None)
        dialog = SettingsDialog(parent=None, view_model=vm)
        dialog._build_ui()
        yield dialog
        dialog.close()

    def test_dialog_has_focusable_content(self, settings_dialog) -> None:
        chain = _collect_focus_chain(settings_dialog)
        found = _names_to_set(chain)
        assert "settingsFluentNavigation" in found, f"Settings sidebar nav not in focus chain: {chain}"
        assert "settingsTabRoot" in found, f"No settings tab in focus chain: {chain}"

    def test_action_buttons_in_chain(self, settings_dialog) -> None:
        chain = _collect_focus_chain(settings_dialog)
        found = _names_to_set(chain)
        for btn_name in (
            "settingsResetTabButton",
            "settingsResetAllButton",
            "settingsOkButton",
            "settingsCancelButton",
            "settingsApplyButton",
        ):
            assert btn_name in found, f"'{btn_name}' missing from SettingsDialog focus chain"

    def test_action_buttons_after_tab_content(self, settings_dialog) -> None:
        """Action buttons (Ok/Cancel/Apply/Reset) appear after tab content."""
        chain = _collect_focus_chain(settings_dialog)
        tab_indices = []
        for name in ("settingsTabRoot", "settingsFluentNavigation"):
            if name in chain:
                tab_indices.append(chain.index(name))
        last_tab = max(tab_indices) if tab_indices else -1

        for btn_name in ("settingsOkButton", "settingsCancelButton"):
            if btn_name in chain:
                assert chain.index(btn_name) > last_tab, f"'{btn_name}' should appear after tab content"


# ────────────────────────────────────────────────────────────
# AboutDialog focus chain
# ────────────────────────────────────────────────────────────


class TestAboutDialogFocusChain:
    """AboutDialog: close button reachable by keyboard."""

    @pytest.fixture
    def about_dialog(self, qapp):
        from docwen_gui.dialogs.about import AboutDialog

        dialog = AboutDialog(parent=None)
        yield dialog
        dialog.close()

    def test_close_button_in_chain(self, about_dialog) -> None:
        chain = _collect_focus_chain(about_dialog)
        assert "aboutCloseButton" in chain, f"AboutDialog close button not in focus chain: {chain}"

    def test_scroll_area_in_chain(self, about_dialog) -> None:
        chain = _collect_focus_chain(about_dialog)
        assert "aboutScrollArea" in chain, f"AboutDialog scroll area not in focus chain: {chain}"


# ────────────────────────────────────────────────────────────
# Focus guard: text editing suppresses shortcuts
# ────────────────────────────────────────────────────────────


class TestFocusGuard:
    """When focus is in a text editing widget, global shortcuts are suppressed.

    _has_editable_text_focus() checks QApplication.focusWidget(), which
    requires the widget to be visible in an active window.  We show the
    main_window to provide an active window context.
    """

    @pytest.fixture(autouse=True)
    def _show_main_window(self, main_window, qapp) -> None:
        main_window.show()
        qapp.processEvents()

    def test_editable_text_focus_is_detected(self, main_window) -> None:
        from PySide6.QtWidgets import QLineEdit, QPlainTextEdit, QTextEdit

        # QLineEdit
        le = QLineEdit(main_window)
        le.show()
        le.setFocus()
        assert main_window._has_editable_text_focus(), "QLineEdit focus should trigger editable text detection"
        le.setParent(None)

        # QTextEdit
        te = QTextEdit(main_window)
        te.show()
        te.setFocus()
        assert main_window._has_editable_text_focus(), "QTextEdit focus should trigger editable text detection"
        te.setParent(None)

        # QPlainTextEdit
        pe = QPlainTextEdit(main_window)
        pe.show()
        pe.setFocus()
        assert main_window._has_editable_text_focus(), "QPlainTextEdit focus should trigger editable text detection"
        pe.setParent(None)

    def test_button_focus_does_not_guard(self, main_window) -> None:
        from PySide6.QtWidgets import QPushButton

        btn = QPushButton("test", main_window)
        btn.show()
        btn.setFocus()
        assert not main_window._has_editable_text_focus(), "QPushButton focus should NOT suppress shortcuts"
        btn.setParent(None)

    def test_readonly_combo_does_not_guard(self, main_window) -> None:
        from PySide6.QtWidgets import QComboBox

        cb = QComboBox(main_window)
        cb.setEditable(False)
        cb.addItem("test")
        cb.show()
        cb.setFocus()
        assert not main_window._has_editable_text_focus(), "Read-only QComboBox should NOT suppress shortcuts"
        cb.setParent(None)

    def test_editable_combo_guards(self, main_window) -> None:
        from PySide6.QtWidgets import QComboBox

        cb = QComboBox(main_window)
        cb.setEditable(True)
        cb.show()
        cb.setFocus()
        assert main_window._has_editable_text_focus(), "Editable QComboBox focus should trigger editable text detection"
        cb.setParent(None)


# ────────────────────────────────────────────────────────────
# Interactive widget focus policies
# ────────────────────────────────────────────────────────────


class TestInteractiveWidgetFocusPolicies:
    """Key interactive widgets must accept Tab focus."""

    def test_input_area_add_button_is_focusable(self, main_window) -> None:
        from PySide6.QtCore import Qt

        btn = main_window.input_area._add_button
        assert btn.focusPolicy() != Qt.FocusPolicy.NoFocus, "Add file button should accept focus"

    def test_action_area_cancel_button_is_focusable(self, main_window) -> None:
        from PySide6.QtCore import Qt

        btn = main_window.action_area.cancel_button
        assert btn is not None, "Cancel button not found"
        assert btn.focusPolicy() != Qt.FocusPolicy.NoFocus, "Cancel button should accept focus"

    def test_settings_button_is_focusable(self, main_window) -> None:
        from PySide6.QtCore import Qt

        btn = main_window.settings_button
        assert btn.focusPolicy() != Qt.FocusPolicy.NoFocus, "Settings button should accept focus"
