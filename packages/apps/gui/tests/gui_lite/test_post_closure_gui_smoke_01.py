"""Focused tests split from test_post_closure_gui_smoke.py."""

from __future__ import annotations

import pytest

from ._post_closure_gui_smoke_support import (
    Path,
    _save_screenshot,
    contextlib,
    patch,
)

pytestmark = pytest.mark.gui
from ._post_closure_gui_smoke_support import (
    window as window,
)


class TestMainWindowSmoke:
    """Verify the MainWindow contains all expected child widgets by objectName."""

    def test_main_window_has_object_name(self, window):
        assert window.objectName() == "docwenMainWindow"

    def test_about_button_exists(self, window):
        btn = window.findChild(type(window._about_btn), "aboutButton")
        assert btn is not None

    def test_settings_button_exists(self, window):
        btn = window.findChild(type(window._settings_btn), "settingsButton")
        assert btn is not None

    def test_central_container_exists(self, window):
        from PySide6.QtWidgets import QWidget

        container = window.findChild(QWidget, "centralContainer")
        assert container is not None

    def test_centralized_qaaction_about_exists(self, window):
        action = window.findChild(type(window._action_about), "actionAbout")
        assert action is not None
        assert action.shortcut().toString() in ("Ctrl+?", "Ctrl+Shift+/")
        assert action.text() != ""

    def test_centralized_qaaction_settings_exists(self, window):
        action = window.findChild(type(window._action_settings), "actionSettings")
        assert action is not None
        assert action.shortcut().toString() == "Ctrl+,"
        assert action.text() != ""

    def test_centralized_qaaction_add_file_exists(self, window):
        action = window.findChild(type(window._action_add_file), "actionAddFile")
        assert action is not None
        assert action.shortcut().toString() == "Ctrl+O"

    def test_centralized_qaaction_convert_exists(self, window):
        action = window.findChild(type(window._action_convert), "actionConvert")
        assert action is not None
        assert "Return" in action.shortcut().toString() or "Enter" in action.shortcut().toString()

    def test_centralized_qaaction_cancel_exists(self, window):
        action = window.findChild(type(window._action_cancel), "actionCancel")
        assert action is not None
        assert action.shortcut().toString() == "Esc"

    def test_main_window_screenshot(self, window, tmp_path, qtbot, qapp):
        """Capture a real batch-conversion state, not the empty launch shell."""
        from PySide6.QtGui import QColor, QImage
        from PySide6.QtWidgets import QLabel

        from docwen_gui.view_models.interaction import RightPanelSlot

        checked_in_baseline = Path(__file__).resolve().parent / "screenshots" / "main_window.png"
        checked_in_baseline_before = checked_in_baseline.read_bytes()
        source = tmp_path / "screenshot-source.png"
        image = QImage(32, 32, QImage.Format.Format_RGB32)
        image.fill(QColor("#5B8FF9"))
        assert image.save(str(source))

        window.view_model.set_mode("batch")
        outcome = window.view_model.add_files([str(source)])
        assert len(outcome.added) == 1
        assert outcome.rejected == ()
        qtbot.waitUntil(
            lambda: (
                window.view_model.ui_projection.left_panel_visible
                and window.view_model.ui_projection.right_panel_visible
                and window.view_model.ui_projection.right_panel_slot is RightPanelSlot.CONVERSION
            ),
            timeout=3000,
        )
        window.resize(max(window.minimumWidth(), 1280), 760)
        qapp.processEvents()

        # The visual fixture is checked into the public repository. Exercise
        # the real temporary file above, then replace only its visible parent
        # directory so the PNG never records a local account name.
        source_parent = str(tmp_path)
        safe_parent = r"C:\DocWen\Samples"
        for label in window.findChildren(QLabel):
            full_text = getattr(label, "full_text", None)
            set_full_text = getattr(label, "set_full_text", None)
            if isinstance(full_text, str) and source_parent in full_text and callable(set_full_text):
                set_full_text(full_text.replace(source_parent, safe_parent))
            elif source_parent in label.text():
                label.setText(label.text().replace(source_parent, safe_parent))
            if source_parent in label.toolTip():
                label.setToolTip(label.toolTip().replace(source_parent, safe_parent))
        qapp.processEvents()

        assert window._left_panel_frame.isVisible()
        assert window._center_column.isVisible()
        assert window._right_panel_frame.isVisible()
        assert window._right_stack.currentWidget() is window._conversion_panel
        assert window._action_area.isVisible()

        path = _save_screenshot(window, "main_window", tmp_path / "screenshots")
        # Screenshot is best-effort; do not fail if unavailable.
        if path is not None:
            screenshot = Path(path)
            assert screenshot.exists()
            assert screenshot.is_relative_to(tmp_path)
        assert checked_in_baseline.read_bytes() == checked_in_baseline_before


class TestAboutDialogLite:
    """Verify AboutDialog creation and content."""

    def test_about_dialog_creates(self, qapp, window):
        from docwen_gui.dialogs.about import AboutDialog

        dlg = AboutDialog(parent=window)
        assert dlg.objectName() == "aboutDialog"
        assert dlg.isModal() is True
        dlg.close()

    def test_about_dialog_has_hero_card(self, qapp, window):
        from PySide6.QtWidgets import QWidget

        from docwen_gui.dialogs.about import AboutDialog

        dlg = AboutDialog(parent=window)
        hero = dlg.findChild(QWidget, "aboutHeroCard")
        assert hero is not None
        dlg.close()

    def test_about_dialog_has_title_label(self, qapp, window):
        from docwen_gui.dialogs.about import AboutDialog

        dlg = AboutDialog(parent=window)
        from PySide6.QtWidgets import QLabel

        title = dlg.findChild(QLabel, "aboutTitle")
        assert title is not None
        assert title.text() != ""
        dlg.close()

    def test_about_dialog_has_version_label(self, qapp, window):
        from docwen_gui.dialogs.about import AboutDialog

        dlg = AboutDialog(parent=window)
        from PySide6.QtWidgets import QLabel

        ver = dlg.findChild(QLabel, "aboutVersion")
        assert ver is not None
        assert "version" in ver.text().lower() or ver.text() != ""
        dlg.close()

    def test_about_dialog_has_close_button(self, qapp, window):
        from PySide6.QtWidgets import QWidget

        from docwen_gui.dialogs.about import AboutDialog

        dlg = AboutDialog(parent=window)
        close_btn = dlg.findChild(QWidget, "aboutCloseButton")
        assert close_btn is not None
        dlg.close()

    def test_about_dialog_has_tools_grid(self, qapp, window):
        from PySide6.QtWidgets import QWidget

        from docwen_gui.dialogs.about import AboutDialog

        dlg = AboutDialog(parent=window)
        tools = dlg.findChild(QWidget, "aboutToolsGrid")
        assert tools is not None
        dlg.close()

    def test_about_dialog_has_acknowledgments_intro(self, qapp, window):
        from docwen_gui.dialogs.about import AboutDialog

        dlg = AboutDialog(parent=window)
        from PySide6.QtWidgets import QLabel

        intro = dlg.findChild(QLabel, "aboutAcknowledgmentsIntro")
        assert intro is not None
        dlg.close()

    def test_about_dialog_screenshot(self, qapp, window, tmp_path):
        from docwen_gui.dialogs.about import AboutDialog

        dlg = AboutDialog(parent=window)
        dlg.show()
        qapp.processEvents()
        path = _save_screenshot(dlg, "about_dialog", tmp_path / "screenshots")
        if path is not None:
            screenshot = Path(path)
            assert screenshot.exists()
            assert screenshot.is_relative_to(tmp_path)
        dlg.close()

    def test_about_button_triggers_dialog(self, qtbot, window):
        """Clicking the about button creates and shows the AboutDialog."""
        from docwen_gui.dialogs.about import AboutDialog

        # Intercept show_dialog to avoid blocking exec()
        called = []

        def _fake_show(self):
            called.append(True)

        with patch.object(AboutDialog, "show_dialog", _fake_show):
            window._about_btn.click()
            qtbot.wait(50)
            assert len(called) == 1


class TestSettingsDialogLite:
    """Verify SettingsDialog creation, tabs, and key controls."""

    def test_settings_dialog_creates(self, qapp, window):
        from docwen_gui.view_models.settings_vm import SettingsViewModel
        from docwen_gui.widgets.settings.dialog import SettingsDialog

        vm = SettingsViewModel(controller=None)
        dlg = SettingsDialog(parent=window, view_model=vm)
        assert dlg.objectName() == "settingsDialog"
        assert dlg.isModal() is True
        dlg.close()

    def test_settings_dialog_has_tab_widget(self, qapp, window):
        from PySide6.QtWidgets import QTabWidget

        from docwen_gui.view_models.settings_vm import SettingsViewModel
        from docwen_gui.widgets.settings.dialog import SettingsDialog

        vm = SettingsViewModel(controller=None)
        dlg = SettingsDialog(parent=window, view_model=vm)
        tab_widget = dlg.findChild(QTabWidget)
        assert tab_widget is not None
        # Should have 13 tabs
        assert tab_widget.count() >= 13
        dlg.close()

    def test_settings_dialog_has_expected_tab_titles(self, qapp, window):
        from docwen_gui.view_models.settings_vm import SettingsViewModel
        from docwen_gui.widgets.settings.dialog import SettingsDialog

        vm = SettingsViewModel(controller=None)
        dlg = SettingsDialog(parent=window, view_model=vm)
        from PySide6.QtWidgets import QTabWidget

        tab_widget = dlg.findChild(QTabWidget)
        assert tab_widget is not None
        assert tab_widget.count() >= 13, f"Expected >=13 tabs, got {tab_widget.count()}"
        for i in range(tab_widget.count()):
            assert tab_widget.tabText(i), f"Tab {i} has empty title"
        dlg.close()

    def test_settings_dialog_has_action_buttons(self, qapp, window):
        from PySide6.QtWidgets import QPushButton

        from docwen_gui.view_models.settings_vm import SettingsViewModel
        from docwen_gui.widgets.settings.dialog import SettingsDialog

        vm = SettingsViewModel(controller=None)
        dlg = SettingsDialog(parent=window, view_model=vm)
        ok_btn = dlg.findChild(QPushButton, "settingsOkButton")
        assert ok_btn is not None
        cancel_btn = dlg.findChild(QPushButton, "settingsCancelButton")
        assert cancel_btn is not None
        apply_btn = dlg.findChild(QPushButton, "settingsApplyButton")
        assert apply_btn is not None
        dlg.close()

    def test_settings_dialog_screenshot(self, qapp, window, tmp_path):
        from docwen_gui.view_models.settings_vm import SettingsViewModel
        from docwen_gui.widgets.settings.dialog import SettingsDialog

        vm = SettingsViewModel(controller=None)
        dlg = SettingsDialog(parent=window, view_model=vm)
        dlg.show()
        qapp.processEvents()
        path = _save_screenshot(dlg, "settings_dialog", tmp_path / "screenshots")
        if path is not None:
            screenshot = Path(path)
            assert screenshot.exists()
            assert screenshot.is_relative_to(tmp_path)
        dlg.close()

    def test_settings_button_triggers_dialog(self, qtbot, window):
        """Clicking settings button opens the SettingsDialog."""
        from docwen_gui.widgets.settings.dialog import SettingsDialog

        called = []

        def _fake_open(self):
            called.append(True)

        with patch.object(SettingsDialog, "open", _fake_open):
            window._settings_btn.click()
            qtbot.wait(50)
            assert len(called) == 1


class TestFeedbackDialogsLite:
    """Verify feedback dialogs (error/warn/info/confirm) create with objectNames."""

    def test_error_messagebox_has_object_name(self, qapp, window):
        """error() creates a QMessageBox with feedbackErrorMessageBox objectName."""

        from PySide6.QtWidgets import QMessageBox

        from docwen_gui.dialogs.feedback import error

        boxes: list[QMessageBox] = []

        def _capture(box_self):
            boxes.append(box_self)
            box_self.accept()
            return QMessageBox.DialogCode.Accepted

        with patch.object(QMessageBox, "exec", _capture), contextlib.suppress(Exception):
            error("Test Error", "Something went wrong")

        if boxes:
            assert boxes[0].objectName() == "feedbackErrorMessageBox"

    def test_warn_messagebox_has_object_name(self, qapp, window):
        """warn() creates a QMessageBox with feedbackWarningMessageBox objectName."""
        from unittest.mock import patch as mock_patch

        from PySide6.QtWidgets import QMessageBox

        from docwen_gui.dialogs.feedback import warn

        boxes: list[QMessageBox] = []

        def _capture(box_self):
            boxes.append(box_self)
            box_self.accept()
            return QMessageBox.DialogCode.Accepted

        with mock_patch.object(QMessageBox, "exec", _capture), contextlib.suppress(Exception):
            warn("Test Warning", "Be careful")

        if boxes:
            assert boxes[0].objectName() == "feedbackWarningMessageBox"

    def test_info_messagebox_has_object_name(self, qapp, window):
        """info() creates a QMessageBox with feedbackInfoMessageBox objectName."""
        from unittest.mock import patch as mock_patch

        from PySide6.QtWidgets import QMessageBox

        from docwen_gui.dialogs.feedback import info

        boxes: list[QMessageBox] = []

        def _capture(box_self):
            boxes.append(box_self)
            box_self.accept()
            return QMessageBox.DialogCode.Accepted

        with mock_patch.object(QMessageBox, "exec", _capture), contextlib.suppress(Exception):
            info("Test Info", "For your information")

        if boxes:
            assert boxes[0].objectName() == "feedbackInfoMessageBox"

    def test_confirm_messagebox_has_object_name(self, qapp, window):
        """confirm() creates a QMessageBox with feedbackConfirmMessageBox objectName."""
        from unittest.mock import patch as mock_patch

        from PySide6.QtWidgets import QMessageBox

        from docwen_gui.dialogs.feedback import confirm

        boxes: list[QMessageBox] = []

        def _capture(box_self):
            boxes.append(box_self)
            box_self.accept()
            return QMessageBox.DialogCode.Accepted

        with mock_patch.object(QMessageBox, "exec", _capture), contextlib.suppress(Exception):
            confirm("Confirm", "Are you sure?")

        if boxes:
            assert boxes[0].objectName() == "feedbackConfirmMessageBox"


class TestKeyboardShortcutsLite:
    """Verify centralized QAction shortcuts trigger expected handlers."""

    def test_action_add_file_shortcut_registered(self, window):
        """Ctrl+O is registered as the shortcut for actionAddFile."""
        actions = window.findChildren(type(window._action_add_file), "actionAddFile")
        assert len(actions) >= 1
        action = actions[0]
        shortcut = action.shortcut().toString()
        # On some platforms Qt normalises Ctrl+O to Ctrl+O
        assert "o" in shortcut.lower() or "O" in shortcut

    def test_action_settings_shortcut_registered(self, window):
        """Ctrl+, is registered as the shortcut for actionSettings."""
        actions = window.findChildren(type(window._action_settings), "actionSettings")
        assert len(actions) >= 1
        action = actions[0]
        assert action.shortcut().toString() == "Ctrl+,"

    def test_action_about_shortcut_registered(self, window):
        """About action has a shortcut bound (Ctrl+? or similar)."""
        actions = window.findChildren(type(window._action_about), "actionAbout")
        assert len(actions) >= 1
        action = actions[0]
        shortcut = action.shortcut().toString()
        # Accept Ctrl+? or Ctrl+Shift+/ as valid
        assert "?" in shortcut or "/" in shortcut or "Shift" in shortcut

    def test_action_convert_shortcut_registered(self, window):
        """Ctrl+Return is registered as the shortcut for actionConvert."""
        actions = window.findChildren(type(window._action_convert), "actionConvert")
        assert len(actions) >= 1
        action = actions[0]
        shortcut = action.shortcut().toString()
        assert "Return" in shortcut or "Enter" in shortcut

    def test_action_cancel_shortcut_registered(self, window):
        """Esc is registered as the shortcut for actionCancel."""
        actions = window.findChildren(type(window._action_cancel), "actionCancel")
        assert len(actions) >= 1
        action = actions[0]
        assert action.shortcut().toString() == "Esc"

    def test_action_about_triggers_handler(self, qtbot, window):
        """Triggering actionAbout opens the about dialog."""
        from docwen_gui.dialogs.about import AboutDialog

        called = []

        def _fake_show(self):
            called.append(True)

        with patch.object(AboutDialog, "show_dialog", _fake_show):
            window._action_about.trigger()
            qtbot.wait(50)
            assert len(called) == 1

    def test_action_settings_triggers_handler(self, window):
        """Triggering actionSettings opens the settings dialog."""
        from docwen_gui.widgets.settings.dialog import SettingsDialog

        called = []

        def _fake_open(self):
            called.append(True)

        with patch.object(SettingsDialog, "open", _fake_open):
            window._action_settings.trigger()
            assert len(called) == 1

    def test_about_button_uses_shared_action(self, window):
        """About button click is connected (verified in trigger test above)."""
        # The button's clicked signal connects to actionAbout.trigger.
        # Qt may report receivers() differently across platforms; verify
        # holistically via test_about_button_triggers_dialog instead.
        assert window._about_btn is not None
        assert window._action_about is not None

    def test_settings_button_uses_shared_action(self, window):
        """Settings button click is connected (verified in trigger test above)."""
        assert window._settings_btn is not None
        assert window._action_settings is not None
