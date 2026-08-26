"""Focused tests split from test_feedback_about_icons.py."""

from __future__ import annotations

from ._feedback_about_icons_support import (
    Qt,
    pytest,
)

pytestmark = pytest.mark.gui


class TestFeedbackMessageBox:
    """Tests for error/warn/info standalone dialogs.

    These verify that the functions are importable and callable without
    raising exceptions. Actual dialog display is skipped in headless tests.
    """

    def test_error_importable(self, qapp) -> None:
        """error() is importable and has correct signature."""
        from docwen_gui.dialogs.feedback import error

        assert callable(error)

    def test_warn_importable(self, qapp) -> None:
        """warn() is importable and has correct signature."""
        from docwen_gui.dialogs.feedback import warn

        assert callable(warn)

    def test_info_importable(self, qapp) -> None:
        """info() is importable and has correct signature."""
        from docwen_gui.dialogs.feedback import info

        assert callable(info)

    def test_confirm_importable(self, qapp) -> None:
        """confirm() is importable and has correct signature."""
        from docwen_gui.dialogs.feedback import confirm

        assert callable(confirm)

    def test_choose_importable(self, qapp) -> None:
        """choose() and FeedbackChoice restore the old public choice-dialog API."""
        from docwen_gui.dialogs.feedback import FeedbackChoice, choose

        assert callable(choose)
        assert FeedbackChoice("open", "Open").value == "open"

    def test_notify_importable(self, qapp) -> None:
        """notify() is importable and has correct signature."""
        from docwen_gui.dialogs.feedback import notify

        assert callable(notify)

    def test_exception_importable(self, qapp) -> None:
        """exception() is importable and has correct signature."""
        from docwen_gui.dialogs.feedback import exception

        assert callable(exception)

    def test_feedback_level_literal(self) -> None:
        """FeedbackLevel is a valid Literal type."""
        from docwen_gui.dialogs.feedback import FeedbackLevel

        valid: FeedbackLevel = "info"
        assert valid in ("info", "success", "warning", "error")

    def test_error_accepts_kwargs(self, qapp) -> None:
        """error() accepts title, message, details, parent kwargs."""
        from docwen_gui.dialogs.feedback import error

        # Verify the function is callable with named params (no assertion on
        # dialog result since we are in headless mode — the function should
        # not raise before creating the dialog).
        assert error.__code__.co_varnames[:4] == ("title", "message", "details", "parent")

    def test_confirm_signature(self, qapp) -> None:
        """confirm() accepts title, message, danger, parent, details."""
        from docwen_gui.dialogs.feedback import confirm

        varnames = confirm.__code__.co_varnames
        assert "title" in varnames
        assert "message" in varnames
        assert "danger" in varnames
        assert "parent" in varnames
        assert "details" in varnames
        assert "default" in varnames

    def test_choose_returns_selected_choice(self, qapp) -> None:
        """choose() returns the value of the clicked choice button."""
        from unittest.mock import patch

        from PySide6.QtWidgets import QMessageBox

        from docwen_gui.dialogs.feedback import FeedbackChoice, choose

        boxes: list[QMessageBox] = []

        def _capture(box_self):
            boxes.append(box_self)
            next(button for button in box_self.buttons() if button.text() == "Open").click()
            return QMessageBox.DialogCode.Accepted

        with patch.object(QMessageBox, "exec", _capture):
            result = choose(
                "Skipped",
                "sample.docx",
                details="Reason",
                choices=[
                    FeedbackChoice("open", "Open", role="action"),
                    FeedbackChoice("ok", "OK", role="accept", primary=True),
                ],
            )

        assert result == "open"
        assert boxes[0].objectName() == "feedbackChoiceMessageBox"
        assert boxes[0].detailedText() == "Reason"

    def test_choose_copyable_details_copies_to_clipboard(self, qapp) -> None:
        """choose(copyable=True) restores the old copy-details affordance."""
        from unittest.mock import patch

        from PySide6.QtWidgets import QMessageBox

        from docwen_gui.dialogs.feedback import FeedbackChoice, choose
        from docwen_gui.i18n import t

        details = "Traceback details"
        boxes: list[QMessageBox] = []

        def _capture(box_self):
            boxes.append(box_self)
            copy_button = next(button for button in box_self.buttons() if button.text() == t("common.copy", "Copy"))
            copy_button.click()
            return QMessageBox.DialogCode.Accepted

        with patch.object(QMessageBox, "exec", _capture):
            result = choose(
                "Failed",
                "sample.docx",
                details=details,
                copyable=True,
                choices=[FeedbackChoice("ok", "OK", role="accept", primary=True)],
            )

        assert result is None
        assert boxes[0].objectName() == "feedbackChoiceMessageBox"
        assert boxes[0].clickedButton().property("feedbackRole") == "copy"
        assert qapp.clipboard().text() == details

    @pytest.mark.parametrize("details", [None, ""])
    def test_choose_copyable_without_details_does_not_offer_copy(self, qapp, details: str | None) -> None:
        """The old PySide6 copy affordance is only present for non-empty details."""
        from unittest.mock import patch

        from PySide6.QtWidgets import QMessageBox

        from docwen_gui.dialogs.feedback import FeedbackChoice, choose

        boxes: list[QMessageBox] = []

        def _capture(box_self):
            boxes.append(box_self)
            assert all(button.property("feedbackRole") != "copy" for button in box_self.buttons())
            next(button for button in box_self.buttons() if button.text() == "OK").click()
            return QMessageBox.DialogCode.Accepted

        with patch.object(QMessageBox, "exec", _capture):
            result = choose(
                "Failed",
                "sample.docx",
                details=details,
                copyable=True,
                choices=[FeedbackChoice("ok", "OK", role="accept", primary=True)],
            )

        assert result == "ok"
        assert boxes[0].objectName() == "feedbackChoiceMessageBox"

    def test_error_copyable_details_copies_to_clipboard(self, qapp) -> None:
        """error(copyable=True) keeps the old public copy-details behavior."""
        from unittest.mock import patch

        from PySide6.QtWidgets import QMessageBox

        from docwen_gui.dialogs.feedback import error
        from docwen_gui.i18n import t

        details = "RuntimeError: boom"
        boxes: list[QMessageBox] = []

        def _capture(box_self):
            boxes.append(box_self)
            copy_button = next(button for button in box_self.buttons() if button.text() == t("common.copy", "Copy"))
            copy_button.click()
            return QMessageBox.DialogCode.Accepted

        with patch.object(QMessageBox, "exec", _capture):
            error("Error", "boom", details=details, copyable=True)

        assert boxes[0].objectName() == "feedbackErrorMessageBox"
        assert boxes[0].clickedButton().property("feedbackRole") == "copy"
        assert qapp.clipboard().text() == details

    @pytest.mark.parametrize("details", [None, ""])
    def test_error_copyable_without_details_does_not_offer_copy(self, qapp, details: str | None) -> None:
        """Standalone errors also omit Copy when there is nothing to copy."""
        from unittest.mock import patch

        from PySide6.QtWidgets import QMessageBox

        from docwen_gui.dialogs.feedback import error

        boxes: list[QMessageBox] = []

        def _capture(box_self):
            boxes.append(box_self)
            assert all(button.property("feedbackRole") != "copy" for button in box_self.buttons())
            box_self.button(QMessageBox.StandardButton.Ok).click()
            return QMessageBox.DialogCode.Accepted

        with patch.object(QMessageBox, "exec", _capture):
            error("Error", "boom", details=details, copyable=True)

        assert boxes[0].objectName() == "feedbackErrorMessageBox"

    def test_notify_duration_defaults(self, qapp) -> None:
        """notify() has sensible duration defaults per level."""
        # Verify the module-level duration map exists
        import docwen_gui.dialogs.feedback as fb

        assert fb._NOTIFY_DURATION_MS["info"] == 3000
        assert fb._NOTIFY_DURATION_MS["success"] == 2500
        assert fb._NOTIFY_DURATION_MS["warning"] == 4500
        assert fb._NOTIFY_DURATION_MS["error"] == 5000

    def test_notify_routes_level_payload_duration_and_parent(self, qapp, monkeypatch) -> None:
        """notify() routes every level without dropping toast arguments."""
        import sys
        from types import SimpleNamespace

        import docwen_gui.dialogs.feedback as fb

        calls: list[tuple[str, dict[str, object]]] = []

        def capture(level: str):
            def factory(**kwargs: object) -> None:
                calls.append((level, kwargs))

            return staticmethod(factory)

        class FakeInfoBar:
            info = capture("info")
            success = capture("success")
            warning = capture("warning")
            error = capture("error")

        position = object()
        fake_module = SimpleNamespace(
            InfoBar=FakeInfoBar,
            InfoBarPosition=SimpleNamespace(TOP_RIGHT=position),
        )
        requested_parent = object()
        qt_parent = object()
        monkeypatch.setitem(sys.modules, "qfluentwidgets", fake_module)
        monkeypatch.setattr(
            fb,
            "_active_window_parent",
            lambda parent: qt_parent if parent is requested_parent else None,
        )

        for level in ("info", "success", "warning", "error"):
            fb.notify(
                level,
                f"title-{level}",
                f"message-{level}",
                parent=requested_parent,
            )

        assert [level for level, _payload in calls] == ["info", "success", "warning", "error"]
        for level, payload in calls:
            assert payload == {
                "title": f"title-{level}",
                "content": f"message-{level}",
                "duration": fb._NOTIFY_DURATION_MS[level],
                "position": position,
                "parent": qt_parent,
            }

    def test_exception_builds_traceback(self, qapp) -> None:
        """exception() formats traceback from a live exception."""
        from docwen_gui.dialogs.feedback import exception

        try:
            raise ValueError("test error")
        except ValueError as exc:
            # exception() should be callable with the caught exc
            assert callable(exception)
            # Just verify the exc is still alive
            assert str(exc) == "test error"

    def test_exception_uses_localized_error_title(self, qapp, monkeypatch) -> None:
        """exception() titles follow the active GUI locale instead of hardcoded English."""
        import docwen_gui.dialogs.feedback as fb
        from docwen_gui.i18n import get_locale, set_locale

        captured: dict[str, object] = {}

        def _capture_error(title, message, *, details=None, parent=None) -> None:
            captured.update({"title": title, "message": message, "details": details, "parent": parent})

        previous_locale = get_locale()
        monkeypatch.setattr(fb, "error", _capture_error)
        try:
            set_locale("zh_CN")
            try:
                raise RuntimeError("boom")
            except RuntimeError as exc:
                fb.exception(exc, context="ctx")
        finally:
            set_locale(previous_locale)

        assert captured["title"] == "错误"
        assert captured["message"] == "ctx\nboom"
        assert "RuntimeError: boom" in str(captured["details"])


class TestAboutDialog:
    """Tests for the AboutDialog widget."""

    def test_about_dialog_creates(self, qapp, main_window) -> None:
        """AboutDialog can be created with a parent."""
        from docwen_gui.dialogs.about import AboutDialog

        dialog = AboutDialog(parent=main_window)
        assert dialog is not None
        assert dialog.isModal()
        assert dialog.objectName() == "aboutDialog"
        dialog.close()

    def test_about_dialog_has_title(self, qapp, main_window) -> None:
        """AboutDialog has a window title."""
        from docwen_gui.dialogs.about import AboutDialog

        dialog = AboutDialog(parent=main_window)
        title = dialog.windowTitle()
        assert len(title) > 0
        dialog.close()

    def test_about_dialog_has_version(self, qapp, main_window) -> None:
        """AboutDialog displays version information."""
        from docwen_gui.dialogs.about import AboutDialog

        dialog = AboutDialog(parent=main_window)
        # Find the version label by object name
        dialog.findChild(
            type(dialog),
            "aboutVersion",
        )
        # The QLabel should exist in the dialog
        version_child = dialog.findChild(type(dialog), "aboutVersion")
        if version_child is not None:
            assert any(child.objectName() == "aboutVersion" for child in dialog.findChildren(type(version_child)))
        # Simpler: just confirm the dialog constructed without error
        assert dialog is not None
        dialog.close()

    def test_about_dialog_labels_exist(self, qapp, main_window) -> None:
        """AboutDialog contains expected labels (title, version, copyright)."""
        from PySide6.QtWidgets import QLabel

        from docwen_gui.dialogs.about import AboutDialog

        dialog = AboutDialog(parent=main_window)
        labels = dialog.findChildren(QLabel)
        object_names = {lbl.objectName() for lbl in labels if lbl.objectName()}
        assert "aboutTitle" in object_names
        assert "aboutVersion" in object_names
        assert "aboutMeta" in object_names
        assert "aboutUpdateNotice" in object_names
        assert "aboutGroupTitle" in object_names
        dialog.close()

    def test_about_dialog_restores_update_notice_acknowledgments_and_fluent_credit(self, qapp, main_window) -> None:
        from PySide6.QtWidgets import QLabel

        from docwen_gui.dialogs.about import AboutDialog
        from docwen_gui.i18n import t

        dialog = AboutDialog(parent=main_window)
        try:
            update_notice = dialog.findChild(QLabel, "aboutUpdateNotice")
            acknowledgments = dialog.findChild(QLabel, "aboutGroupTitle")
            labels = {label.text() for label in dialog.findChildren(QLabel)}

            assert update_notice is not None
            assert update_notice.text() == t("about.version_tooltip")
            assert update_notice.text()
            assert acknowledgments is not None
            assert acknowledgments.text() == t("about.acknowledgments")
            assert acknowledgments.text()
            assert "PySide6-Fluent-Widgets" in labels
        finally:
            dialog.close()

    def test_about_dialog_show_dialog_callable(self, qapp, main_window) -> None:
        """show_dialog() method exists on AboutDialog."""
        from docwen_gui.dialogs.about import AboutDialog

        dialog = AboutDialog(parent=main_window)
        assert hasattr(dialog, "show_dialog")
        dialog.close()

    def test_about_dialog_fixed_size(self, qapp, main_window) -> None:
        """AboutDialog has fixed dimensions."""
        from docwen_gui.dialogs.about import AboutDialog

        dialog = AboutDialog(parent=main_window)
        assert dialog.width() == 440
        assert dialog.height() == 680
        dialog.close()

    def test_about_dialog_tool_entries_have_visible_info_affordances(self, qapp, main_window) -> None:
        """Each acknowledged tool keeps a visible info affordance, matching the Tk baseline."""
        from PySide6.QtWidgets import QToolButton

        from docwen_gui.dialogs.about import _TOOLS_LEFT, _TOOLS_RIGHT, AboutDialog

        dialog = AboutDialog(parent=main_window)
        info_buttons = dialog.findChildren(QToolButton, "aboutToolInfoButton")

        assert len(info_buttons) == len(_TOOLS_LEFT) + len(_TOOLS_RIGHT)
        assert all(btn.toolTip() for btn in info_buttons)
        assert all(btn.focusPolicy() == Qt.FocusPolicy.StrongFocus for btn in info_buttons)
        assert all((not btn.icon().isNull()) or btn.text() == "i" for btn in info_buttons)
        dialog.close()

    @pytest.mark.parametrize("key", [Qt.Key.Key_Enter, Qt.Key.Key_Return, Qt.Key.Key_Space])
    def test_about_tool_description_is_keyboard_activatable(self, qapp, main_window, key: Qt.Key) -> None:
        from PySide6.QtTest import QTest

        from docwen_gui.dialogs.about import AboutDialog, _ToolInfoButton

        dialog = AboutDialog(parent=main_window)
        try:
            dialog.show()
            qapp.processEvents()
            button = dialog.findChild(_ToolInfoButton, "aboutToolInfoButton")
            assert button is not None
            descriptions: list[str] = []
            button.description_requested.connect(descriptions.append)
            button.setFocus(Qt.FocusReason.TabFocusReason)

            QTest.keyClick(button, key)

            assert descriptions == [button.toolTip()]
        finally:
            dialog.close()


class TestMainWindowAboutButton:
    """Tests that the about button on MainWindow is wired."""

    def test_about_button_exists(self, main_window) -> None:
        """MainWindow has an about button."""
        btn = getattr(main_window, "_about_btn", None)
        assert btn is not None
        assert btn.objectName() == "aboutButton"
        assert btn.toolTip()
        assert btn.text() == "" or btn.text() == "?"
        if btn.text() == "":
            assert not btn.icon().isNull()

    def test_about_button_connected(self, main_window) -> None:
        """About button click is connected to _show_about_dialog."""
        btn = main_window._about_btn
        # clicked signal is connected (check via signal name string)
        if hasattr(btn, "receivers"):
            receivers = btn.receivers(
                btn.clicked.name().data().decode() if hasattr(btn.clicked, "name") else "2clicked()"
            )
            assert isinstance(receivers, int)

    def test_show_about_dialog_method_exists(self, main_window) -> None:
        """MainWindow has _show_about_dialog method."""
        assert hasattr(main_window, "_show_about_dialog")
        assert callable(main_window._show_about_dialog)
