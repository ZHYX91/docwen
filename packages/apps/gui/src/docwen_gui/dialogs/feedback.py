"""GUI feedback helpers — error/warning/info message boxes and toast notifications.

Provides standalone dialog functions that can be called from any part of
the GUI without depending on the main window or ViewModel infrastructure.
"""

from __future__ import annotations

import traceback
from dataclasses import dataclass
from typing import Any, Literal
from typing import cast as _cast

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QMessageBox, QWidget

from ..i18n import t

FeedbackLevel = Literal["info", "success", "warning", "error"]
FeedbackRole = Literal["accept", "reject", "action"]


@dataclass(frozen=True)
class FeedbackChoice:
    value: str
    label: str
    role: FeedbackRole = "action"
    primary: bool = False


def _resolve_parent(parent: Any) -> QWidget | None:
    """Resolve a parent argument to a QWidget for dialog parenting."""
    if parent is None:
        return None
    if isinstance(parent, QWidget):
        return parent
    root = getattr(parent, "root", None)
    if isinstance(root, QWidget):
        return root
    if hasattr(parent, "metaObject"):
        return parent
    return None


def _active_window_parent(parent: Any) -> QWidget | None:
    """Resolve parent, falling back to the active application window."""
    qt_parent = _resolve_parent(parent)
    if qt_parent is not None:
        return qt_parent
    app = QApplication.instance()
    if app is None:
        return None
    active = _cast(QApplication, app).activeWindow()
    return active if isinstance(active, QWidget) else None


def _focus_snapshot(parent: QWidget | None) -> QWidget | None:
    app = QApplication.instance()
    focus_widget = _cast(QApplication, app).focusWidget() if app is not None else None
    if focus_widget is None:
        return None
    if parent is None or focus_widget.window() is parent.window():
        return focus_widget
    return None


def _restore_focus(widget: QWidget | None) -> None:
    if widget is None:
        return
    try:
        if widget.isVisible() and widget.isEnabled():
            widget.setFocus(Qt.FocusReason.OtherFocusReason)
    except RuntimeError:
        return


def _message_box(
    level: FeedbackLevel,
    title: str,
    message: str,
    *,
    details: str | None = None,
    parent: Any = None,
    copyable: bool = False,
) -> None:
    """Show a QMessageBox with the given level, title, and message."""
    icon_map = {
        "info": QMessageBox.Icon.Information,
        "success": QMessageBox.Icon.Information,
        "warning": QMessageBox.Icon.Warning,
        "error": QMessageBox.Icon.Critical,
    }
    box = QMessageBox(_active_window_parent(parent))
    box.setObjectName(f"feedback{level.title()}MessageBox")
    box.setWindowTitle(title)
    box.setIcon(icon_map.get(level, QMessageBox.Icon.Information))
    box.setText(message)
    if details:
        box.setDetailedText(details)
    box.setStandardButtons(QMessageBox.StandardButton.Ok)
    copy_button = None
    copy_details = details if copyable and details else None
    if copy_details is not None:
        copy_button = box.addButton(t("common.copy", "Copy"), QMessageBox.ButtonRole.ActionRole)
        copy_button.setProperty("feedbackRole", "copy")
    box.exec()
    if copy_button is not None and box.clickedButton() is copy_button and copy_details is not None:
        _copy_text_to_clipboard(copy_details)


def _copy_text_to_clipboard(text: str) -> bool:
    app = QApplication.instance()
    if app is None:
        return False
    clipboard = _cast(QApplication, app).clipboard()
    clipboard.setText(text)
    return True


def error(
    title: str,
    message: str,
    *,
    details: str | None = None,
    parent: Any = None,
    copyable: bool = True,
) -> None:
    """Show an error dialog."""
    _message_box("error", title, message, details=details, parent=parent, copyable=copyable)


def warn(
    title: str,
    message: str,
    *,
    details: str | None = None,
    parent: Any = None,
) -> None:
    """Show a warning dialog."""
    _message_box("warning", title, message, details=details, parent=parent)


def info(
    title: str,
    message: str,
    *,
    details: str | None = None,
    parent: Any = None,
) -> None:
    """Show an info dialog."""
    _message_box("info", title, message, details=details, parent=parent)


def choose(
    title: str,
    message: str,
    *,
    choices: list[FeedbackChoice],
    parent: Any = None,
    default: str | None = None,
    details: str | None = None,
    copyable: bool = False,
    danger: bool = False,
    level: FeedbackLevel = "info",
    _object_name: str = "feedbackChoiceMessageBox",
) -> str | None:
    """Show a choice dialog and return the selected choice value."""
    qt_parent = _active_window_parent(parent)
    focus_widget = _focus_snapshot(qt_parent)

    icon_map = {
        "info": QMessageBox.Icon.Information,
        "success": QMessageBox.Icon.Information,
        "warning": QMessageBox.Icon.Warning,
        "error": QMessageBox.Icon.Critical,
    }
    role_map = {
        "accept": QMessageBox.ButtonRole.AcceptRole,
        "reject": QMessageBox.ButtonRole.RejectRole,
        "action": QMessageBox.ButtonRole.ActionRole,
    }

    box = QMessageBox(qt_parent)
    box.setObjectName(_object_name)
    box.setWindowTitle(title)
    box.setIcon(QMessageBox.Icon.Warning if danger else icon_map.get(level, QMessageBox.Icon.Information))
    box.setText(message)
    if details:
        box.setDetailedText(details)

    button_values: dict[object, str] = {}
    copy_button = None
    copy_details = details if copyable and details else None
    if copy_details is not None:
        copy_button = box.addButton(t("common.copy", "Copy"), QMessageBox.ButtonRole.ActionRole)
        copy_button.setProperty("feedbackRole", "copy")

    default_button = None
    primary_button = None
    for choice in choices:
        button = box.addButton(choice.label, role_map[choice.role])
        button.setProperty("feedbackRole", choice.role)
        if danger and choice.role == "accept":
            button.setProperty("feedbackDanger", True)
        button_values[button] = choice.value
        if choice.value == default:
            default_button = button
        if choice.primary:
            primary_button = button

    if default_button is not None:
        box.setDefaultButton(default_button)
    elif primary_button is not None:
        box.setDefaultButton(primary_button)

    try:
        box.exec()
        clicked = box.clickedButton()
        if copy_button is not None and clicked is copy_button and copy_details is not None:
            _copy_text_to_clipboard(copy_details)
            return None
        return button_values.get(clicked)
    finally:
        _restore_focus(focus_widget)
        box.deleteLater()


def confirm(
    title: str,
    message: str,
    *,
    danger: bool = False,
    parent: Any = None,
    details: str | None = None,
    default: str | None = None,
    confirm_label: str | None = None,
    cancel_label: str | None = None,
) -> bool:
    """Show a Yes/No confirmation dialog. Returns True if confirmed."""
    safe_default = default
    if safe_default is None:
        safe_default = "cancel" if danger else "confirm"
    if safe_default == "yes":
        safe_default = "confirm"
    elif safe_default == "no":
        safe_default = "cancel"

    choice = choose(
        title,
        message,
        choices=[
            FeedbackChoice(
                "confirm",
                confirm_label or t("common.ok", "OK"),
                role="accept",
                primary=safe_default == "confirm",
            ),
            FeedbackChoice(
                "cancel",
                cancel_label or t("common.cancel", "Cancel"),
                role="reject",
                primary=safe_default == "cancel",
            ),
        ],
        parent=parent,
        default=safe_default,
        details=details,
        danger=danger,
        level="warning" if danger else "info",
        _object_name="feedbackConfirmMessageBox",
    )
    return choice == "confirm"


_NOTIFY_DURATION_MS: dict[str, int] = {
    "info": 3000,
    "success": 2500,
    "warning": 4500,
    "error": 5000,
}


def notify(
    level: FeedbackLevel,
    title: str,
    message: str,
    *,
    parent: Any = None,
    duration_ms: int | None = None,
) -> None:
    """Show a transient notification toast using qfluentwidgets InfoBar.

    Falls back to a simple message box if qfluentwidgets is unavailable.
    """
    qt_parent = _active_window_parent(parent)
    duration = _NOTIFY_DURATION_MS.get(level, 3000) if duration_ms is None else duration_ms

    try:
        from qfluentwidgets import InfoBar, InfoBarPosition

        factory = {
            "info": InfoBar.info,
            "success": InfoBar.success,
            "warning": InfoBar.warning,
            "error": InfoBar.error,
        }.get(level, InfoBar.info)

        factory(
            title=title,
            content=message,
            duration=duration,
            position=InfoBarPosition.TOP_RIGHT,
            parent=qt_parent,
        )
    except ImportError:
        # Fallback: use a small auto-closing message if InfoBar is not available
        if qt_parent is not None:
            _message_box(level, title, message, parent=qt_parent)


def exception(
    exc: BaseException,
    *,
    context: str | None = None,
    parent: Any = None,
) -> None:
    """Show an error dialog with formatted traceback details."""
    title = t("common.error", "Error")
    message = str(exc) or exc.__class__.__name__
    if context:
        message = f"{context}\n{message}"
    details = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    error(title, message, details=details, parent=parent)


__all__ = [
    "FeedbackChoice",
    "FeedbackLevel",
    "choose",
    "confirm",
    "error",
    "exception",
    "info",
    "notify",
    "warn",
]
