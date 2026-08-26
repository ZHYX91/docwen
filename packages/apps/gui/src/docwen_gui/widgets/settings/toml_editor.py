"""TOML configuration editor widget.

Provides a raw TOML text editor with config_name switching, syntax
validation via tomlkit, and save-to-disk with config reload callback.

Placed in ``widgets/settings/`` because it is used from the settings
dialog text tab as the numbering scheme TOML raw editor.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)

from docwen_runtime.config import atomic_write_text

from ...i18n import t

logger = logging.getLogger(__name__)

# ── Combo helpers (self-contained, no old adapter imports) ───────────────


def _combobox_add_item(combo: QComboBox, display: str, data: Any, *, tooltip: str = "") -> None:
    combo.addItem(display, data)
    if tooltip:
        combo.setItemData(combo.count() - 1, tooltip, Qt.ItemDataRole.ToolTipRole)


def _combobox_current_data(combo: QComboBox) -> Any:
    return combo.currentData()


def _combobox_set_current_data(combo: QComboBox, data: Any) -> bool:
    for i in range(combo.count()):
        if combo.itemData(i) == data:
            combo.setCurrentIndex(i)
            return True
    return False


def _prepare_combobox_for_long_text(combo: QComboBox) -> None:
    from PySide6.QtWidgets import QListView

    combo.setMaxVisibleItems(20)
    view = QListView()
    view.setTextElideMode(Qt.TextElideMode.ElideMiddle)
    combo.setView(view)


# ── Widget ───────────────────────────────────────────────────────────────


class TomlEditorWidget(QWidget):
    """TOML text editor widget.

    - Supports switching ``config_name`` via a combo box populated from
      ``choices``.
    - Validates TOML syntax via ``tomlkit`` before saving.
    - Calls a ``reload_callback`` after a successful save so the
      surrounding dialog can refresh its state.
    - Shows the resolved file path below the combo.

    The widget does **not** depend on a concrete config loader directly.
    Persistence is pluggable:

    - If a ``save_callback`` ``(config_name, content) -> bool`` is
      provided, ``save_to_disk`` delegates the write to it (the callback
      is normally supplied by the Settings view model and its injected
      ``ConfigPort`` so the save uses registry validation and the configured
      in-memory source). This is the production path.
    - Otherwise it falls back to ``path_resolver`` plus the runtime's atomic
      text writer (used by tests and non-config ad-hoc editors).

    In both cases ``reload_callback`` is invoked after a successful save.
    """

    def __init__(
        self,
        parent: QWidget | None,
        *,
        config_name: str,
        choices: list[tuple[str, str]] | None = None,
        path_resolver: Any | None = None,
        reload_callback: Any | None = None,
        save_callback: Any | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("tomlEditorWidget")
        self._config_name = config_name
        self._choices = list(choices or [])
        self._path_resolver = path_resolver
        self._reload_callback = reload_callback
        self._save_callback = save_callback

        self._combo: QComboBox | None = None
        self._path_label: QLabel | None = None
        self._editor: QPlainTextEdit | None = None

        self._build_ui()
        self.reload_from_disk()

    # ── UI construction ────────────────────────────────────────────────

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(8)

        header = QHBoxLayout()
        header.setSpacing(8)

        label = QLabel(t("settings.toml_editor.config_label", "Config:") + " ", self)
        header.addWidget(label)

        if self._choices:
            combo = QComboBox(self)
            self._combo = combo
            _prepare_combobox_for_long_text(combo)
            for display, name in self._choices:
                _combobox_add_item(combo, display, name, tooltip=display)
            _combobox_set_current_data(combo, self._config_name)
            combo.currentIndexChanged.connect(self._on_config_changed)
            header.addWidget(combo, 1)
        else:
            header.addStretch(1)

        reload_btn = QDialogButtonBox(QDialogButtonBox.StandardButton.Reset, parent=self)
        rst = reload_btn.button(QDialogButtonBox.StandardButton.Reset)
        if rst is not None:
            rst.setText(t("common.reload", "Reload"))
            rst.clicked.connect(self.reload_from_disk)
        header.addWidget(reload_btn)

        root.addLayout(header)

        path_label = QLabel("", self)
        path_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self._path_label = path_label
        root.addWidget(path_label)

        editor = QPlainTextEdit(self)
        editor.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self._editor = editor
        root.addWidget(editor, 1)

        hint = QLabel(
            t("settings.toml_editor.hint", "Edit the TOML directly. Save will validate syntax."),
            self,
        )
        hint.setWordWrap(True)
        root.addWidget(hint)

    # ── Public API ─────────────────────────────────────────────────────

    @property
    def config_name(self) -> str:
        return self._config_name

    def set_config_name(self, config_name: str) -> None:
        if config_name == self._config_name:
            return
        self._config_name = config_name
        self.reload_from_disk()

    def reload_from_disk(self) -> None:
        path = self._resolve_path()
        if self._path_label is not None:
            self._path_label.setText(str(path))
        try:
            content = path.read_text(encoding="utf-8")
        except FileNotFoundError:
            content = ""
        except Exception as exc:
            logger.warning("Failed to read %s: %s", path, exc)
            content = ""
        if self._editor is not None:
            self._editor.setPlainText(content)

    def save_to_disk(self, *, show_success: bool = True) -> bool:
        if self._editor is None:
            return False
        content = self._editor.toPlainText()
        if not self._validate_toml(content):
            return False

        if callable(self._save_callback):
            # Production path: delegate to the Settings VM / injected port so
            # the write goes through registry validation and state wiring.
            try:
                ok = bool(self._save_callback(self._config_name, content))
            except Exception as exc:
                self._show_error(
                    t("settings.toml_editor.save_failed", "Save failed"),
                    str(exc),
                )
                return False
            if not ok:
                self._show_error(
                    t("settings.toml_editor.save_failed", "Save failed"),
                    t(
                        "settings.toml_editor.save_failed_message",
                        "The configuration could not be saved.",
                    ),
                )
                return False
        else:
            # Fallback path: atomic write via path_resolver (tests / ad-hoc).
            path = self._resolve_path()
            try:
                atomic_write_text(path, content)
            except Exception as exc:
                self._show_error(
                    t("settings.toml_editor.save_failed", "Save failed"),
                    str(exc),
                )
                return False

        # Notify caller that config changed
        if callable(self._reload_callback):
            try:
                reload_result = self._reload_callback()
            except Exception as exc:
                logger.warning("Reload callback failed: %s", exc)
                self._show_error(
                    t("settings.toml_editor.save_failed", "Save failed"),
                    str(exc),
                )
                return False
            if reload_result is False:
                self._show_error(
                    t("settings.toml_editor.save_failed", "Save failed"),
                    t(
                        "settings.toml_editor.save_failed_message",
                        "The configuration could not be saved.",
                    ),
                )
                return False

        if show_success:
            from ...dialogs.feedback import notify

            notify(
                "success",
                t("settings.toml_editor.save_success_title", "Saved"),
                t(
                    "settings.toml_editor.save_success_message",
                    "Configuration saved successfully.",
                ),
                parent=self,
            )
        return True

    # ── Internal ───────────────────────────────────────────────────────

    def _resolve_path(self) -> Path:
        if callable(self._path_resolver):
            return Path(str(self._path_resolver(self._config_name)))
        raise RuntimeError("path_resolver is not set")

    def _validate_toml(self, content: str) -> bool:
        try:
            from docwen_core.toml_tools import parse_toml_text

            parse_toml_text(content or "")
            return True
        except Exception as exc:
            self._show_error(
                t("settings.toml_editor.toml_syntax_error", "TOML Syntax Error"),
                str(exc),
            )
            return False

    def _show_error(self, title: str, message: str) -> None:
        from ...dialogs.feedback import error

        error(title, message, parent=self)

    def _on_config_changed(self, _index: int) -> None:
        if self._combo is None:
            return
        data = _combobox_current_data(self._combo)
        if isinstance(data, str) and data:
            self.set_config_name(data)


class TomlEditorDialog(QWidget):
    """Container widget that wraps a ``TomlEditorWidget`` with Save/Reload buttons.

    Intended to be embedded as a settings tab page or a standalone dialog.
    """

    def __init__(self, parent: QWidget | None, *, editor: TomlEditorWidget) -> None:
        super().__init__(parent)
        self.setObjectName("tomlEditorDialog")
        self.editor = editor
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        layout.addWidget(editor, 1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Reset,
            parent=self,
        )
        save_btn = buttons.button(QDialogButtonBox.StandardButton.Save)
        if save_btn is not None:
            save_btn.setText(t("common.save", "Save"))
            save_btn.clicked.connect(lambda: self.editor.save_to_disk(show_success=True))
        reset_btn = buttons.button(QDialogButtonBox.StandardButton.Reset)
        if reset_btn is not None:
            reset_btn.setText(t("common.reload", "Reload"))
            reset_btn.clicked.connect(self.editor.reload_from_disk)
        layout.addWidget(buttons)
