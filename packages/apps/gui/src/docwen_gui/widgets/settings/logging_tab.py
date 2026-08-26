"""Logging settings tab — levels, file prefix, retention, directory, console.

Matches old LoggingTab: 11 config items + 4 read-only runtime state fields +
copy path / open directory buttons.
"""

from __future__ import annotations

import os
import re
import tempfile
from pathlib import Path
from typing import cast as _cast

from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QWidget,
)

from ...i18n import t
from ...styles.theme_semantics import apply_theme_class
from ...view_models.settings_vm import SECTION_LOGGING, SettingsViewModel
from .base_tab import BaseSettingsTab, _create_info_button


def _log_levels() -> list[tuple[str, str]]:
    return [
        (t("settings.logging.levels.debug", "Debug"), "debug"),
        (t("settings.logging.levels.info", "Info"), "info"),
        (t("settings.logging.levels.warning", "Warning"), "warning"),
        (t("settings.logging.levels.error", "Error"), "error"),
        (t("settings.logging.levels.critical", "Critical"), "critical"),
    ]


def _dir_modes() -> list[tuple[str, str]]:
    return [
        (t("settings.logging.dir_modes.user", "System Default"), "user"),
        (t("settings.logging.dir_modes.temp", "Temp Directory"), "temp"),
        (t("settings.logging.dir_modes.custom", "Custom Directory"), "custom"),
    ]


def _console_colorize_modes() -> list[tuple[str, str]]:
    return [
        (t("settings.logging.console_colorize.auto", "Auto"), "auto"),
        (t("settings.logging.console_colorize.always", "Always"), "always"),
        (t("settings.logging.console_colorize.never", "Never"), "never"),
    ]


class LoggingTab(BaseSettingsTab):
    """Logging system settings tab."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        self._vm = view_model
        self._enable_checkbox: QCheckBox = _cast(QCheckBox, None)
        self._level_combo: QComboBox = _cast(QComboBox, None)
        self._file_prefix_edit: QLineEdit = _cast(QLineEdit, None)
        self._retention_days: QSpinBox = _cast(QSpinBox, None)
        self._console_enable: QCheckBox = _cast(QCheckBox, None)
        self._console_level: QComboBox = _cast(QComboBox, None)
        self._console_format: QLineEdit = _cast(QLineEdit, None)
        self._console_colorize: QComboBox = _cast(QComboBox, None)
        self._dir_mode: QComboBox = _cast(QComboBox, None)
        self._dir_edit: QLineEdit = _cast(QLineEdit, None)
        self._browse_btn: QPushButton = _cast(QPushButton, None)
        self._resolved_path: QLineEdit = _cast(QLineEdit, None)
        self._fallback_reason: QLineEdit = _cast(QLineEdit, None)
        self._override_source: QLineEdit = _cast(QLineEdit, None)
        self._dir_notice: QLineEdit = _cast(QLineEdit, None)
        super().__init__()
        self._load_values()

    def _create_interface(self) -> None:
        # ── System card ─────────────────────────────────────────────────
        _c1, f1 = self.add_settings_card(
            t("settings.logging.system_section", "System"),
            t("settings.logging.system_desc", "Enable or disable logging and set log level."),
        )

        toggle_widget, self._enable_checkbox = self._create_toggle_with_info(
            t("settings.logging.enable_logging_label", "Enable Logging"),
            t("settings.logging.enable_logging_tooltip", "Enable file and console logging"),
        )
        f1.addRow(toggle_widget)
        self._enable_checkbox.toggled.connect(lambda v: self._vm.set_field(SECTION_LOGGING, "enable", v))

        self._level_combo = self.create_combobox(_log_levels(), t("settings.logging.log_level_tooltip", "Log level"))
        self.add_form_row(f1, t("settings.logging.log_level_label", "Log Level:"), self._level_combo)
        self._level_combo.currentIndexChanged.connect(
            lambda _: self._vm.set_field(SECTION_LOGGING, "level", self.get_combo_data(self._level_combo))
        )

        # ── File card ───────────────────────────────────────────────────
        _c2, f2 = self.add_settings_card(
            t("settings.logging.file_section", "File Logging"),
            t("settings.logging.file_desc", "Configure log file naming and retention."),
        )
        self._file_prefix_edit = QLineEdit(self._scroll_container)
        self._file_prefix_edit.setToolTip(
            t("settings.logging.file_prefix_tooltip", "Log file name prefix (no spaces or special chars)")
        )
        self._file_prefix_edit.textChanged.connect(self._on_file_prefix_changed)
        self.add_form_row(f2, t("settings.logging.file_prefix_label", "File Prefix:"), self._file_prefix_edit)

        self._retention_days = self.create_spinbox(
            1, 365, t("settings.logging.retention_days_tooltip", "Days to keep log files"), default=30
        )
        self._retention_days.valueChanged.connect(lambda v: self._vm.set_field(SECTION_LOGGING, "retention_days", v))
        self.add_form_row(f2, t("settings.logging.retention_days_label", "Retention (days):"), self._retention_days)

        # ── Location card ───────────────────────────────────────────────
        _c3, f3 = self.add_settings_card(
            t("settings.logging.location_section", "Output Location"),
            t("settings.logging.location_desc", "Control where log files are written."),
        )
        self._dir_mode = self.create_combobox(
            _dir_modes(), t("settings.logging.directory_mode_tooltip", "Directory mode")
        )
        self._dir_mode.currentIndexChanged.connect(self._on_dir_mode_changed)
        self.add_form_row(f3, t("settings.logging.directory_mode_label", "Directory Mode:"), self._dir_mode)

        dir_row = QWidget(self._scroll_container)
        dir_layout = QHBoxLayout(dir_row)
        dir_layout.setContentsMargins(0, 0, 0, 0)
        dir_layout.setSpacing(8)
        self._dir_edit = QLineEdit(self._scroll_container)
        self._dir_edit.setToolTip(
            t("settings.logging.custom_directory_tooltip", "Custom log directory (only used in Custom mode)")
        )
        self._dir_edit.textChanged.connect(self._on_directory_changed)
        dir_layout.addWidget(self._dir_edit, 1)
        self._browse_btn = QPushButton(t("common.browse", "Browse"), self._scroll_container)
        self._browse_btn.clicked.connect(self._browse_directory)
        dir_layout.addWidget(self._browse_btn)
        self.add_form_row(f3, t("settings.logging.custom_directory_label", "Custom Directory:"), dir_row)

        # Read-only runtime state fields
        self._resolved_path = QLineEdit(self._scroll_container)
        self._resolved_path.setReadOnly(True)
        self.add_form_row(f3, t("settings.logging.actual_log_file_label", "Actual Log File:"), self._resolved_path)

        self._fallback_reason = QLineEdit(self._scroll_container)
        self._fallback_reason.setReadOnly(True)
        self.add_form_row(f3, t("settings.logging.fallback_reason_label", "Fallback Reason:"), self._fallback_reason)

        self._override_source = QLineEdit(self._scroll_container)
        self._override_source.setReadOnly(True)
        self.add_form_row(f3, t("settings.logging.override_source_label", "Override Source:"), self._override_source)

        self._dir_notice = QLineEdit(self._scroll_container)
        self._dir_notice.setReadOnly(True)
        self.add_form_row(f3, t("settings.logging.directory_notice_label", "Directory Notice:"), self._dir_notice)

        btn_row = QWidget(self._scroll_container)
        btn_layout = QHBoxLayout(btn_row)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(8)
        copy_btn = QPushButton(t("info_area.copy_path", "Copy path"), btn_row)
        copy_btn.setObjectName("settingsLoggingCopyPathButton")
        apply_theme_class(copy_btn, "secondary")
        copy_btn.clicked.connect(self._copy_path)
        btn_layout.addWidget(copy_btn)
        open_btn = QPushButton(t("settings.logging.open_directory", "Open Directory"), btn_row)
        open_btn.setObjectName("settingsLoggingOpenDirectoryButton")
        apply_theme_class(open_btn, "secondary")
        open_btn.clicked.connect(self._open_directory)
        btn_layout.addWidget(open_btn)
        btn_layout.addStretch(1)
        f3.addRow(btn_row)

        # ── Console card ────────────────────────────────────────────────
        _c4, f4 = self.add_settings_card(
            t("settings.logging.console_section", "Console"),
            t("settings.logging.console_desc", "Control console log output and formatting."),
        )
        console_toggle, self._console_enable = self._create_toggle_with_info(
            t("settings.logging.enable_console_label", "Enable Console Logging"),
            t("settings.logging.enable_console_tooltip", "Print log messages to console"),
        )
        f4.addRow(console_toggle)
        self._console_enable.toggled.connect(lambda v: self._vm.set_field(SECTION_LOGGING, "console_enable", v))

        self._console_level = self.create_combobox(
            _log_levels(), t("settings.logging.console_level_tooltip", "Console log level")
        )
        self.add_form_row(f4, t("settings.logging.console_level_label", "Console Level:"), self._console_level)
        self._console_level.currentIndexChanged.connect(
            lambda _: self._vm.set_field(SECTION_LOGGING, "console_level", self.get_combo_data(self._console_level))
        )

        self._console_format = QLineEdit(self._scroll_container)
        self._console_format.setToolTip(
            t("settings.logging.console_format_tooltip", "Console format string (empty = reuse file format)")
        )
        self._console_format.textChanged.connect(lambda t: self._vm.set_field(SECTION_LOGGING, "console_format", t))
        self.add_form_row(f4, t("settings.logging.console_format_label", "Console Format:"), self._console_format)

        self._console_colorize = self.create_combobox(
            _console_colorize_modes(),
            t("settings.logging.console_colorize_tooltip", "Control console log color policy."),
        )
        self._console_colorize.currentIndexChanged.connect(
            lambda _: self._vm.set_field(
                SECTION_LOGGING, "console_colorize", self.get_combo_data(self._console_colorize)
            )
        )
        self.add_form_row(
            f4,
            t("settings.logging.console_colorize_label", "Console color:"),
            self._console_colorize,
        )

    def _create_toggle_with_info(self, text: str, tooltip: str) -> tuple[QWidget, QCheckBox]:
        container = QWidget(self._scroll_container)
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        cb = self.create_settings_toggle(text, tooltip)
        layout.addWidget(cb)
        if tooltip:
            layout.addWidget(_create_info_button(tooltip, container))
        layout.addStretch(1)
        return container, cb

    def _on_file_prefix_changed(self, value: str) -> None:
        self._vm.set_field(SECTION_LOGGING, "file_prefix", value)
        self._refresh_runtime_state()

    def _on_directory_changed(self, value: str) -> None:
        self._vm.set_field(SECTION_LOGGING, "directory", value)
        self._refresh_runtime_state()

    def _on_dir_mode_changed(self, _idx: int) -> None:
        mode = self.get_combo_data(self._dir_mode)
        if mode:
            self._vm.set_field(SECTION_LOGGING, "directory_mode", mode)
        self._update_dir_controls()

    def _env_override_source(self) -> str:
        if os.environ.get("DOCWEN_LOG_DIR", "").strip():
            return "DOCWEN_LOG_DIR"
        if os.environ.get("DOCWEN_LOG_TO_TEMP", "").strip().lower() in {"1", "true", "yes", "on"}:
            return "DOCWEN_LOG_TO_TEMP"
        return ""

    def _update_dir_controls(self, *, refresh_runtime: bool = True) -> None:
        is_custom = self.get_combo_data(self._dir_mode) == "custom"
        override = bool(self._env_override_source())
        if self._override_source is not None:
            override = override or bool(self._override_source.text().strip())
        if self._dir_mode:
            self._dir_mode.setEnabled(not override)
        enabled = is_custom and not override
        if self._dir_edit:
            self._dir_edit.setEnabled(enabled)
        if self._browse_btn:
            self._browse_btn.setEnabled(enabled)
        if refresh_runtime:
            self._refresh_runtime_state()

    def _browse_directory(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            t("settings.logging.browse_title", "Select Log Directory"),
            (self._dir_edit.text() or str(Path.home())) if self._dir_edit else str(Path.home()),
        )
        if selected and self._dir_edit is not None:
            self._dir_edit.setText(selected)
            self._vm.set_field(SECTION_LOGGING, "directory", selected)
            self._refresh_runtime_state()

    def _copy_path(self) -> None:
        if self._resolved_path:
            QApplication.clipboard().setText(self._resolved_path.text())

    def _open_directory(self) -> None:
        if self._resolved_path is None:
            return
        target = self._resolved_path.text().strip()
        if not target:
            return
        import subprocess
        import sys

        parent_dir = str(Path(target).parent)
        if sys.platform == "win32":
            subprocess.Popen(["explorer", parent_dir])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", parent_dir])
        else:
            subprocess.Popen(["xdg-open", parent_dir])

    def _refresh_runtime_state(self) -> None:
        """Display runtime log file path and env override source."""
        prefix = self._file_prefix_edit.text().strip() if self._file_prefix_edit else "docwen"
        dir_mode = self.get_combo_data(self._dir_mode) if self._dir_mode else "user"
        custom_dir = self._dir_edit.text() if self._dir_edit else ""

        env_override = self._env_override_source()
        fallback_reason = ""
        config = {
            "file_prefix": prefix,
            "directory_mode": dir_mode,
            "directory": custom_dir,
        }
        try:
            from docwen_runtime.logging import (
                get_logging_runtime_state,
                resolve_log_file_path,
            )

            runtime_state = get_logging_runtime_state()
            env_override = runtime_state.overridden_by_env or env_override
            if env_override and runtime_state.active_log_file:
                resolved = runtime_state.active_log_file
            else:
                resolved = resolve_log_file_path(config)
            fallback_reason = runtime_state.fallback_reason or ""
        except Exception:
            resolved = self._fallback_resolve_log_file_path(config)

        if self._resolved_path:
            self._resolved_path.setText(resolved)
        if self._fallback_reason:
            self._fallback_reason.setText(fallback_reason)
        if self._override_source:
            self._override_source.setText(env_override)

        notice = ""
        if env_override:
            notice = t(
                "settings.logging.directory_env_override_notice",
                "Log directory is controlled by an environment variable; directory controls are disabled.",
            )
        elif dir_mode != "custom":
            notice = t(
                "settings.logging.directory_auto_notice",
                "Directory is auto-determined by mode; custom input is disabled.",
            )
        if self._dir_notice:
            self._dir_notice.setText(notice)
        self._update_dir_controls(refresh_runtime=False)

    @staticmethod
    def _fallback_resolve_log_file_path(config: dict[str, object]) -> str:
        prefix = re.sub(r'[\\/*?:"<>|\x00-\x1f]', "", str(config.get("file_prefix") or ""))
        prefix = prefix.strip().rstrip(". ") or "docwen"
        dir_mode = str(config.get("directory_mode") or "user").strip().lower()
        if dir_mode not in {"user", "temp", "custom"}:
            dir_mode = "user"
        custom_dir = str(config.get("directory") or "")
        if dir_mode == "temp":
            log_dir = str(Path(tempfile.gettempdir()) / "docwen" / "logs")
        elif dir_mode == "custom" and custom_dir:
            log_dir = custom_dir
        else:
            log_dir = str(Path.home() / ".docwen" / "logs")
        resolved_dir = Path(log_dir).resolve(strict=False)
        resolved_path = (resolved_dir / f"{prefix}.log").resolve(strict=False)
        try:
            resolved_path.relative_to(resolved_dir)
        except ValueError:  # pragma: no cover - defense after prefix normalization
            return str(resolved_dir / "docwen.log")
        return str(resolved_path)

    def _load_values(self) -> None:
        log = self._vm.config.logging
        self._enable_checkbox.setChecked(log.enable)
        self.set_combo_data(self._level_combo, log.level)
        if self._file_prefix_edit:
            self._file_prefix_edit.setText(log.file_prefix)
        self._retention_days.setValue(log.retention_days)
        self._console_enable.setChecked(log.console_enable)
        self.set_combo_data(self._console_level, log.console_level)
        if self._console_format:
            self._console_format.setText(log.console_format)
        self.set_combo_data(self._console_colorize, log.console_colorize)
        self.set_combo_data(self._dir_mode, log.directory_mode)
        if self._dir_edit:
            self._dir_edit.setText(log.directory)
        self._update_dir_controls()

    def reload_from_config(self) -> None:
        self._load_values()

    def validate(self) -> list[str]:
        """Validate logging-specific inputs."""
        errors: list[str] = []
        prefix = self._file_prefix_edit.text().strip() if self._file_prefix_edit else ""
        if not prefix:
            errors.append(t("settings.logging.validation_prefix_required", "Log file prefix cannot be empty."))
        elif re.search(r'[\\/*?:"<>|]', prefix):
            errors.append(
                t(
                    "settings.logging.validation_prefix_invalid",
                    'Log file prefix contains invalid characters: \\ / : * ? " < > |',
                )
            )

        dir_mode = self.get_combo_data(self._dir_mode) if self._dir_mode else "user"
        override = bool((self._override_source.text() or "").strip()) if self._override_source else False
        custom_dir = self._dir_edit.text().strip() if self._dir_edit else ""
        if dir_mode == "custom" and not override and not custom_dir:
            errors.append(
                t(
                    "settings.logging.validation_custom_directory_required",
                    "Custom directory mode requires a non-empty directory path.",
                )
            )
        return errors
