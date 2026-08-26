"""Output settings tab — directory mode, date subfolder, auto-open.

Matches old OutputTab (6 items: output_mode, custom_path+browse, date_subfolder,
date_format, auto_open, save_intermediate).
"""

from __future__ import annotations

from typing import cast as _cast

from PySide6.QtWidgets import QCheckBox, QComboBox, QFileDialog, QHBoxLayout, QLineEdit, QPushButton, QWidget

from ...i18n import t
from ...view_models.settings_vm import SECTION_OUTPUT, SettingsViewModel
from .base_tab import BaseSettingsTab


class OutputTab(BaseSettingsTab):
    """Output directory and behavior settings tab."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        self._vm = view_model
        self._output_mode: QComboBox = _cast(QComboBox, None)
        self._custom_path: QLineEdit = _cast(QLineEdit, None)
        self._create_date_subfolder: QCheckBox = _cast(QCheckBox, None)
        self._date_format: QComboBox = _cast(QComboBox, None)
        self._auto_open_folder: QCheckBox = _cast(QCheckBox, None)
        self._save_intermediate: QCheckBox = _cast(QCheckBox, None)
        super().__init__()
        self._load_values()

    def _create_interface(self) -> None:
        # ── Intermediate files card ─────────────────────────────────────
        _g1, f1 = self.add_settings_card(
            t("settings.output.intermediate_section", "Intermediate Files"),
            t("settings.output.save_intermediate_desc", "Save intermediate processing files alongside final output."),
            object_name="outputIntermediateCard",
        )
        self._save_intermediate = self.create_settings_toggle(
            t("settings.output.save_intermediate_label", "Save intermediate files to output"),
            t("settings.output.save_intermediate_tooltip", "Retain temporary files generated during conversion"),
        )
        self.add_form_row(f1, "", self._save_intermediate)
        self._save_intermediate.toggled.connect(
            lambda v: self._vm.set_field(SECTION_OUTPUT, "save_intermediate_files", v)
        )

        # ── Output directory card ───────────────────────────────────────
        _g2, f2 = self.add_settings_card(
            t("settings.output.directory_section", "Output Directory"),
            t("settings.output.output_mode_tooltip", "Where to save converted files."),
            object_name="outputDirectoryCard",
        )
        self._output_mode = self.create_combobox(
            [
                (t("settings.output.output_modes.source", "Same as Source"), "source"),
                (t("settings.output.output_modes.custom", "Custom Directory"), "custom"),
            ],
            t("settings.output.output_mode_tooltip", "Output directory mode"),
        )
        self.add_form_row(f2, t("settings.output.output_mode_label", "Directory Mode:"), self._output_mode)
        self._output_mode.currentIndexChanged.connect(
            lambda _: self._vm.set_field(SECTION_OUTPUT, "output_mode", self.get_combo_data(self._output_mode))
        )

        path_row = QWidget()
        path_row.setObjectName("outputCustomPathRow")
        path_layout = QHBoxLayout(path_row)
        path_layout.setContentsMargins(0, 0, 0, 0)
        path_layout.setSpacing(4)
        self._custom_path = QLineEdit()
        self._custom_path.setObjectName("outputCustomPathEdit")
        browse_btn = QPushButton(t("common.browse", "Browse"))
        browse_btn.setObjectName("outputBrowseButton")
        browse_btn.clicked.connect(self._browse_path)
        self._custom_path.textChanged.connect(lambda t: self._vm.set_field(SECTION_OUTPUT, "custom_path", t))
        path_layout.addWidget(self._custom_path)
        path_layout.addWidget(browse_btn)
        self.add_form_row(f2, t("settings.output.custom_path_label", "Custom Path:"), path_row)

        # ── Date subfolder card ─────────────────────────────────────────
        _g3, f3 = self.add_settings_card(
            t("settings.output.date_folder.section", "Date Subfolder"),
            t("settings.output.date_folder.create_tooltip", "Create a date-named subfolder for output files."),
            object_name="outputDateFolderCard",
        )
        self._create_date_subfolder = self.create_settings_toggle(
            t("settings.output.date_folder.create_label", "Create date subfolder"),
            t("settings.output.date_folder.create_tooltip", "Place output files in a folder named by date"),
        )
        self.add_form_row(f3, "", self._create_date_subfolder)
        self._create_date_subfolder.toggled.connect(
            lambda v: self._vm.set_field(SECTION_OUTPUT, "create_date_subfolder", v)
        )

        self._date_format = self.create_combobox(
            [
                (t("settings.output.date_formats.iso", "ISO (YYYY-MM-DD)"), "%Y-%m-%d"),
                (t("settings.output.date_formats.compact", "Compact (YYYYMMDD)"), "%Y%m%d"),
                (t("settings.output.date_formats.chinese", "Chinese (YYYY-MM-DD)"), "%Y年%m月%d日"),
            ],
            t("settings.output.date_folder.format_label", "Date format for subfolder naming"),
        )
        self.add_form_row(f3, t("settings.output.date_folder.format_label", "Date Format:"), self._date_format)
        self._date_format.currentIndexChanged.connect(
            lambda _: self._vm.set_field(SECTION_OUTPUT, "date_folder_format", self.get_combo_data(self._date_format))
        )

        # ── Behavior card ───────────────────────────────────────────────
        _g4, f4 = self.add_settings_card(
            t("settings.output.behavior.section", "Output Behavior"),
            t("settings.output.behavior.auto_open_folder_tooltip", "Additional output actions."),
            object_name="outputBehaviorCard",
        )
        self._auto_open_folder = self.create_settings_toggle(
            t("settings.output.behavior.auto_open_folder_label", "Auto-open output folder after conversion"),
            t(
                "settings.output.behavior.auto_open_folder_tooltip",
                "Open the output directory in file manager when done",
            ),
        )
        self.add_form_row(f4, "", self._auto_open_folder)
        self._auto_open_folder.toggled.connect(lambda v: self._vm.set_field(SECTION_OUTPUT, "auto_open_folder", v))

    def _browse_path(self) -> None:
        path = QFileDialog.getExistingDirectory(self, t("settings.output.browse_title", "Select Output Directory"))
        if path and self._custom_path is not None:
            self._custom_path.setText(path)

    def _load_values(self) -> None:
        out = self._vm.config.output
        self.set_combo_data(self._output_mode, out.output_mode)
        if self._custom_path is not None:
            self._custom_path.setText(out.custom_path)
        self._create_date_subfolder.setChecked(out.create_date_subfolder)
        self.set_combo_data(self._date_format, out.date_folder_format)
        self._auto_open_folder.setChecked(out.auto_open_folder)
        self._save_intermediate.setChecked(out.save_intermediate_files)

    def reload_from_config(self) -> None:
        self._load_values()
