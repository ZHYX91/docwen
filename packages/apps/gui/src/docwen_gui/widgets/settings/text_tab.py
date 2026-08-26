"""Text settings tab — numbering schemes, templates, field processors.

Matches old TextTab behavior:
- Remove/add numbering checkboxes with scheme combo (linked state)
- Numbering scheme editor + pattern editor buttons
- MD default template (docx/xlsx)
- Field processor dynamic checkbox list
"""

from __future__ import annotations

from collections.abc import Mapping
from typing import cast as _cast

from PySide6.QtCore import QSignalBlocker
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from docwen_gui import numbering_schemes

from ...i18n import t
from ...view_models.settings_vm import SECTION_GUI, SECTION_TEXT, SettingsViewModel
from ..template_selector_tabbed import TabbedTemplateSelector
from .base_tab import BaseSettingsTab


def _draft_mapping(value: object) -> dict[str, object]:
    """Project a dynamically assigned Settings draft value as a plain mapping."""
    if not isinstance(value, Mapping):
        return {}
    return {str(key): item for key, item in value.items()}


# ── TextTab ─────────────────────────────────────────────────────────────────


class TextTab(BaseSettingsTab):
    """Text/numbering/template settings tab."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        self._vm = view_model
        self._remove_numbering: QCheckBox = _cast(QCheckBox, None)
        self._add_numbering: QCheckBox = _cast(QCheckBox, None)
        self._scheme_combo: QComboBox = _cast(QComboBox, None)
        self._template_combo: QComboBox = _cast(QComboBox, None)
        self._template_selector: TabbedTemplateSelector = _cast(TabbedTemplateSelector, None)
        self._field_processor_checkboxes: dict[str, QCheckBox] = {}
        self._field_processors_by_id: dict[str, dict] = {}
        self._field_processors_empty_label: QLabel = _cast(QLabel, None)
        super().__init__()
        self._load_values()

    def _create_interface(self) -> None:
        # ── MD to DOCX numbering card ───────────────────────────────────
        _card1, form1 = self.add_settings_card(
            t("settings.text.md_to_docx_section", "MD to DOCX — Numbering Options"),
            object_name="textNumberingCard",
        )
        self._remove_numbering = self.create_settings_toggle(
            t("settings.text.remove_numbering", "Remove existing numbering from headings"),
            t("settings.text.remove_numbering_tooltip", "Strip heading numbers when converting MD to DOCX"),
        )
        self.add_form_row(form1, "", self._remove_numbering)

        self._add_numbering = self.create_settings_toggle(
            t("settings.text.add_numbering", "Add numbering to headings"),
            t("settings.text.add_numbering_tooltip", "Add hierarchical numbering to headings in the output DOCX"),
        )
        self.add_form_row(form1, "", self._add_numbering)

        # Numbering scheme combo (enabled only when add_numbering is checked)
        self._scheme_combo = self.create_combobox(
            numbering_schemes.get_numbering_scheme_items(
                config_data=self._vm.config.text.numbering_schemes,
            ),
            t("settings.text.scheme_tooltip", "Choose numbering scheme for heading numbering"),
        )
        self._add_numbering.stateChanged.connect(lambda state: self._scheme_combo.setEnabled(bool(state)))
        self._remove_numbering.stateChanged.connect(self._on_remove_numbering_changed)
        self._add_numbering.stateChanged.connect(self._on_add_numbering_changed)
        self._scheme_combo.currentIndexChanged.connect(self._on_scheme_changed)
        self.add_form_row(form1, t("settings.text.scheme_label", "Numbering Scheme:"), self._scheme_combo)

        # Heading numbering output mode (render mode)
        self._render_mode_combo = self.create_combobox(
            [
                (t("settings.text.render_mode_text", "Text"), "text"),
                (t("settings.text.render_mode_word_native", "Word native"), "word_native"),
            ],
            t("settings.text.render_mode_tooltip", "Heading numbering output mode"),
        )
        self._render_mode_combo.setEnabled(self._add_numbering.isChecked())
        self._add_numbering.stateChanged.connect(lambda state: self._render_mode_combo.setEnabled(bool(state)))
        self._render_mode_combo.currentIndexChanged.connect(self._on_render_mode_changed)
        self.add_form_row(
            form1,
            t("settings.text.render_mode_label", "Heading numbering output mode:"),
            self._render_mode_combo,
        )

        # Help text
        help_label = QLabel(
            t(
                "settings.text.render_mode_hint",
                "Text = write numbers into headings; Word native = use Word multilevel list numbering",
            )
        )
        help_label.setWordWrap(True)
        help_label.setProperty("class", "secondary")
        form1.addRow(help_label)

        # ── Numbering editor buttons ────────────────────────────────────
        _card2, form2 = self.add_settings_card(
            t("settings.text.numbering_settings_section", "Numbering Settings"),
            t("settings.text.numbering_settings_desc", "Edit numbering addition schemes and removal rules."),
            object_name="textNumberingEditorsCard",
        )
        button_row = QWidget(self)
        button_layout = QVBoxLayout(button_row)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(8)

        add_btn = QPushButton(t("settings.text.edit_numbering_add", "Edit Numbering Addition Schemes"), button_row)
        add_btn.clicked.connect(self._open_numbering_scheme_editor)
        button_layout.addWidget(add_btn)

        clean_btn = QPushButton(t("settings.text.edit_numbering_clean", "Edit Numbering Removal Rules"), button_row)
        clean_btn.clicked.connect(self._open_numbering_clean_editor)
        button_layout.addWidget(clean_btn)

        form2.addRow(button_row)

        # ── Field processors card ─────────────────────────────────────
        _card_fp, self._field_processors_form = self.add_settings_card(
            t("settings.text.field_processors_section", "Field Optimizations (Markdown to Document)"),
            object_name="textFieldProcessorsCard",
        )
        self._field_processors_empty_label = QLabel(
            t("settings.text.field_processors_empty", "No field optimizations available for the current language."),
            self,
        )
        self._field_processors_empty_label.setWordWrap(True)
        self._field_processors_empty_label.setProperty("class", "secondary")
        self._field_processors_form.addRow(self._field_processors_empty_label)
        self._build_field_processor_rows()

        # ── Template card ───────────────────────────────────────────────
        _card3, form3 = self.add_settings_card(
            t("settings.text.template_section", "Default Template"),
            t("settings.text.template_tooltip", "Select a default template for MD to document conversions."),
            object_name="textTemplateCard",
        )
        self._template_selector = TabbedTemplateSelector(
            on_template_selected=self._on_template_selector_changed,
            on_tab_changed=self._on_template_type_changed,
        )
        self.add_form_row(form3, "", self._template_selector)
        # Connect ViewModel template signals
        self._vm.template_lists_changed.connect(self._on_templates_loaded)

    # ── Value loading ───────────────────────────────────────────────────────

    def _load_values(self) -> None:
        config = self._vm.config
        text = config.text
        self._refresh_scheme_combo_items(text.default_scheme)

        with QSignalBlocker(self._remove_numbering):
            self._remove_numbering.setChecked(text.remove_numbering)
        with QSignalBlocker(self._add_numbering):
            self._add_numbering.setChecked(text.add_numbering)
        if self._scheme_combo is not None:
            with QSignalBlocker(self._scheme_combo):
                self._scheme_combo.setEnabled(text.add_numbering)
                self.set_combo_data(self._scheme_combo, text.default_scheme)
        if hasattr(self, "_render_mode_combo") and self._render_mode_combo is not None:
            with QSignalBlocker(self._render_mode_combo):
                self._render_mode_combo.setEnabled(text.add_numbering)
                self.set_combo_data(self._render_mode_combo, text.heading_numbering_render_mode)
        if self._template_selector is not None:
            templates = self._vm.get_templates()
            if templates:
                self._template_selector.load_all_templates(templates)
            # Restore previous selections from ViewModel
            selected = self._vm.selected_templates
            if selected:
                for tt, name in selected.items():
                    sel = self._template_selector.get_selector(tt)
                    if sel is not None and sel.has_template(name):
                        sel.select_template(name, selection_source="restore")
            self._template_selector.restore_current_tab(config.gui.md_default_template)
        self._load_field_processor_values()

    def reload_from_config(self) -> None:
        self._load_values()

    def _refresh_scheme_combo_items(self, selected_scheme: str = "") -> None:
        if self._scheme_combo is None:
            return
        items = numbering_schemes.get_numbering_scheme_items(
            config_data=self._vm.config.text.numbering_schemes,
        )
        with QSignalBlocker(self._scheme_combo):
            self._scheme_combo.clear()
            for label, scheme_id in items:
                self._scheme_combo.addItem(label, scheme_id)
            if selected_scheme:
                self.set_combo_data(self._scheme_combo, selected_scheme)

    def _build_field_processor_rows(self) -> None:
        processors = self._vm.get_available_field_processors()
        self._field_processors_by_id = {str(item.get("id")): item for item in processors}
        self._field_processors_empty_label.setVisible(not processors)
        for item in processors:
            self._add_field_processor_row(item)

    def _add_field_processor_row(self, item: dict) -> None:
        processor_id = str(item.get("id") or "")
        if not processor_id or processor_id in self._field_processor_checkboxes:
            return
        name = str(item.get("name") or processor_id)
        # Field processors may come from mutable user/plugin configuration.
        # Their display name is data, not authority to address DocWen's core
        # locale catalogue with an arbitrary dynamic key.
        label = name
        description = str(item.get("description") or "")
        if item.get("load_error"):
            description = t(
                "settings.text.field_processors_load_error",
                "Load error: {error}",
                error=item.get("load_error"),
            )
        checkbox = self.create_settings_toggle(label, description, bool(item.get("enabled", True)))
        checkbox.stateChanged.connect(
            lambda state, pid=processor_id: self._on_field_processor_toggled(pid, bool(state))
        )
        self._field_processor_checkboxes[processor_id] = checkbox
        self._field_processors_form.addRow(checkbox)

    def _load_field_processor_values(self) -> None:
        processors = self._vm.get_available_field_processors()
        self._field_processors_by_id = {str(item.get("id")): item for item in processors}
        self._field_processors_empty_label.setVisible(not processors)
        for item in processors:
            self._add_field_processor_row(item)
        for processor_id, checkbox in self._field_processor_checkboxes.items():
            item = self._field_processors_by_id.get(processor_id)
            checkbox.setVisible(item is not None)
            if item is None:
                continue
            with QSignalBlocker(checkbox):
                checkbox.setChecked(bool(item.get("enabled", True)))

    # ── Signal handlers ─────────────────────────────────────────────────────

    def _on_template_selector_changed(self, template_type: str, name: str) -> None:
        """Handle template selection without dirtying on list restoration."""
        feedback = self._template_selector.peek_callback_selection_feedback()
        selection_source = feedback[2].selection_source if feedback is not None else "user"
        if selection_source == "user":
            self._vm.set_field(SECTION_GUI, "md_default_template", template_type)
        self._vm.select_template(template_type, name)

    def _on_template_type_changed(self, template_type: str, _previous_type: str) -> None:
        """Persist an explicit user tab switch independently of selection restore."""
        self._vm.set_field(SECTION_GUI, "md_default_template", template_type)

    def _on_templates_loaded(self, data: dict) -> None:
        """Load template lists into the selector widget when ViewModel provides them."""
        if self._template_selector is not None:
            self._template_selector.load_all_templates(data)
            self._template_selector.restore_current_tab(self._vm.config.gui.md_default_template)

    def _on_remove_numbering_changed(self, state: int) -> None:
        self._vm.set_field(SECTION_TEXT, "remove_numbering", bool(state))

    def _on_add_numbering_changed(self, state: int) -> None:
        enabled = bool(state)
        self._scheme_combo.setEnabled(enabled)
        self._vm.set_field(SECTION_TEXT, "add_numbering", enabled)

    def _on_scheme_changed(self, _index: int) -> None:
        scheme_id = self._scheme_combo.currentData()
        if not isinstance(scheme_id, str) or not scheme_id:
            return

        config = self._vm.config
        numbering_schemes = _draft_mapping(config.text.numbering_schemes)
        settings = _draft_mapping(numbering_schemes.get("settings"))
        settings["default_scheme"] = scheme_id
        numbering_schemes["settings"] = settings
        self._vm.set_field_batch(
            SECTION_TEXT,
            {
                "default_scheme": scheme_id,
                "numbering_schemes": numbering_schemes,
            },
        )

    def _on_render_mode_changed(self, _index: int) -> None:
        """Handle heading numbering render mode combo change."""
        mode = self._render_mode_combo.currentData()
        if not isinstance(mode, str) or not mode:
            return
        self._vm.set_field(SECTION_TEXT, "heading_numbering_render_mode", mode)

    def _on_field_processor_toggled(self, processor_id: str, enabled: bool) -> None:
        if self._vm.set_field_processor_enabled(processor_id, enabled):
            return
        checkbox = self._field_processor_checkboxes.get(processor_id)
        processor = self._field_processors_by_id.get(processor_id, {})
        name = str(processor.get("name") or processor_id)
        if checkbox is not None:
            with QSignalBlocker(checkbox):
                checkbox.setChecked(not enabled)
        from PySide6.QtWidgets import QMessageBox

        QMessageBox.warning(
            self,
            t("common.error", "Error"),
            t(
                "settings.text.field_processors_toggle_failed",
                "Failed to update “{name}”. The previous value has been restored.",
                name=name,
            ),
        )

    # ── Editor dialogs ──────────────────────────────────────────────────────

    def _open_numbering_scheme_editor(self) -> None:
        """Open the full numbering addition scheme editor (NumberingAddDialog)."""
        from .numbering_add_editor import NumberingAddDialog

        config = self._vm.config
        ns = _draft_mapping(config.text.numbering_schemes)
        settings = _draft_mapping(ns.get("settings"))
        order = settings.get("order", [])
        current_data = {
            "number_styles": _draft_mapping(ns.get("number_styles")),
            "schemes": _draft_mapping(ns.get("schemes")),
            "settings": {
                "default_scheme": config.text.default_scheme,
                "order": list(order) if isinstance(order, list) else [],
            },
        }

        dlg = NumberingAddDialog(
            self,
            config_data=current_data,
            on_save=self._on_numbering_schemes_saved,
        )
        dlg.exec()

    def _open_numbering_clean_editor(self) -> None:
        """Open the numbering removal rule editor (NumberingCleanDialog)."""
        from .numbering_clean_editor import NumberingCleanDialog

        config = self._vm.config
        cr = _draft_mapping(config.text.numbering_clean_rules)
        settings = _draft_mapping(cr.get("settings"))
        order = settings.get("order", [])
        rules = cr.get("rules", [])
        current_data = {
            "settings": {"order": list(order) if isinstance(order, list) else []},
            "rules": rules if isinstance(rules, (list, dict)) else [],
        }
        dlg = NumberingCleanDialog(
            self,
            config_data=current_data,
            on_save=self._on_numbering_clean_rules_saved,
        )
        dlg.exec()

    def _on_numbering_schemes_saved(self, schemes_data: dict) -> bool:
        ok = self._vm.persist_numbering_schemes_source(schemes_data)
        if not ok:
            from PySide6.QtWidgets import QMessageBox

            QMessageBox.warning(
                self,
                t("common.save_failed", "Save Failed"),
                t(
                    "settings.text.save_numbering_schemes_failed",
                    "Failed to save numbering schemes to disk. Changes were not persisted.",
                ),
            )
            return False
        default_scheme = schemes_data.get("settings", {}).get("default_scheme", "")
        updates: dict[str, object] = {"numbering_schemes": schemes_data}
        if isinstance(default_scheme, str) and default_scheme:
            updates["default_scheme"] = default_scheme
        self._vm.set_field_batch(SECTION_TEXT, updates)
        self._refresh_scheme_combo_items(str(default_scheme or ""))
        return True

    def _on_numbering_clean_rules_saved(self, rules_data: dict) -> bool:
        ok = self._vm.persist_numbering_clean_rules_source(rules_data)
        if not ok:
            from PySide6.QtWidgets import QMessageBox

            QMessageBox.warning(
                self,
                t("common.save_failed", "Save Failed"),
                t(
                    "settings.text.save_numbering_clean_rules_failed",
                    "Failed to save numbering clean rules to disk. Changes were not persisted.",
                ),
            )
            return False
        self._vm.set_field(SECTION_TEXT, "numbering_clean_rules", rules_data)
        return True
