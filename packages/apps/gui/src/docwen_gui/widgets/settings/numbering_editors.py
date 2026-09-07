"""Shared numbering editors opened from either input settings page."""

from __future__ import annotations

from collections.abc import Callable, Mapping

from PySide6.QtWidgets import QPushButton

from ...i18n import t
from ...view_models.settings_vm import SECTION_TEXT, SettingsViewModel
from .base_tab import BaseSettingsTab


def _draft_mapping(value: object) -> dict[str, object]:
    return {str(key): item for key, item in value.items()} if isinstance(value, Mapping) else {}


class NumberingEditors:
    def __init__(self, tab: BaseSettingsTab, view_model: SettingsViewModel, on_changed: Callable[[], None]) -> None:
        self._tab = tab
        self._vm = view_model
        self._on_changed = on_changed
        self._vm.config_reloaded.connect(on_changed)
        _card, form = tab.add_settings_card(
            t("settings.text.numbering_settings_section", "Numbering Settings"),
            t("settings.text.numbering_settings_desc", "Edit numbering addition schemes and removal rules."),
            object_name="numberingEditorsCard",
        )
        self.add_button = QPushButton(t("settings.text.edit_numbering_add", "Edit Numbering Addition Schemes"), tab)
        self.add_button.clicked.connect(self._open_numbering_scheme_editor)
        form.addRow(self.add_button)
        self.clean_button = QPushButton(t("settings.text.edit_numbering_clean", "Edit Numbering Removal Rules"), tab)
        self.clean_button.clicked.connect(self._open_numbering_clean_editor)
        form.addRow(self.clean_button)

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
            self._tab,
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
            self._tab,
            config_data=current_data,
            on_save=self._on_numbering_clean_rules_saved,
        )
        dlg.exec()

    def _on_numbering_schemes_saved(self, schemes_data: dict) -> bool:
        ok = self._vm.persist_numbering_schemes_source(schemes_data)
        if not ok:
            from PySide6.QtWidgets import QMessageBox

            QMessageBox.warning(
                self._tab,
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
        self._on_changed()
        return True

    def _on_numbering_clean_rules_saved(self, rules_data: dict) -> bool:
        ok = self._vm.persist_numbering_clean_rules_source(rules_data)
        if not ok:
            from PySide6.QtWidgets import QMessageBox

            QMessageBox.warning(
                self._tab,
                t("common.save_failed", "Save Failed"),
                t(
                    "settings.text.save_numbering_clean_rules_failed",
                    "Failed to save numbering clean rules to disk. Changes were not persisted.",
                ),
            )
            return False
        self._vm.set_field(SECTION_TEXT, "numbering_clean_rules", rules_data)
        self._on_changed()
        return True
