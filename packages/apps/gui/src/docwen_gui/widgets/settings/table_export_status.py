"""Read-only projection of the shared Structural Tables output policy."""

from __future__ import annotations

from PySide6.QtCore import QObject, Slot
from PySide6.QtWidgets import QLabel, QWidget

from ...i18n import t
from ...view_models.settings_vm import SettingsViewModel


class TableExportStatus(QObject):
    """Keep status current across tabs, presets and configuration reloads."""

    def __init__(self, label: QLabel, view_model: SettingsViewModel) -> None:
        super().__init__(label)
        self._label = label
        self._vm = view_model
        self._vm.config_changed.connect(self._on_changed)
        self._vm.config_reloaded.connect(self.refresh)
        self.refresh()

    @classmethod
    def attach(cls, tab: QWidget, view_model: SettingsViewModel) -> TableExportStatus:
        label = tab.findChild(QLabel, "tableExportStatus")
        if label is None:
            raise RuntimeError("Missing table export status label")
        return cls(label, view_model)

    @Slot(str, str, object)
    def _on_changed(self, section: str, key: str, _value: object) -> None:
        if section == "formatting" and key == "markdown_extensions":
            self.refresh()

    @Slot()
    def refresh(self) -> None:
        enabled = self._vm.config.formatting.markdown_extensions["output"]["structural_tables"]
        if enabled:
            text = t(
                "settings.table_export.structural_active",
                "Standard document/spreadsheet export: Structural Tables is on; Fill/Empty is retained but not used.",
            )
        else:
            text = t(
                "settings.table_export.ordinary_active",
                "Standard document/spreadsheet export: Structural Tables is off; each input tab's Fill/Empty choice applies.",
            )
        self._label.setText(text)
