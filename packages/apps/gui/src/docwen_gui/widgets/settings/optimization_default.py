"""One category's default selection, shared by the built-in optimization pages."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QPushButton, QVBoxLayout, QWidget

from docwen_gui.i18n import t
from docwen_gui.styles.design_tokens import Spacing
from docwen_gui.styles.ui_scale import set_metric
from docwen_gui.view_models.settings_vm import SECTION_CONVERSION_DEFAULTS, SettingsViewModel
from docwen_gui.widgets.panel_card import WrappingLabel


class DefaultOptimizationControl(QWidget):
    """Show the shared choice and explicitly select or clear this page's resource."""

    def __init__(
        self,
        view_model: SettingsViewModel,
        *,
        category: str,
        resource_id: str,
        resource_label: str,
        parent: QWidget,
    ) -> None:
        super().__init__(parent)
        self._vm = view_model
        self._category = category
        self._resource_id = resource_id
        self._resource_label = resource_label
        self.setObjectName(f"{category}OptimizationDefault")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        set_metric(layout, "setSpacing", Spacing.SM)
        self._current_label = WrappingLabel("", self)
        self._current_label.setObjectName(f"{category}OptimizationCurrent")
        self._current_label.setTextFormat(Qt.TextFormat.PlainText)
        layout.addWidget(self._current_label)
        self._action_button = QPushButton(self)
        self._action_button.setObjectName(f"{category}OptimizationDefaultButton")
        self._action_button.setAutoDefault(False)
        self._action_button.clicked.connect(self._change_default)
        layout.addWidget(self._action_button)
        self._availability_label = WrappingLabel("", self)
        self._availability_label.setObjectName(f"{category}OptimizationAvailability")
        self._availability_label.setTextFormat(Qt.TextFormat.PlainText)
        self._availability_label.setProperty("settingsRole", "sectionDescription")
        layout.addWidget(self._availability_label)
        self._vm.config_changed.connect(self._on_config_changed)
        self._vm.config_reloaded.connect(self.reload_from_config)
        self.reload_from_config()

    def _on_config_changed(self, section: str, key: str, _value: object) -> None:
        if section == SECTION_CONVERSION_DEFAULTS and key == f"{self._category}.to_md_optimization":
            self.reload_from_config()

    def reload_from_config(self) -> None:
        current = self._vm.get_default_optimization(self._category)
        result = self._vm.get_optimization_choices_result(source_category=self._category)
        choice = result.get(current)
        known_labels = {
            "none": t("settings.optimization.no_default"),
            "gongwen": t("settings.gongwen.title"),
            "invoice_cn": t("settings.invoice.title"),
        }
        current_label = choice.label if choice is not None else known_labels.get(current, current)
        self._current_label.setText(t("settings.optimization.current_default", name=current_label))
        selected = current == self._resource_id
        available = result.get(self._resource_id) is not None
        action = t("settings.optimization.clear_default") if selected else t("settings.optimization.set_default")
        self._action_button.setText(action)
        self._action_button.setAccessibleName(f"{self._resource_label} · {action}")
        self._action_button.setAccessibleDescription(self._current_label.text())
        # A saved unavailable choice must remain removable.
        self._action_button.setEnabled(selected or available)
        reason = ""
        if result.status == "failed":
            reason = t("settings.optimization.discovery_failed")
        elif not available:
            reason = t("settings.optimization.unavailable", name=self._resource_label)
        self._availability_label.setText(reason)
        self._availability_label.setVisible(bool(reason))
        self._availability_label.setToolTip(str(result.error or ""))
        self._action_button.setToolTip(reason if not selected else "")

    def _change_default(self) -> None:
        if self._vm.get_default_optimization(self._category) == self._resource_id:
            self._vm.clear_default_optimization(self._category, self._resource_id)
        else:
            self._vm.set_default_optimization(self._category, self._resource_id)
        self.reload_from_config()
