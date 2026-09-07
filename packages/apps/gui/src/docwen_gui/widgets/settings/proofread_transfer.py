"""Proofread rule import preview and complete export through the settings port."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHeaderView,
    QLabel,
    QMessageBox,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from docwen_runtime.config import atomic_write_text
from docwen_runtime.config.proofread_transfer import RuleImportPlan, plan_rule_import

from ...i18n import t
from ...styles.design_tokens import Spacing
from ...styles.theme_semantics import apply_theme_class
from ...view_models.settings_vm import SettingsViewModel
from ..value_controls import ScrollSafeComboBox

_CHANGE_LABEL_KEYS = {
    "added": "settings.rule_transfer.added",
    "unchanged": "settings.rule_transfer.unchanged",
    "conflict": "settings.rule_transfer.conflict",
    "existing_only": "settings.rule_transfer.existing_only",
}


class _ChangesModel(QAbstractTableModel):
    def __init__(self, plan: RuleImportPlan, parent=None) -> None:
        super().__init__(parent)
        self.plan = plan

    def rowCount(self, parent=None) -> int:
        return 0 if parent is not None and parent.isValid() else len(self.plan.changes)

    def columnCount(self, parent=None) -> int:
        return 0 if parent is not None and parent.isValid() else 4

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or role not in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole):
            return None
        change = self.plan.changes[index.row()]
        return (t(_CHANGE_LABEL_KEYS[change.kind]), change.key, change.current, change.incoming)[index.column()]

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return (
                t("settings.rule_transfer.change"),
                t("settings.rule_transfer.rule"),
                t("settings.rule_transfer.current"),
                t("settings.rule_transfer.incoming"),
            )[section]
        return None


class RuleImportDialog(QDialog):
    def __init__(self, plan: RuleImportPlan, parent: QWidget) -> None:
        super().__init__(parent)
        self.setObjectName("ruleImportDialog")
        self.setWindowTitle(t("settings.rule_transfer.preview"))
        self.resize(720, 480)
        self.setMinimumSize(360, 320)
        self.plan = plan
        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            Spacing.CARD_PADDING, Spacing.CARD_PADDING, Spacing.CARD_PADDING, Spacing.CARD_PADDING
        )
        layout.setSpacing(Spacing.GROUP_GAP)
        counts = {key: sum(change.kind == key for change in plan.changes) for key in _CHANGE_LABEL_KEYS}
        summary = QLabel(
            t(
                "settings.rule_transfer.summary",
                added=counts["added"],
                conflict=counts["conflict"],
                unchanged=counts["unchanged"],
                existing_only=counts["existing_only"],
            ),
            self,
        )
        summary.setWordWrap(True)
        layout.addWidget(summary)
        table = QTableView(self)
        table.setModel(_ChangesModel(plan, table))
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setAlternatingRowColors(True)
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        table.verticalHeader().hide()
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.setAccessibleName(t("settings.rule_transfer.preview"))
        layout.addWidget(table, 1)
        self.strategy = ScrollSafeComboBox(self)
        self.strategy.setAccessibleName(t("settings.rule_transfer.strategy"))
        if plan.merge_text is not None:
            self.strategy.addItem(t("settings.rule_transfer.merge"), "merge")
        self.strategy.addItem(t("settings.rule_transfer.replace"), "replace")
        layout.addWidget(self.strategy)
        hint = QLabel(t("settings.rule_transfer.strategy_hint"), self)
        hint.setWordWrap(True)
        layout.addWidget(hint)
        if plan.merge_error:
            warning = QLabel(t("settings.rule_transfer.merge_unavailable") + "\n" + plan.merge_error, self)
            warning.setTextFormat(Qt.TextFormat.PlainText)
            warning.setWordWrap(True)
            layout.addWidget(warning)
        actions = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        actions.button(QDialogButtonBox.StandardButton.Ok).setText(t("settings.rule_transfer.confirm"))
        actions.button(QDialogButtonBox.StandardButton.Cancel).setText(t("common.cancel"))
        for button in actions.buttons():
            button.setObjectName("secondaryActionButton")
            apply_theme_class(button, "secondary")
        actions.accepted.connect(self.accept)
        actions.rejected.connect(self.reject)
        layout.addWidget(actions)

    @property
    def selected_text(self) -> str:
        if self.strategy.currentData() == "merge" and self.plan.merge_text is not None:
            return self.plan.merge_text
        return self.plan.replace_text


class ProofreadRuleTransfer:
    def __init__(self, parent: QWidget, view_model: SettingsViewModel) -> None:
        self.parent = parent
        self.view_model = view_model

    def import_rules(self, config_name: str) -> None:
        filename, _ = QFileDialog.getOpenFileName(self.parent, t("settings.rule_transfer.import"), "", "TOML (*.toml)")
        if not filename:
            return
        try:
            source = self.view_model.read_config_file_text(config_name)
            if source is None:
                raise ValueError(t("settings.rule_transfer.read_failed"))
            plan = plan_rule_import(config_name, source, Path(filename).read_text(encoding="utf-8-sig"))
            dialog = RuleImportDialog(plan, self.parent)
            if dialog.exec() != QDialog.DialogCode.Accepted:
                return
            if self.view_model.read_config_file_text(config_name) != source:
                raise ValueError(t("settings.rule_transfer.source_changed"))
            if not self.view_model.save_config_file_text(config_name, dialog.selected_text, expected_text=source):
                changed = self.view_model.read_config_file_text(config_name) != source
                raise ValueError(
                    t("settings.rule_transfer.source_changed") if changed else t("settings.rule_transfer.save_failed")
                )
        except Exception as error:
            self._error(t("settings.rule_transfer.import_failed"), error)
            return
        QMessageBox.information(self.parent, t("settings.rule_transfer.import"), t("settings.rule_transfer.imported"))

    def export_rules(self, config_name: str) -> None:
        filename, _ = QFileDialog.getSaveFileName(
            self.parent, t("settings.rule_transfer.export"), Path(config_name).name, "TOML (*.toml)"
        )
        if not filename:
            return
        try:
            source = self.view_model.read_config_file_text(config_name)
            if source is None:
                raise ValueError(t("settings.rule_transfer.read_failed"))
            atomic_write_text(Path(filename), source)
        except Exception as error:
            self._error(t("settings.rule_transfer.export_failed"), error)
            return
        QMessageBox.information(self.parent, t("settings.rule_transfer.export"), t("settings.rule_transfer.exported"))

    def _error(self, title: str, error: Exception) -> None:
        dialog = QMessageBox(QMessageBox.Icon.Warning, title, str(error), QMessageBox.StandardButton.Ok, self.parent)
        dialog.setTextFormat(Qt.TextFormat.PlainText)
        dialog.exec()
