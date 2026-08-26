"""Heading numbering removal/clean rule editor.

Port of the old ``NumberingCleanDialog`` — regex-based heading number
removal rule CRUD, regex test playground with live match feedback,
direct regex patterns (no placeholders), and TOML persistence.

Placed in ``widgets/settings/``; opened from ``TextTab`` editor buttons.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ...i18n import t

# ── Data ──────────────────────────────────────────────────────────────────


@dataclass
class CleanRule:
    """A single numbering removal rule.

    Fields map directly to the new TOML schema:
    - ``id``: unique rule identifier
    - ``name`` / ``name_key``: optional human-readable list label
    - ``enabled``: whether the rule is active
    - ``pattern``: direct regex pattern (no placeholders)
    - ``description``: human-readable description
    - ``description_key``: optional localized description key
    - ``level``: implied heading level 1-5
    """

    id: str
    name: str = ""
    name_key: str = ""
    enabled: bool = True
    pattern: str = ""
    description: str = ""
    description_key: str = ""
    level: int = 1
    is_system: bool = False

    @classmethod
    def from_dict(cls, rule_id: str, data: dict[str, Any]) -> CleanRule:
        """Create a rule from the canonical flat ``[[rules]]`` schema."""
        return cls(
            id=rule_id,
            name=str(data.get("name", "")),
            name_key=str(data.get("name_key", "")),
            enabled=bool(data.get("enabled", True)),
            pattern=str(data.get("pattern", "")),
            description=str(data.get("description", "")),
            description_key=str(data.get("description_key", "")),
            level=int(data.get("level", 1)),
            is_system=bool(data.get("is_system", False)),
        )

    def display_name(self) -> str:
        """Return the localized list label, with stable-id fallback."""
        if self.name_key:
            key = f"editors.numbering_clean.names.{self.name_key}"
            translated = t(key, "")
            if translated and translated not in {key, f"[{key}]"}:
                return translated
        return self.name or self.id

    def display_description(self) -> str:
        """Return the localized description, with configured-text fallback."""
        if self.description_key:
            key = f"editors.numbering_clean.descriptions.{self.description_key}"
            translated = t(key, "")
            if translated and translated not in {key, f"[{key}]"}:
                return translated
        return self.description

    def to_dict(self) -> dict[str, Any]:
        """Serialise to the new TOML rule dict (no id — id is the key)."""
        return {
            "name": self.name,
            "name_key": self.name_key,
            "enabled": bool(self.enabled),
            "pattern": self.pattern,
            "description": self.description,
            "description_key": self.description_key,
            "level": int(self.level),
        }

    def copy(self, new_id: str, new_name: str) -> CleanRule:
        """Create a custom copy with localized display text frozen as user data."""
        return CleanRule(
            id=new_id,
            name=new_name,
            enabled=self.enabled,
            pattern=self.pattern,
            description=self.display_description(),
            level=self.level,
            is_system=False,
        )


# ── Dialog ────────────────────────────────────────────────────────────────


class NumberingCleanDialog(QDialog):
    """Dialog for managing heading numbering removal rules.

    Features:
    - Rule list with CRUD (create/copy/delete) and reorder (move up/down).
    - Pattern editor with enabled/disabled toggle and level selector.
    - Live regex test playground.
    - Dirty-state tracking; save hands data to ``on_save`` callback (the
      Settings view model persists and reloads through its injected port).
    """

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        config_data: dict[str, Any] | None = None,
        on_save: Any | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("numberingCleanDialog")
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self._on_save = on_save

        self.rules: dict[str, CleanRule] = {}
        self.order: list[str] = []
        self.current_rule_id: str | None = None
        self._loading_form = False
        self._dirty = False
        self._saved_state: tuple[tuple[str, ...], tuple[tuple[str, tuple[Any, ...]], ...]] = ((), ())
        self._base_window_title = t("editors.numbering_clean.window_title", "Numbering Removal Rule Editor")

        self.setWindowTitle(self._base_window_title)
        self.resize(1000, 720)

        self._build_ui()

        if config_data is not None:
            self._load_from_dict(config_data)

    # ── UI construction ──────────────────────────────────────────────────

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        content = QHBoxLayout()
        root.addLayout(content, 1)

        # -- left panel: rule list + buttons --
        left = QVBoxLayout()
        content.addLayout(left, 1)
        left.addWidget(QLabel(t("editors.numbering_clean.rule_list", "Rules:"), self))
        self.rule_list = QListWidget(self)
        self.rule_list.setTextElideMode(Qt.TextElideMode.ElideMiddle)
        self.rule_list.currentRowChanged.connect(self._on_rule_selected)
        left.addWidget(self.rule_list, 1)

        left_btns = QHBoxLayout()
        self.new_btn = QPushButton(t("editors.numbering_clean.new_rule", "New Rule"), self)
        self.new_btn.clicked.connect(self._create_new_rule)
        left_btns.addWidget(self.new_btn)
        self.copy_btn = QPushButton(t("editors.common.copy", "Copy"), self)
        self.copy_btn.clicked.connect(self._copy_selected_rule)
        left_btns.addWidget(self.copy_btn)
        self.delete_btn = QPushButton(t("editors.common.delete", "Delete"), self)
        self.delete_btn.clicked.connect(self._delete_selected_rule)
        left_btns.addWidget(self.delete_btn)
        left.addLayout(left_btns)

        order_btns = QHBoxLayout()
        self.up_btn = QPushButton(t("editors.common.move_up", "Move Up"), self)
        self.up_btn.clicked.connect(self._move_selected_rule_up)
        order_btns.addWidget(self.up_btn)
        self.down_btn = QPushButton(t("editors.common.move_down", "Move Down"), self)
        self.down_btn.clicked.connect(self._move_selected_rule_down)
        order_btns.addWidget(self.down_btn)
        left.addLayout(order_btns)

        # -- right panel: form + regex test --
        right = QVBoxLayout()
        content.addLayout(right, 2)
        form = QFormLayout()

        self.enabled_check = QCheckBox(t("editors.numbering_clean.enabled", "Enabled"), self)
        self.enabled_check.toggled.connect(self._on_enabled_changed)
        form.addRow("", self.enabled_check)

        self.pattern_edit = QLineEdit(self)
        self.pattern_edit.textChanged.connect(self._on_pattern_changed)
        form.addRow(t("editors.numbering_clean.regex", "Pattern:"), self.pattern_edit)

        self.desc_edit = QLineEdit(self)
        self.desc_edit.textChanged.connect(self._on_description_changed)
        form.addRow(t("editors.numbering_clean.description", "Description:"), self.desc_edit)

        self.level_spin = QSpinBox(self)
        self.level_spin.setRange(1, 5)
        self.level_spin.valueChanged.connect(self._on_level_changed)
        form.addRow(t("editors.numbering_add.level", "Level:"), self.level_spin)
        right.addLayout(form)

        # Regex test area
        right.addWidget(QLabel(t("editors.numbering_clean.regex_test", "Regex Test:"), self))
        self.test_input = QLineEdit(self)
        self.test_input.setPlaceholderText(t("editors.numbering_clean.input", "Enter test text here..."))
        self.test_input.textChanged.connect(self._run_regex_test)
        right.addWidget(self.test_input)

        self.test_result = QTextEdit(self)
        self.test_result.setReadOnly(True)
        right.addWidget(QLabel(t("editors.numbering_clean.result", "Result:"), self))
        right.addWidget(self.test_result, 1)

        self.status_label = QLabel("", self)
        right.addWidget(self.status_label)

        # -- bottom bar --
        bottom = QHBoxLayout()
        bottom.addStretch(1)
        self.save_btn = QPushButton(t("editors.common.save", "Save"), self)
        self.save_btn.clicked.connect(self._save)
        bottom.addWidget(self.save_btn)
        self.cancel_btn = QPushButton(t("editors.common.cancel", "Cancel"), self)
        self.cancel_btn.clicked.connect(self.reject)
        bottom.addWidget(self.cancel_btn)
        root.addLayout(bottom)

    # ── Load ─────────────────────────────────────────────────────────────

    def _load_from_dict(self, data: dict[str, Any]) -> None:
        settings = data.get("settings", {})
        if not isinstance(settings, dict):
            settings = {}
        self.order = [str(item) for item in settings.get("order", []) if isinstance(item, str)]

        rules_raw = data.get("rules", [])
        self.rules = {}

        if isinstance(rules_raw, list):
            for entry in rules_raw:
                if isinstance(entry, dict):
                    rid = str(entry.get("id", ""))
                    if rid:
                        self.rules[rid] = CleanRule.from_dict(rid, entry)

        self.order = [rid for rid in self.order if rid in self.rules]
        for rid in self.rules:
            if rid not in self.order:
                self.order.append(rid)

        self._refresh_rule_list()
        if self.order:
            self._select_rule(self.order[0])
        self._saved_state = self._capture_state()
        self._refresh_dirty_state()

    # ── Rule list views ─────────────────────────────────────────────────

    @staticmethod
    def _update_rule_list_item(item: QListWidgetItem, rule: CleanRule) -> None:
        suffix = "" if rule.enabled else " [off]"
        item.setText(f"{rule.display_name()}{suffix}")
        description = rule.display_description()
        item.setToolTip(f"{description}\n{rule.pattern}" if description else rule.pattern)

    def _refresh_rule_list(self) -> None:
        selected_id = self.current_rule_id
        self.rule_list.blockSignals(True)
        self.rule_list.clear()
        for rid in self.order:
            rule = self.rules[rid]
            item = QListWidgetItem()
            self._update_rule_list_item(item, rule)
            self.rule_list.addItem(item)
        self.rule_list.blockSignals(False)
        if selected_id and selected_id in self.order:
            self._select_rule(selected_id)
        else:
            self._refresh_action_states()

    def _select_rule(self, rule_id: str) -> None:
        if rule_id not in self.rules:
            return
        self.current_rule_id = rule_id
        row = self.order.index(rule_id)
        self.rule_list.blockSignals(True)
        self.rule_list.setCurrentRow(row)
        self.rule_list.blockSignals(False)
        self._populate_form(self.rules[rule_id])

    def _on_rule_selected(self, row: int) -> None:
        if row < 0 or row >= len(self.order):
            return
        self.current_rule_id = self.order[row]
        self._populate_form(self.rules[self.current_rule_id])

    def _populate_form(self, rule: CleanRule) -> None:
        self._loading_form = True
        try:
            self.enabled_check.setChecked(rule.enabled)
            self.pattern_edit.setText(rule.pattern)
            self.desc_edit.setText(rule.display_description())
            self.level_spin.setValue(rule.level)
            self._run_regex_test()
            self._refresh_action_states()
        finally:
            self._loading_form = False

    # ── Dirty tracking ──────────────────────────────────────────────────

    def _mark_dirty(self) -> None:
        self._refresh_dirty_state()

    def _capture_state(self) -> tuple[tuple[str, ...], tuple[tuple[str, tuple[Any, ...]], ...]]:
        items: list[tuple[str, tuple[Any, ...]]] = []
        for rid in sorted(self.rules):
            rule = self.rules[rid]
            items.append(
                (
                    rid,
                    (
                        bool(rule.enabled),
                        rule.name,
                        rule.name_key,
                        rule.pattern,
                        rule.description,
                        rule.description_key,
                        int(rule.level),
                        bool(rule.is_system),
                    ),
                )
            )
        return (tuple(self.order), tuple(items))

    def _refresh_dirty_state(self) -> None:
        self._dirty = self._capture_state() != self._saved_state
        self.setWindowTitle(f"{self._base_window_title} *" if self._dirty else self._base_window_title)
        self.save_btn.setEnabled(self._dirty and self._can_save())

    def _can_save(self) -> bool:
        err_prefix = t("editors.numbering_clean.regex_error", "Regex Error")
        return not self.status_label.text().startswith(f"{err_prefix}:")

    def _refresh_action_states(self) -> None:
        rule = self._current_rule()
        has = rule is not None and rule.id in self.order
        if not has or rule is None:
            self.delete_btn.setEnabled(False)
            self.up_btn.setEnabled(False)
            self.down_btn.setEnabled(False)
            return
        idx = self.order.index(rule.id)
        self.delete_btn.setEnabled(not rule.is_system)
        self.up_btn.setEnabled(idx > 0)
        self.down_btn.setEnabled(idx < len(self.order) - 1)

    def _current_rule(self) -> CleanRule | None:
        if not self.current_rule_id:
            return None
        return self.rules.get(self.current_rule_id)

    # ── Field change handlers ───────────────────────────────────────────

    def _on_enabled_changed(self, checked: bool) -> None:
        if self._loading_form:
            return
        rule = self._current_rule()
        if rule is None:
            return
        rule.enabled = checked
        self._refresh_rule_list()
        self._mark_dirty()

    def _on_pattern_changed(self, text: str) -> None:
        if self._loading_form:
            return
        rule = self._current_rule()
        if rule is None:
            return
        rule.pattern = text
        self._run_regex_test()
        self._mark_dirty()

    def _on_description_changed(self, text: str) -> None:
        if self._loading_form:
            return
        rule = self._current_rule()
        if rule is None:
            return
        rule.description = text
        rule.description_key = ""
        if rule.id in self.order:
            item = self.rule_list.item(self.order.index(rule.id))
            if item is not None:
                self._update_rule_list_item(item, rule)
        self._mark_dirty()

    def _on_level_changed(self, value: int) -> None:
        if self._loading_form:
            return
        rule = self._current_rule()
        if rule is None:
            return
        rule.level = value
        self._mark_dirty()

    # ── Rule CRUD ───────────────────────────────────────────────────────

    def _generate_unique_id(self, base: str) -> str:
        candidate = re.sub(r"[^a-zA-Z0-9_]+", "_", base).strip("_").lower() or "custom_rule"
        if candidate not in self.rules:
            return candidate
        idx = 2
        while f"{candidate}_{idx}" in self.rules:
            idx += 1
        return f"{candidate}_{idx}"

    def _create_new_rule(self) -> None:
        base = "new_rule"
        rid = self._generate_unique_id(base)
        rule = CleanRule(
            id=rid,
            name=t("editors.numbering_clean.new_rule", "New Rule"),
            pattern=r"^\s*",
        )
        self.rules[rid] = rule
        self.order.append(rid)
        self._refresh_rule_list()
        self._select_rule(rid)
        self._mark_dirty()

    def _copy_selected_rule(self) -> None:
        rule = self._current_rule()
        if rule is None:
            self._notify_info(t("editors.common.select_item_first", "Please select an item first"))
            return
        rid = self._generate_unique_id(f"{rule.id}_copy")
        new_name = f"{rule.display_name()} {t('editors.numbering_clean.copy_suffix', 'Copy')}"
        copied = rule.copy(rid, new_name)
        self.rules[rid] = copied
        insert_at = self.order.index(rule.id) + 1
        self.order.insert(insert_at, rid)
        self._refresh_rule_list()
        self._select_rule(rid)
        self._mark_dirty()

    def _delete_selected_rule(self) -> None:
        rule = self._current_rule()
        if rule is None:
            self._notify_info(t("editors.common.select_item_first", "Please select an item first"))
            return
        if rule.is_system:
            self._notify_warning(
                t("editors.common.system_item_hint", "System rules can be disabled but cannot be deleted.")
            )
            return
        if not self._confirm(
            t("editors.common.confirm_delete", "Confirm Delete"),
            t("editors.common.confirm_delete_message", "Delete rule '{name}'?", name=rule.id),
        ):
            return
        self.rules.pop(rule.id, None)
        self.order = [item for item in self.order if item != rule.id]
        self.current_rule_id = None
        self._refresh_rule_list()
        if self.order:
            self._select_rule(self.order[0])
        self._mark_dirty()

    def _move_selected_rule_up(self) -> None:
        rule = self._current_rule()
        if rule is None:
            return
        idx = self.order.index(rule.id)
        if idx <= 0:
            return
        self.order[idx - 1], self.order[idx] = self.order[idx], self.order[idx - 1]
        self._refresh_rule_list()
        self._select_rule(rule.id)
        self._mark_dirty()

    def _move_selected_rule_down(self) -> None:
        rule = self._current_rule()
        if rule is None:
            return
        idx = self.order.index(rule.id)
        if idx >= len(self.order) - 1:
            return
        self.order[idx + 1], self.order[idx] = self.order[idx], self.order[idx + 1]
        self._refresh_rule_list()
        self._select_rule(rule.id)
        self._mark_dirty()

    # ── Regex test ──────────────────────────────────────────────────────

    def _run_regex_test(self) -> None:
        rule = self._current_rule()
        if rule is None:
            self.status_label.clear()
            self.test_result.clear()
            return
        text = self.test_input.text()
        pattern = rule.pattern
        try:
            match = re.match(pattern, text)
        except re.error as exc:
            self.status_label.setText(f"{t('editors.numbering_clean.regex_error', 'Regex Error')}: {exc}")
            self.test_result.setPlainText(pattern)
            self._refresh_dirty_state()
            return

        if match:
            self.status_label.setText(t("editors.numbering_clean.match_success", "Match Success"))
            cleaned = re.sub(pattern, "", text)
            self.test_result.setPlainText(cleaned)
        else:
            self.status_label.setText(t("editors.numbering_clean.no_match", "No Match"))
            self.test_result.setPlainText(text)
        self._refresh_dirty_state()

    # ── Save / Cancel ───────────────────────────────────────────────────

    def _save(self) -> bool:
        if not self._can_save():
            return False

        doc_data = self._build_toml_dict()
        if callable(self._on_save) and self._on_save(doc_data) is False:
            return False

        self._saved_state = self._capture_state()
        self._refresh_dirty_state()
        self.accept()
        return True

    def _build_toml_dict(self) -> dict[str, Any]:
        """Build a dict ready for TOML serialisation.

        Returns:
            ``{"settings": {"order": [...]}, "rules": [...]}`` — rules
            is an array of tables (new schema).
        """
        rules_list: list[dict[str, Any]] = []
        for rid in self.order:
            rule = self.rules[rid]
            entry: dict[str, Any] = {
                "id": rid,
                "name": rule.name,
                "name_key": rule.name_key,
                "enabled": bool(rule.enabled),
                "pattern": rule.pattern,
                "description": rule.description,
                "description_key": rule.description_key,
                "level": int(rule.level),
                "is_system": bool(rule.is_system),
            }
            rules_list.append(entry)

        return {
            "settings": {"order": list(self.order)},
            "rules": rules_list,
        }

    def reject(self) -> None:
        if self._dirty and not self._confirm(
            t("editors.common.unsaved_changes_hint", "Unsaved Changes"),
            t("editors.common.unsaved_close_message", "Discard unsaved changes?"),
        ):
            return
        super().reject()

    # ── Helpers ─────────────────────────────────────────────────────────

    def _confirm(self, title: str, message: str) -> bool:
        from PySide6.QtWidgets import QMessageBox

        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setText(message)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        box.setDefaultButton(QMessageBox.StandardButton.No)
        return box.exec() == QMessageBox.StandardButton.Yes

    def _notify_info(self, message: str, *, title: str = "") -> None:
        from PySide6.QtWidgets import QMessageBox

        box = QMessageBox(self)
        box.setWindowTitle(title or t("editors.numbering_clean.prompt", "Info"))
        box.setText(message)
        box.setIcon(QMessageBox.Icon.Information)
        box.setStandardButtons(QMessageBox.StandardButton.Ok)
        box.exec()

    def _notify_error(self, title: str, message: str) -> None:
        from PySide6.QtWidgets import QMessageBox

        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setText(message)
        box.setIcon(QMessageBox.Icon.Critical)
        box.setStandardButtons(QMessageBox.StandardButton.Ok)
        box.exec()

    def _notify_warning(self, message: str, *, title: str = "") -> None:
        from PySide6.QtWidgets import QMessageBox

        box = QMessageBox(self)
        box.setWindowTitle(title or t("common.type_warning", "Warning"))
        box.setText(message)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setStandardButtons(QMessageBox.StandardButton.Ok)
        box.exec()

    # ── Public API ──────────────────────────────────────────────────────

    def get_rules_data(self) -> dict[str, Any]:
        """Return the current rules as a dict ready for TOML serialization."""
        return self._build_toml_dict()
