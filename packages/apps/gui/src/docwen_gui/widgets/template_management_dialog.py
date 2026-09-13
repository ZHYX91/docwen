"""Template management UI for built-in and user-owned templates."""

from __future__ import annotations

import logging
from pathlib import Path

from PySide6.QtCore import QSignalBlocker, Qt, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from docwen_runtime.templates import TemplateManagementError, TemplateManager

from ..i18n import t
from ..styles.design_tokens import Sizing, Spacing
from ..styles.theme_semantics import apply_theme_class

logger = logging.getLogger(__name__)

_TEMPLATE_ID_ROLE = Qt.ItemDataRole.UserRole
_TEMPLATE_CUSTOM_ROLE = Qt.ItemDataRole.UserRole + 1


class TemplateManagementDialog(QDialog):
    """Immediate-apply management surface for DOCX/XLSX templates."""

    catalog_changed = Signal()

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        manager: TemplateManager | None = None,
        initial_target: str = "docx",
    ) -> None:
        super().__init__(parent)
        self._manager = manager or TemplateManager.default()
        self._lists: dict[str, QListWidget] = {}
        self._current_ids: dict[str, str | None] = {"docx": None, "xlsx": None}
        self.setWindowTitle(t("settings.templates.title", "Templates"))
        self.setModal(True)
        self.resize(760, 620)
        self.setMinimumSize(620, 480)
        self._build_ui()
        self.refresh()
        self._tabs.setCurrentIndex(1 if initial_target == "xlsx" else 0)
        self._sync_actions()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(Spacing.MD, Spacing.MD, Spacing.MD, Spacing.MD)
        root.setSpacing(Spacing.SM)

        intro = QLabel(
            t(
                "settings.templates.description",
                "Built-in templates are read-only. Copy one to create an editable custom template. "
                "Changes on this page take effect immediately.",
            ),
            self,
        )
        intro.setWordWrap(True)
        intro.setObjectName("templateManagementDescription")
        root.addWidget(intro)

        self._tabs = QTabWidget(self)
        self._tabs.setObjectName("templateManagementTabs")
        for target, title in (
            ("docx", t("components.template_selector_tabbed.document_templates", "DOCX templates")),
            ("xlsx", t("components.template_selector_tabbed.spreadsheet_templates", "XLSX templates")),
        ):
            page = QWidget(self._tabs)
            layout = QVBoxLayout(page)
            layout.setContentsMargins(0, Spacing.SM, 0, 0)
            template_list = QListWidget(page)
            template_list.setObjectName(f"templateManagementList-{target}")
            template_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
            template_list.itemSelectionChanged.connect(self._on_selection_changed)
            template_list.itemChanged.connect(self._on_item_changed)
            layout.addWidget(template_list, 1)
            self._lists[target] = template_list
            self._tabs.addTab(page, title)
        self._tabs.currentChanged.connect(lambda _index: self._sync_actions())
        root.addWidget(self._tabs, 1)

        primary_row = QHBoxLayout()
        primary_row.setSpacing(Spacing.XS)
        self._import_button = self._button(
            t("settings.templates.import", "Import templates"),
            self._import_templates,
            primary=True,
        )
        self._copy_button = self._button(
            t("settings.templates.copy_edit", "Copy and edit"),
            self._copy_and_edit,
        )
        self._edit_button = self._button(t("settings.templates.edit", "Edit"), self._edit_selected)
        self._rename_button = self._button(t("settings.templates.rename", "Rename"), self._rename_selected)
        self._export_button = self._button(t("settings.templates.export", "Export"), self._export_selected)
        self._delete_button = self._button(t("settings.templates.delete", "Delete"), self._delete_selected)
        for button in (
            self._import_button,
            self._copy_button,
            self._edit_button,
            self._rename_button,
            self._export_button,
            self._delete_button,
        ):
            primary_row.addWidget(button)
        primary_row.addStretch(1)
        root.addLayout(primary_row)

        secondary_row = QHBoxLayout()
        secondary_row.setSpacing(Spacing.XS)
        self._up_button = self._button(t("settings.templates.move_up", "Move up"), lambda: self._move_selected(-1))
        self._down_button = self._button(
            t("settings.templates.move_down", "Move down"),
            lambda: self._move_selected(1),
        )
        self._folder_button = self._button(
            t("settings.templates.open_user_folder", "Open custom template folder"),
            self._open_user_folder,
        )
        self._refresh_button = self._button(t("settings.templates.refresh", "Refresh"), self.refresh)
        secondary_row.addWidget(self._up_button)
        secondary_row.addWidget(self._down_button)
        secondary_row.addStretch(1)
        secondary_row.addWidget(self._folder_button)
        secondary_row.addWidget(self._refresh_button)
        root.addLayout(secondary_row)

        note = QLabel(
            t(
                "settings.templates.status_note",
                "Disabled templates stay on this page but are hidden from the main template selector.",
            ),
            self,
        )
        note.setWordWrap(True)
        note.setObjectName("templateManagementNote")
        root.addWidget(note)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, parent=self)
        close_button = buttons.button(QDialogButtonBox.StandardButton.Close)
        if close_button is not None:
            close_button.setText(t("common.close", "Close"))
            close_button.setMinimumHeight(Sizing.CONTROL_HEIGHT)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def _button(self, text: str, callback, *, primary: bool = False) -> QPushButton:
        button = QPushButton(text, self)
        button.setMinimumHeight(Sizing.CONTROL_HEIGHT)
        button.clicked.connect(callback)
        apply_theme_class(button, "primary" if primary else "secondary")
        return button

    def _target(self) -> str:
        return "xlsx" if self._tabs.currentIndex() == 1 else "docx"

    def _selected_item(self) -> QListWidgetItem | None:
        items = self._lists[self._target()].selectedItems()
        return items[0] if items else None

    def _selected_id(self) -> str | None:
        item = self._selected_item()
        if item is None:
            return None
        value = item.data(_TEMPLATE_ID_ROLE)
        return str(value) if value else None

    def refresh(self, *, select_id: str | None = None) -> None:
        for target, widget in self._lists.items():
            remembered = select_id if select_id is not None else self._current_ids.get(target)
            templates = self._manager.list_templates(target, include_disabled=True)
            with QSignalBlocker(widget):
                widget.clear()
                for template in templates:
                    custom = self._manager.is_custom(template)
                    source = (
                        t("settings.templates.custom", "Custom")
                        if custom
                        else t("settings.templates.builtin", "Built-in")
                    )
                    item = QListWidgetItem(f"{template.name}    [{source}]")
                    item.setData(_TEMPLATE_ID_ROLE, template.id)
                    item.setData(_TEMPLATE_CUSTOM_ROLE, custom)
                    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                    item.setCheckState(
                        Qt.CheckState.Checked if self._manager.is_enabled(template.id) else Qt.CheckState.Unchecked
                    )
                    item.setToolTip(
                        str(template.path)
                        if custom
                        else t(
                            "settings.templates.builtin_hint",
                            "Built-in template: copy it before editing.",
                        )
                    )
                    widget.addItem(item)
                    if template.id == remembered:
                        item.setSelected(True)
                        widget.setCurrentItem(item)
            if widget.currentItem() is None and widget.count() > 0:
                widget.setCurrentRow(0)
            current = widget.currentItem()
            self._current_ids[target] = str(current.data(_TEMPLATE_ID_ROLE)) if current is not None else None
        self._sync_actions()

    def _on_selection_changed(self) -> None:
        item = self._selected_item()
        if item is not None:
            self._current_ids[self._target()] = str(item.data(_TEMPLATE_ID_ROLE))
        self._sync_actions()

    def _on_item_changed(self, item: QListWidgetItem) -> None:
        template_id = item.data(_TEMPLATE_ID_ROLE)
        if not template_id:
            return
        try:
            self._manager.set_enabled(str(template_id), item.checkState() == Qt.CheckState.Checked)
        except Exception as exc:
            self._show_error(exc)
            self.refresh(select_id=str(template_id))
            return
        self.catalog_changed.emit()

    def _sync_actions(self) -> None:
        item = self._selected_item()
        has_selection = item is not None
        custom = bool(item.data(_TEMPLATE_CUSTOM_ROLE)) if item is not None else False
        self._copy_button.setEnabled(has_selection and not custom)
        for button in (self._edit_button, self._rename_button, self._export_button, self._delete_button):
            button.setEnabled(has_selection and custom)
        self._up_button.setEnabled(has_selection)
        self._down_button.setEnabled(has_selection)

    def _import_templates(self) -> None:
        paths, _selected_filter = QFileDialog.getOpenFileNames(
            self,
            t("settings.templates.import", "Import templates"),
            "",
            t("settings.templates.file_filter", "Office templates (*.docx *.xlsx);;All files (*)"),
        )
        if not paths:
            return
        last_id: str | None = None
        errors: list[str] = []
        for path in paths:
            try:
                imported = self._manager.import_template(path)
                last_id = imported.id
            except Exception as exc:
                errors.append(f"{Path(path).name}: {exc}")
        self.refresh(select_id=last_id)
        if last_id is not None:
            self.catalog_changed.emit()
        if errors:
            QMessageBox.warning(self, t("common.error", "Error"), "\n".join(errors))

    def _copy_and_edit(self) -> None:
        template_id = self._selected_id()
        if template_id is None:
            return
        template = next(
            (item for item in self._manager.list_templates(include_disabled=True) if item.id == template_id),
            None,
        )
        if template is None:
            return
        name, accepted = QInputDialog.getText(
            self,
            t("settings.templates.copy_edit", "Copy and edit"),
            t("settings.templates.name_prompt", "Template name:"),
            text=f"{template.name} - custom",
        )
        if not accepted:
            return
        try:
            copied = self._manager.copy_builtin_as_custom(template_id, custom_name=name)
        except Exception as exc:
            self._show_error(exc)
            return
        self.refresh(select_id=copied.id)
        self.catalog_changed.emit()
        self._open_file(copied.path)

    def _edit_selected(self) -> None:
        template = self._selected_template()
        if template is not None and self._manager.is_custom(template):
            self._open_file(template.path)

    def _rename_selected(self) -> None:
        template = self._selected_template()
        if template is None or not self._manager.is_custom(template):
            return
        name, accepted = QInputDialog.getText(
            self,
            t("settings.templates.rename", "Rename"),
            t("settings.templates.name_prompt", "Template name:"),
            text=template.name,
        )
        if not accepted:
            return
        try:
            renamed = self._manager.rename_custom(template.id, name)
        except Exception as exc:
            self._show_error(exc)
            return
        self.refresh(select_id=renamed.id)
        self.catalog_changed.emit()

    def _export_selected(self) -> None:
        template = self._selected_template()
        if template is None or not self._manager.is_custom(template):
            return
        path, _selected_filter = QFileDialog.getSaveFileName(
            self,
            t("settings.templates.export", "Export"),
            template.path.name,
            f"*.{template.target}",
        )
        if not path:
            return
        try:
            self._manager.export_custom(template.id, path)
        except Exception as exc:
            self._show_error(exc)

    def _delete_selected(self) -> None:
        template = self._selected_template()
        if template is None or not self._manager.is_custom(template):
            return
        result = QMessageBox.question(
            self,
            t("settings.templates.delete", "Delete"),
            t(
                "settings.templates.delete_confirm",
                "Delete custom template “{name}”?",
                name=template.name,
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if result != QMessageBox.StandardButton.Yes:
            return
        try:
            self._manager.delete_custom(template.id)
        except Exception as exc:
            self._show_error(exc)
            return
        self.refresh()
        self.catalog_changed.emit()

    def _move_selected(self, offset: int) -> None:
        template_id = self._selected_id()
        if template_id is None:
            return
        try:
            self._manager.move(template_id, offset)
        except Exception as exc:
            self._show_error(exc)
            return
        self.refresh(select_id=template_id)
        self.catalog_changed.emit()

    def _open_user_folder(self) -> None:
        try:
            directory = self._manager.ensure_user_directory()
        except Exception as exc:
            self._show_error(exc)
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(directory)))

    def _selected_template(self):
        template_id = self._selected_id()
        if template_id is None:
            return None
        return next(
            (item for item in self._manager.list_templates(include_disabled=True) if item.id == template_id),
            None,
        )

    @staticmethod
    def _open_file(path: Path) -> None:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def _show_error(self, error: Exception) -> None:
        logger.exception("Template management operation failed", exc_info=error)
        message = str(error)
        if isinstance(error, TemplateManagementError):
            message = str(error)
        QMessageBox.warning(self, t("common.error", "Error"), message)


__all__ = ["TemplateManagementDialog"]
