"""Template management UI for built-in and user-owned templates."""

from __future__ import annotations

import logging
from pathlib import Path

from PySide6.QtCore import QSignalBlocker, Qt, QTimer, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import Pivot, PushButton

from docwen_runtime.templates import TemplateManagementError

from ...dialogs import feedback
from ...i18n import t
from ...resources import set_action_icon
from ...styles.design_tokens import Sizing, Spacing
from ...styles.theme_semantics import apply_theme_class
from ...view_models.template_vm import TemplateViewModel
from ..action_button import ActionButton
from ..template_item_delegate import CUSTOM_ROLE, DEFAULT_ROLE, SOURCE_ROLE, TemplateItemDelegate
from .base_tab import BaseSettingsTab

logger = logging.getLogger(__name__)

_TEMPLATE_ID_ROLE = Qt.ItemDataRole.UserRole
_TEMPLATE_CUSTOM_ROLE = Qt.ItemDataRole.UserRole + 1


class TemplatesTab(BaseSettingsTab):
    """Immediate-apply management surface for DOCX/XLSX templates."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        view_model: TemplateViewModel,
        initial_target: str = "docx",
    ) -> None:
        self._vm = view_model
        self._manager = view_model.manager
        self._lists: dict[str, QListWidget] = {}
        self._current_ids: dict[str, str | None] = {"docx": None, "xlsx": None}
        super().__init__(parent)
        self._vm.changed.connect(self._render)
        self._vm.failed.connect(self._show_catalog_error)
        app = QApplication.instance()
        if isinstance(app, QApplication):
            app.applicationStateChanged.connect(self._on_application_state_changed)
        self.refresh()
        self._tabs.setCurrentIndex(1 if initial_target == "xlsx" else 0)
        self._sync_actions()

    def _create_interface(self) -> None:
        root = self._scroll_layout
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
        self._catalog_error = QWidget(self)
        self._catalog_error.setObjectName("templateCatalogError")
        error_layout = QHBoxLayout(self._catalog_error)
        error_layout.setContentsMargins(0, 0, 0, 0)
        error_label = QLabel(t("settings.templates.load_failed"), self._catalog_error)
        error_label.setWordWrap(True)
        error_layout.addWidget(error_label, 1)
        details_button = PushButton(t("activity.details"), self._catalog_error)
        details_button.clicked.connect(lambda: self._show_error(RuntimeError(self._vm.error or "")))
        error_layout.addWidget(details_button)
        self._catalog_error.hide()
        root.addWidget(self._catalog_error)

        self._pivot = Pivot(self)
        self._pivot.setObjectName("templateManagementPivot")
        root.addWidget(self._pivot)
        self._tabs = QStackedWidget(self)
        self._tabs.setObjectName("templateManagementTabs")
        self._tabs.setMinimumHeight(240)
        for target, title in (
            ("docx", t("components.template_selector_tabbed.document_templates", "DOCX templates")),
            ("xlsx", t("components.template_selector_tabbed.spreadsheet_templates", "XLSX templates")),
        ):
            page = QWidget(self._tabs)
            layout = QVBoxLayout(page)
            layout.setContentsMargins(0, Spacing.SM, 0, 0)
            template_list = QListWidget(page)
            template_list.setObjectName(f"templateManagementList-{target}")
            template_list.setItemDelegate(TemplateItemDelegate(template_list, draggable=True))
            template_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
            template_list.itemSelectionChanged.connect(self._on_selection_changed)
            template_list.itemChanged.connect(self._on_item_changed)
            template_list.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
            template_list.setDefaultDropAction(Qt.DropAction.MoveAction)
            template_list.setTextElideMode(Qt.TextElideMode.ElideMiddle)
            template_list.setSpacing(2)
            template_list.model().rowsMoved.connect(
                lambda *args, tt=target: QTimer.singleShot(0, lambda: self._persist_order(tt))
            )
            layout.addWidget(template_list, 1)
            self._lists[target] = template_list
            index = self._tabs.addWidget(page)
            self._pivot.addItem(target, title, onClick=lambda _checked=False, i=index: self._tabs.setCurrentIndex(i))
        self._pivot.setCurrentItem("docx")
        self._tabs.currentChanged.connect(self._on_target_changed)
        root.addWidget(self._tabs, 1)

        primary_row = QGridLayout()
        primary_row.setSpacing(Spacing.XS)
        self._import_button = self._button(
            t("settings.templates.import", "Import templates"),
            self._import_templates,
            primary=True,
        )
        self._import_button.setObjectName("templateImportButton")
        self._copy_button = self._button(
            t("settings.templates.copy_edit", "Copy and edit"),
            self._copy_and_edit,
        )
        self._edit_button = self._button(t("settings.templates.edit", "Edit"), self._edit_selected)
        primary_row.addWidget(self._import_button, 0, 0)
        primary_row.addWidget(self._copy_button, 0, 1)
        primary_row.addWidget(self._edit_button, 0, 1)
        self._more_button = PushButton(t("settings.templates.more", "More"), self)
        menu = QMenu(self._more_button)
        for label, callback, icon_name in (
            (t("settings.templates.rename", "Rename"), self._rename_selected, "text.svg"),
            (t("settings.templates.export", "Export"), self._export_selected, "export.svg"),
            (t("settings.templates.delete", "Delete"), self._delete_selected, "delete.svg"),
        ):
            set_action_icon(menu.addAction(label, callback), icon_name)
        self._more_button.setMenu(menu)
        primary_row.addWidget(self._more_button, 0, 2)
        root.addLayout(primary_row)

        secondary_row = QGridLayout()
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
        secondary_row.addWidget(self._up_button, 0, 0)
        secondary_row.addWidget(self._down_button, 0, 1)
        secondary_row.addWidget(self._folder_button, 1, 0, 1, 2)
        secondary_row.addWidget(self._refresh_button, 1, 2)
        self._default_button = self._button(t("settings.templates.set_default", "Set as default"), self._set_default)
        secondary_row.addWidget(self._default_button, 0, 2)
        root.addLayout(secondary_row)
        for button, icon_name in (
            (self._copy_button, "copy.svg"),
            (self._more_button, "other.svg"),
            (self._up_button, "move_up.svg"),
            (self._down_button, "move_down.svg"),
            (self._folder_button, "open_folder.svg"),
            (self._refresh_button, "refresh.svg"),
        ):
            set_action_icon(button, icon_name)

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

        self._import_summary = QLabel(self)
        self._import_summary.setObjectName("templateImportSummary")
        self._import_summary.setWordWrap(True)
        self._import_summary.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self._import_summary.hide()
        root.addWidget(self._import_summary)

    def _button(self, text: str, callback, *, primary: bool = False) -> QPushButton:
        button = ActionButton(text, self)
        button.setMinimumHeight(Sizing.CONTROL_HEIGHT)
        button.clicked.connect(callback)
        apply_theme_class(button, "primary" if primary else "secondary")
        return button

    def _on_target_changed(self, _index: int) -> None:
        self._pivot.setCurrentItem(self._target())
        self._sync_actions()

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
        self._vm.refresh()
        self._render(select_id=select_id)

    def focus_template(self, target: str, template_id: str | None) -> None:
        self._tabs.setCurrentIndex(1 if target == "xlsx" else 0)
        self._render(select_id=template_id)

    def _render(self, *, select_id: str | None = None) -> None:
        self._catalog_error.setVisible(self._vm.error is not None)
        for target, widget in self._lists.items():
            remembered = select_id if select_id is not None else self._current_ids.get(target)
            templates = [item for item in self._vm.templates if item.target == target]
            with QSignalBlocker(widget):
                widget.clear()
                for template in templates:
                    custom = self._manager.is_custom(template)
                    source = (
                        t("settings.templates.custom", "Custom")
                        if custom
                        else t("settings.templates.builtin", "Built-in")
                    )
                    item = QListWidgetItem(template.name)
                    item.setData(SOURCE_ROLE, source)
                    item.setData(CUSTOM_ROLE, custom)
                    item.setData(DEFAULT_ROLE, self._vm.defaults.get(target) == template.id)
                    item.setData(Qt.ItemDataRole.AccessibleDescriptionRole, source)
                    item.setData(_TEMPLATE_ID_ROLE, template.id)
                    item.setData(_TEMPLATE_CUSTOM_ROLE, custom)
                    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                    item.setCheckState(
                        Qt.CheckState.Checked if self._vm.enabled.get(template.id, True) else Qt.CheckState.Unchecked
                    )
                    item.setToolTip(
                        f"{template.name}\n{template.path}"
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
        self._vm.refresh()

    def _sync_actions(self) -> None:
        item = self._selected_item()
        has_selection = item is not None
        custom = bool(item.data(_TEMPLATE_CUSTOM_ROLE)) if item is not None else False
        self._copy_button.setVisible(not custom)
        self._edit_button.setVisible(custom)
        self._more_button.setEnabled(has_selection and custom)
        self._copy_button.setEnabled(has_selection and not custom)
        self._edit_button.setEnabled(has_selection and custom)
        row = self._lists[self._target()].currentRow()
        self._up_button.setEnabled(has_selection and row > 0)
        self._down_button.setEnabled(has_selection and row < self._lists[self._target()].count() - 1)
        self._default_button.setEnabled(
            has_selection
            and bool(self._vm.enabled.get(self._selected_id() or "", False))
            and self._vm.defaults.get(self._target()) != self._selected_id()
        )

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
        succeeded = 0
        cancelled = 0
        for path in paths:
            try:
                conflict = self._manager.import_conflict(path)
                replace_id = None
                if conflict is not None:
                    choice = feedback.choose(
                        t("settings.templates.import"),
                        t(
                            "settings.templates.name_conflict",
                            "A custom template named {name} already exists.",
                            name=conflict.name,
                        ),
                        choices=[
                            feedback.FeedbackChoice("keep", t("settings.templates.keep_both"), primary=True),
                            feedback.FeedbackChoice("replace", t("settings.templates.replace"), role="accept"),
                            feedback.FeedbackChoice("cancel", t("common.cancel"), role="reject"),
                        ],
                        parent=self,
                        default="keep",
                        danger=True,
                    )
                    if choice not in {"keep", "replace"}:
                        cancelled += 1
                        continue
                    if choice == "replace":
                        replace_id = conflict.id
                imported = self._manager.import_template(path, replace_id=replace_id)
                succeeded += 1
                last_id = imported.id
                self._tabs.setCurrentIndex(1 if imported.target == "xlsx" else 0)
            except Exception as exc:
                errors.append(f"{Path(path).name}: {exc}")
        self.refresh(select_id=last_id)
        summary = t(
            "settings.templates.import_summary",
            "Imported: {succeeded}. Failed: {failed}. Cancelled: {cancelled}.",
            succeeded=succeeded,
            failed=len(errors),
            cancelled=cancelled,
        )
        self._import_summary.setText(summary)
        self._import_summary.show()
        if errors:
            from docwen_gui.diagnostics import DiagnosticSummary

            report = feedback.warn if succeeded else feedback.error
            report(
                t("settings.templates.import"),
                summary,
                details="\n".join(errors),
                parent=self,
                copyable=True,
                diagnostic=DiagnosticSummary(
                    status="partial" if succeeded else "failed",
                    succeeded_count=succeeded,
                    failed_count=len(errors),
                    cancelled_count=cancelled,
                ),
            )

    def _copy_and_edit(self) -> None:
        template_id = self._selected_id()
        if template_id is None:
            return
        template = self._vm.find(template_id)
        if template is None:
            return
        name, accepted = self._ask_name(
            t("settings.templates.copy_edit"), template.name + " - " + t("settings.templates.custom")
        )
        if not accepted:
            return
        try:
            copied = self._manager.copy_builtin_as_custom(template_id, custom_name=name)
        except Exception as exc:
            self._show_error(exc)
            return
        self.refresh(select_id=copied.id)
        try:
            self._open_file(copied.path)
        except Exception as exc:
            self._show_error(exc)

    def _ask_name(self, title: str, initial: str) -> tuple[str, bool]:
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setMinimumWidth(420)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(Spacing.MD, Spacing.MD, Spacing.MD, Spacing.MD)
        label = QLabel(t("settings.templates.name_prompt"), dialog)
        editor = QLineEdit(initial, dialog)
        label.setBuddy(editor)
        editor.setMinimumHeight(Sizing.CONTROL_HEIGHT)
        editor.selectAll()
        layout.addWidget(label)
        layout.addWidget(editor)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, dialog)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText(t("common.ok"))
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText(t("common.cancel"))
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        editor.setFocus()
        accepted = dialog.exec() == QDialog.DialogCode.Accepted
        return editor.text(), accepted

    def _on_application_state_changed(self, state: Qt.ApplicationState) -> None:
        if state == Qt.ApplicationState.ApplicationActive and self.isVisible():
            self.refresh()

    def _edit_selected(self) -> None:
        template = self._selected_template()
        if template is not None and self._manager.is_custom(template):
            try:
                self._open_file(template.path)
            except Exception as exc:
                self._show_error(exc)

    def _rename_selected(self) -> None:
        template = self._selected_template()
        if template is None or not self._manager.is_custom(template):
            return
        name, accepted = self._ask_name(t("settings.templates.rename"), template.name)
        if not accepted:
            return
        try:
            renamed = self._manager.rename_custom(template.id, name)
        except Exception as exc:
            self._show_error(exc)
            return
        self.refresh(select_id=renamed.id)

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
        prompt = QMessageBox(self)
        prompt.setWindowTitle(t("settings.templates.delete"))
        prompt.setText(t("settings.templates.delete_confirm", name=template.name))
        delete = prompt.addButton(t("settings.templates.delete"), QMessageBox.ButtonRole.DestructiveRole)
        cancel = prompt.addButton(t("common.cancel"), QMessageBox.ButtonRole.RejectRole)
        prompt.setDefaultButton(cancel)
        prompt.exec()
        if prompt.clickedButton() is not delete:
            return
        widget = self._lists[self._target()]
        row = widget.currentRow()
        adjacent = widget.item(row + 1) if row + 1 < widget.count() else widget.item(row - 1)
        next_id = str(adjacent.data(_TEMPLATE_ID_ROLE)) if adjacent is not None else None
        try:
            self._manager.delete_custom(template.id)
        except Exception as exc:
            self._show_error(exc)
            return
        self.refresh(select_id=next_id)

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

    def _open_user_folder(self) -> None:
        try:
            directory = self._manager.ensure_user_directory()
        except Exception as exc:
            self._show_error(exc)
            return
        from docwen_gui.path_actions import open_path

        result = open_path(directory)
        if not result.success:
            self._show_error(
                TemplateManagementError(
                    t("main_window.open_path_failed", "Failed to open path: {path}", path=str(directory))
                )
            )

    def _selected_template(self):
        template_id = self._selected_id()
        if template_id is None:
            return None
        return self._vm.find(template_id)

    @staticmethod
    def _open_file(path: Path) -> None:
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(path))):
            raise TemplateManagementError(
                t("settings.templates.open_failed", "Could not open the template. Check its associated application.")
            )

    def _show_catalog_error(self, detail: str) -> None:
        for widget in self._lists.values():
            widget.clear()
        self._sync_actions()
        self._catalog_error.setVisible(True)
        self._catalog_error.setToolTip(detail)

    def _set_default(self) -> None:
        template_id = self._selected_id()
        if template_id:
            try:
                self._manager.set_default(template_id)
                self.refresh(select_id=template_id)
            except Exception as exc:
                self._show_error(exc)

    def _persist_order(self, target: str) -> None:
        widget = self._lists[target]
        ids = [str(widget.item(row).data(_TEMPLATE_ID_ROLE)) for row in range(widget.count())]
        try:
            self._manager.set_order(target, ids)
            self.refresh()
        except Exception as exc:
            self._show_error(exc)
            self.refresh()

    def _show_error(self, error: Exception) -> None:
        logger.exception("Template management operation failed", exc_info=error)
        from docwen_gui.diagnostics import DiagnosticSummary
        from docwen_gui.dialogs.feedback import error as show_error

        message = str(error) if isinstance(error, TemplateManagementError) else t("settings.templates.operation_failed")
        show_error(
            t("common.error", "Error"),
            message,
            details=str(error),
            parent=self,
            diagnostic=DiagnosticSummary.from_exception(error),
        )


__all__ = ["TemplatesTab"]
