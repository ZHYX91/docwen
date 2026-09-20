"""
模板列表组件 — 基于 PySide6 + qfluentwidgets。

提供单个模板类型的列表显示和选择功能，并提供统一模板管理入口。

架构约定：
- 模板 ID 及展示数据由外部注入，控件不访问文件系统。
- 反馈通知通过 ``template_error`` 信号发出，由外层负责渲染。
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import PushButton

from ..i18n import t
from ..styles.design_tokens import Spacing
from .panel_card import FormRow
from .template_item_delegate import CUSTOM_ROLE, SOURCE_ROLE, TemplateItemDelegate

logger = logging.getLogger(__name__)


# ── Data types ───────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class TemplateItemDetails:
    """单个模板的元数据。"""

    display_name: str | None = None
    resource_id: str | None = None
    usage_hint: str | None = None
    source_label: str | None = None
    source_path: str | None = None
    updated_label: str | None = None


@dataclass(frozen=True)
class TemplateSelectionFeedback:
    """选中操作的回执信息。"""

    selection_source: str  # "user" | "auto_default" | "restore"
    explanation: str | None = None


# ── Widget ───────────────────────────────────────────────────────────────────


class TemplateSelector(QWidget):
    """单个模板类型的列表显示和选择组件。

    参数：
        parent: 父组件。
        template_type: 模板类型（``"docx"`` 或 ``"xlsx"``）。
        on_template_selected: 选中回调 ``fn(template_name: str)``。
        on_manage_templates: 可选的模板管理入口回调 ``fn(template_type)``。
    """

    template_selected = Signal(str)
    """选中模板时发出，携带模板名称。"""

    template_error = Signal(str, str)
    """操作出错时发出: (summary, detail)。"""

    """模板管理改变可见目录后发出。"""

    def __init__(
        self,
        parent: QWidget | None = None,
        template_type: str = "docx",
        on_template_selected: Callable[[str], None] | None = None,
        on_manage_templates: Callable[[str], None] | None = None,
    ):
        super().__init__(parent)
        self.setObjectName("templateSelectorRoot")
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.template_type = template_type
        self._on_template_selected_cb = on_template_selected
        self._on_manage_templates_cb = on_manage_templates
        self._selected: str | None = None
        self._user_selected_template_name: str | None = None
        self._template_details: dict[str, TemplateItemDetails] = {}
        self._selection_feedback: TemplateSelectionFeedback | None = None
        self._pending_selection_source: str | None = None
        self._pending_selection_explanation: str | None = None
        self._load_error: tuple[str, str] | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(Spacing.XS)

        # ── 空状态 ──────────────────────────────────────────────────────
        self._empty_state = QFrame(self)
        self._empty_state.setObjectName("templateSelectorEmptyState")
        empty_layout = QVBoxLayout(self._empty_state)
        empty_layout.setContentsMargins(Spacing.SM, Spacing.SM, Spacing.SM, Spacing.SM)
        empty_layout.setSpacing(Spacing.XS)

        self._empty_label = QLabel(t("components.template_selector.no_templates"))
        self._empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._empty_label.setObjectName("templateSelectorEmptyTitle")
        self._empty_label.setWordWrap(True)
        empty_layout.addWidget(self._empty_label)

        self._empty_hint_label = QLabel(t("components.template_selector.empty_hint"))
        self._empty_hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._empty_hint_label.setObjectName("templateSelectorEmptyHint")
        self._empty_hint_label.setWordWrap(True)
        empty_layout.addWidget(self._empty_hint_label)

        self._empty_manage_button = PushButton(t("components.template_selector.manage_templates", "管理模板"))
        self._empty_manage_button.setObjectName("templateSelectorEmptyManageButton")
        self._empty_manage_button.setMinimumHeight(32)
        self._empty_manage_button.clicked.connect(self._open_template_management)
        empty_layout.addWidget(self._empty_manage_button, alignment=Qt.AlignmentFlag.AlignHCenter)
        layout.addWidget(self._empty_state)

        # ── 模板列表 ────────────────────────────────────────────────────
        self._list = QListWidget()
        self._list.setObjectName("templateSelectorList")
        self._list.setItemDelegate(TemplateItemDelegate(self._list))
        self._list.setTextElideMode(Qt.TextElideMode.ElideMiddle)
        self._list.setWordWrap(True)
        self._list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._list.setResizeMode(QListWidget.ResizeMode.Adjust)
        self._list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._list.currentItemChanged.connect(self._on_item_changed)

        self._list.customContextMenuRequested.connect(self._show_list_context_menu)
        layout.addWidget(self._list, 1)

        # ── 底部信息栏 ───────────────────────────────────────────────────
        self._manage_button = PushButton(t("components.template_selector.manage_templates", "管理模板"), self)
        self._manage_button.setObjectName("templateSelectorManageButton")
        self._manage_button.clicked.connect(self._open_template_management)
        self._footer_row = QWidget(self)
        self._footer_row.setObjectName("templateSelectorFooterRow")
        self._footer_row.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        footer_layout = QVBoxLayout(self._footer_row)
        footer_layout.setContentsMargins(Spacing.SM, Spacing.XS, Spacing.XS, Spacing.XS)
        footer_form = FormRow("", self._manage_button, self._footer_row)
        footer_layout.addWidget(footer_form)
        self._details_label = footer_form.label
        self._details_label.setObjectName("templateSelectorMetaLabel")
        self._details_label.setWordWrap(True)
        self._details_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
        self._details_label.setVisible(False)

        layout.addWidget(self._footer_row)

        self.setTabOrder(self._list, self._manage_button)
        self._show_empty_state()

    # ── Public API ────────────────────────────────────────────────────────────

    def add_templates(
        self,
        template_names: list[str],
        *,
        template_details: dict[str, TemplateItemDetails] | None = None,
    ) -> None:
        """批量加载模板列表。"""
        self._load_error = None
        self._empty_label.setText(t("components.template_selector.no_templates"))
        self._empty_hint_label.setText(t("components.template_selector.empty_hint"))
        self._list.blockSignals(True)
        self._list.clear()
        self._list.blockSignals(False)
        self._template_details = dict(template_details or {})
        if self._selected is not None and self._selected not in template_names:
            self._clear_current_selection()
        for name in template_names:
            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, name)
            self._apply_item_presentation(item, name)
            self._list.addItem(item)
        if template_names:
            self._show_template_list()
            if self._has_manual_selection() and self._user_selected_template_name in template_names:
                self.select_template(self._user_selected_template_name, selection_source="restore")
            elif self._selected is not None and self._selected in template_names:
                self.select_template(self._selected, selection_source="restore")
        else:
            self._show_empty_state()

    def show_load_error(self, summary: str, detail: str) -> None:
        """Render an unavailable state distinct from a successful empty list."""

        self._load_error = (str(summary), str(detail))
        self._list.clear()
        self._template_details.clear()
        self._clear_current_selection()
        self._empty_label.setText(self._load_error[0])
        self._empty_hint_label.setText(self._load_error[1])
        self._show_empty_state()

    def select_template(
        self,
        template_name: str,
        *,
        selection_source: str = "restore",
        explanation: str | None = None,
    ) -> None:
        """选中指定模板。"""
        for i in range(self._list.count()):
            item = self._list.item(i)
            if self._get_item_template_name(item) == template_name:
                self._select_row(i, selection_source=selection_source, explanation=explanation)
                return

    def get_selected(self) -> str | None:
        """获取当前选中的模板名称。"""
        item = self._list.currentItem()
        return self._get_item_template_name(item) if item else None

    def get_selected_resource_id(self) -> str | None:
        """Return the canonical resource ID for the selected display item."""
        name = self.get_selected()
        if name is None:
            return None
        return name

    def has_template(self, template_name: str | None) -> bool:
        """检查指定模板是否在当前列表中。"""
        return bool(template_name) and template_name in {
            self._get_item_template_name(self._list.item(index)) for index in range(self._list.count())
        }

    def _has_manual_selection(self) -> bool:
        return bool(self._user_selected_template_name)

    def consume_selection_feedback(self) -> TemplateSelectionFeedback | None:
        """取出并清空最近一次选中回执。"""
        feedback = self._selection_feedback
        self._selection_feedback = None
        return feedback

    def clear_all(self) -> None:
        """清空模板列表。"""
        self._list.clear()
        self._clear_current_selection()
        self._user_selected_template_name = None
        self._show_empty_state()

    # ── Internal: visibility ──────────────────────────────────────────────────

    def _show_template_list(self) -> None:
        self._empty_state.setVisible(False)
        self._list.setVisible(True)
        self._footer_row.setVisible(True)
        self._details_label.setVisible(bool(self._details_label.text()))

    def _show_empty_state(self) -> None:
        self._list.setVisible(False)
        self._empty_state.setVisible(True)
        self._clear_selection_labels()
        self._footer_row.setVisible(False)

    # ── Internal: selection ───────────────────────────────────────────────────

    def _select_row(self, row: int, *, selection_source: str, explanation: str | None = None) -> None:
        item = self._list.item(row)
        if item is None:
            return
        self._pending_selection_source = selection_source
        self._pending_selection_explanation = explanation
        if self._list.currentRow() == row:
            name = self._get_item_template_name(item)
            if not name:
                return
            self._pending_selection_source = None
            self._pending_selection_explanation = None
            self._apply_selection(name, selection_source=selection_source, explanation=explanation)
            return
        self._list.setCurrentRow(row)

    def _apply_selection(self, name: str, *, selection_source: str, explanation: str | None = None) -> None:
        self._selected = name
        if selection_source == "user":
            self._user_selected_template_name = name
        self._selection_feedback = TemplateSelectionFeedback(selection_source=selection_source, explanation=explanation)
        self._refresh_item_presentations()
        self._update_selection_labels(name)
        logger.debug("模板选中: %s (%s)", name, selection_source)
        self.template_selected.emit(name)
        if self._on_template_selected_cb:
            self._on_template_selected_cb(name)

    def _clear_current_selection(self) -> None:
        self._selected = None
        self._pending_selection_source = None
        self._pending_selection_explanation = None
        self._clear_selection_labels()

    def _clear_selection_labels(self) -> None:
        self._details_label.setText(t("settings.templates.choose_template"))
        self._details_label.setVisible(True)
        self._details_label.setToolTip("")

    # ── Event handlers ────────────────────────────────────────────────────────

    def focusInEvent(self, event: Any) -> None:
        super().focusInEvent(event)
        self._focus_primary_control()

    def _focus_primary_control(self) -> None:
        if self._list.isVisible() and self._list.count() > 0 and self._list.isEnabled():
            self._list.setFocus(Qt.FocusReason.TabFocusReason)
            return
        if self._empty_manage_button.isVisible() and self._empty_manage_button.isEnabled():
            self._empty_manage_button.setFocus(Qt.FocusReason.TabFocusReason)
            return

    def _on_item_changed(self, current: QListWidgetItem | None, _previous: QListWidgetItem | None) -> None:
        if current is None:
            return
        name = self._get_item_template_name(current)
        if not name:
            return
        source = self._pending_selection_source or "user"
        explanation = self._pending_selection_explanation
        self._pending_selection_source = None
        self._pending_selection_explanation = None
        self._apply_selection(name, selection_source=source, explanation=explanation)

    def _show_list_context_menu(self, position: Any) -> None:
        item = self._list.itemAt(position) or self._list.currentItem()
        if item is None:
            return
        if item is not self._list.currentItem():
            self._list.setCurrentItem(item)

        menu = QMenu(self._list)
        manage_action = menu.addAction(t("components.template_selector.manage_templates", "管理模板"))
        manage_action.triggered.connect(lambda _checked=False: self._open_template_management())  # type: ignore[attr-defined]
        menu.exec(self._list.viewport().mapToGlobal(position))

    def _open_template_management(self) -> None:
        if self._on_manage_templates_cb is not None:
            self._on_manage_templates_cb(self.template_type)

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _get_item_template_name(item: QListWidgetItem | None) -> str | None:
        if item is None:
            return None
        value = item.data(Qt.ItemDataRole.UserRole)
        return str(value) if value else item.text()

    def _refresh_item_presentations(self) -> None:
        for i in range(self._list.count()):
            item = self._list.item(i)
            name = self._get_item_template_name(item)
            if item is not None and name:
                self._apply_item_presentation(item, name)

    def _apply_item_presentation(self, item: QListWidgetItem, name: str) -> None:
        details = self._template_details.get(name)
        label = details.display_name if details and details.display_name else name
        item.setText(label)
        item.setData(SOURCE_ROLE, details.source_label if details else None)
        item.setData(CUSTOM_ROLE, bool(details and details.source_label == t("settings.templates.custom")))

        item.setToolTip(self._build_item_tooltip(name))

    def _build_item_tooltip(self, name: str) -> str:
        details = self._template_details.get(name)
        lines = [details.display_name or name] if details else [name]
        if details:
            if details.usage_hint:
                lines.append(t("components.template_selector.usage_line", value=details.usage_hint))
            if details.source_label:
                lines.append(t("components.template_selector.source_line", value=details.source_label))
            if details.source_path:
                lines.append(t("components.template_selector.source_tooltip", value=details.source_path))
            if details.updated_label:
                lines.append(t("components.template_selector.updated_line", value=details.updated_label))
        return "\n".join(lines)

    def _update_selection_labels(self, name: str) -> None:
        details_text = self._build_selection_details_text(name)
        self._details_label.setText(details_text)
        self._details_label.setVisible(bool(details_text))
        tooltip = self._build_selection_details_tooltip(name)
        self._details_label.setToolTip(tooltip)
        self._footer_row.setVisible(True)

    def _build_selection_details_text(self, name: str) -> str:
        details = self._template_details.get(name)
        label = details.display_name if details and details.display_name else name
        lines: list[str] = [t("components.template_selector.selected_name", name=label)]
        if not details:
            return "\n".join(lines)
        if details.source_label:
            lines.append(t("components.template_selector.source_line", value=details.source_label))
        if details.updated_label:
            lines.append(t("components.template_selector.updated_line", value=details.updated_label))
        return "\n".join(lines)

    def _build_selection_details_tooltip(self, name: str) -> str:
        details = self._template_details.get(name)
        if not details:
            return ""
        lines: list[str] = []
        if details.usage_hint:
            lines.append(t("components.template_selector.usage_line", value=details.usage_hint))
        if details.source_label:
            lines.append(t("components.template_selector.source_line", value=details.source_label))
        if details.source_path:
            lines.append(t("components.template_selector.source_tooltip", value=details.source_path))
        if details.updated_label:
            lines.append(t("components.template_selector.updated_line", value=details.updated_label))
        return "\n".join(lines)
