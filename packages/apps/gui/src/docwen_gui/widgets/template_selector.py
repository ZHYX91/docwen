"""
模板列表组件 — 基于 PySide6 + qfluentwidgets。

提供单个模板类型的列表显示和选择功能，支持模板列表展示、选择和文件位置打开。

架构约定：
- 模板数据通过 ``add_templates()`` 从外部注入，本组件不直接加载模板。
- 文件夹/位置操作通过回调（``on_open_location`` / ``on_open_directory``）
  委托给外部，本组件不直接访问文件系统。
- 反馈通知通过 ``template_error`` 信号发出，由外层负责渲染。
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

from PySide6.QtCore import QSize, Qt, Signal
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QToolButton,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import PrimaryPushButton

from ..i18n import t
from ..styles.design_tokens import Spacing

logger = logging.getLogger(__name__)


# ── Data types ───────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class TemplateItemDetails:
    """单个模板的元数据。"""

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
        on_open_location: 打开模板文件位置回调 ``fn(template_type, template_name)``。
            未提供时，打开位置按钮隐藏。
        on_open_directory: 打开模板目录回调 ``fn(template_type)``。
    """

    template_selected = Signal(str)
    """选中模板时发出，携带模板名称。"""

    template_error = Signal(str, str)
    """操作出错时发出: (summary, detail)。"""

    def __init__(
        self,
        parent: QWidget | None = None,
        template_type: str = "docx",
        on_template_selected: Callable[[str], None] | None = None,
        on_open_location: Callable[[str, str], None] | None = None,
        on_open_directory: Callable[[str], None] | None = None,
    ):
        super().__init__(parent)
        self.setObjectName("templateSelectorRoot")
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.template_type = template_type
        self._on_template_selected_cb = on_template_selected
        self._on_open_location_cb = on_open_location
        self._on_open_directory_cb = on_open_directory
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

        self._empty_action_button = PrimaryPushButton(t("components.template_selector.open_template_dir"))
        self._empty_action_button.setObjectName("templateSelectorEmptyActionButton")
        self._empty_action_button.setMinimumHeight(32)
        self._empty_action_button.clicked.connect(self._open_template_directory)
        empty_layout.addWidget(self._empty_action_button, alignment=Qt.AlignmentFlag.AlignHCenter)
        layout.addWidget(self._empty_state)

        # ── 模板列表 ────────────────────────────────────────────────────
        self._list = QListWidget()
        self._list.setObjectName("templateSelectorList")
        self._list.setTextElideMode(Qt.TextElideMode.ElideMiddle)
        self._list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._list.currentItemChanged.connect(self._on_item_changed)
        self._list.itemActivated.connect(self._on_item_activated)
        self._list.customContextMenuRequested.connect(self._show_list_context_menu)
        layout.addWidget(self._list)

        # ── 底部信息栏 ───────────────────────────────────────────────────
        self._footer_row = QWidget(self)
        self._footer_row.setObjectName("templateSelectorFooterRow")
        footer_layout = QHBoxLayout(self._footer_row)
        footer_layout.setContentsMargins(Spacing.SM, Spacing.XS, Spacing.XS, Spacing.XS)
        footer_layout.setSpacing(Spacing.XS)

        self._details_label = QLabel()
        self._details_label.setObjectName("templateSelectorMetaLabel")
        self._details_label.setWordWrap(True)
        self._details_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
        self._details_label.setVisible(False)
        footer_layout.addWidget(self._details_label, 1)

        self._open_location_button = QToolButton(self._footer_row)
        self._open_location_button.setObjectName("templateSelectorOpenButton")
        self._open_location_button.setToolTip(t("components.template_selector.open_location"))
        self._open_location_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonIconOnly)
        # Use theme-compatible icon
        folder_icon = self.style().standardIcon(self.style().StandardPixmap.SP_DirOpenIcon)
        self._location_icon = folder_icon
        self._open_location_button.setIcon(folder_icon)
        self._open_location_button.setIconSize(QSize(16, 16))
        self._open_location_button.setFixedSize(26, 26)
        self._open_location_button.setAutoRaise(True)
        self._open_location_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self._open_location_button.setEnabled(False)
        self._open_location_button.clicked.connect(self._open_selected_template_location)
        # Hide if no location callback provided
        if self._on_open_location_cb is None:
            self._open_location_button.setVisible(False)
        footer_layout.addWidget(
            self._open_location_button,
            0,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop,
        )
        layout.addWidget(self._footer_row)

        self.setTabOrder(self._list, self._open_location_button)
        self._show_empty_state()

    # ── Public API ────────────────────────────────────────────────────────────

    def add_templates(
        self,
        template_names: list[str],
        auto_select_first: bool = False,
        *,
        template_details: dict[str, TemplateItemDetails] | None = None,
    ) -> None:
        """批量加载模板列表。

        Args:
            template_names: 模板名称列表。
            auto_select_first: 是否在没有手动选择时自动选中第一个。
            template_details: 模板元数据字典（name → TemplateItemDetails）。
        """
        self._load_error = None
        self._empty_label.setText(t("components.template_selector.no_templates"))
        self._empty_hint_label.setText(t("components.template_selector.empty_hint"))
        self._empty_action_button.setEnabled(self._on_open_directory_cb is not None)
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
            # Restore selection after list repopulation
            if self._has_manual_selection() and self._user_selected_template_name in template_names:
                self.select_template(self._user_selected_template_name, selection_source="restore")
            elif auto_select_first and not self._has_manual_selection():
                self.activate_first_template(selection_source="restore")
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
        self._empty_action_button.setEnabled(False)
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
        details = self._template_details.get(name)
        return details.resource_id if details is not None else None

    def activate_first_template(
        self,
        *,
        selection_source: str = "restore",
        explanation: str | None = None,
    ) -> str | None:
        """激活首个模板，返回模板名称。"""
        if self._list.count() <= 0:
            return None
        first_name = self._get_item_template_name(self._list.item(0))
        if not first_name:
            return None
        self._select_row(0, selection_source=selection_source, explanation=explanation)
        return first_name

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
        self._user_selected_template_name = None  # reset manual tracking on clear
        self._show_empty_state()

    # ── Internal: visibility ──────────────────────────────────────────────────

    def _show_template_list(self) -> None:
        self._empty_state.setVisible(False)
        self._list.setVisible(True)
        self._footer_row.setVisible(bool(self._details_label.text()) or self._open_location_button.isEnabled())
        self._details_label.setVisible(bool(self._details_label.text()))

    def _show_empty_state(self) -> None:
        self._list.setVisible(False)
        self._empty_state.setVisible(True)
        self._clear_selection_labels()
        self._footer_row.setVisible(False)
        self._sync_open_location_button_state(None)

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
        self._sync_open_location_button_state(name)
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
        self._sync_open_location_button_state(None)
        self._clear_selection_labels()

    def _clear_selection_labels(self) -> None:
        self._details_label.clear()
        self._details_label.setVisible(False)
        self._details_label.setToolTip("")
        self._footer_row.setVisible(False)

    # ── Event handlers ────────────────────────────────────────────────────────

    def focusInEvent(self, event: Any) -> None:
        super().focusInEvent(event)
        self._focus_primary_control()

    def _focus_primary_control(self) -> None:
        if self._list.isVisible() and self._list.count() > 0 and self._list.isEnabled():
            self._list.setFocus(Qt.FocusReason.TabFocusReason)
            return
        if self._empty_action_button.isVisible() and self._empty_action_button.isEnabled():
            self._empty_action_button.setFocus(Qt.FocusReason.TabFocusReason)

    def _on_item_changed(self, current: QListWidgetItem | None, _previous: QListWidgetItem | None) -> None:
        if current is None:
            self._sync_open_location_button_state(None)
            return
        name = self._get_item_template_name(current)
        if not name:
            self._sync_open_location_button_state(None)
            return
        source = self._pending_selection_source or "user"
        explanation = self._pending_selection_explanation
        self._pending_selection_source = None
        self._pending_selection_explanation = None
        self._apply_selection(name, selection_source=source, explanation=explanation)

    def _on_item_activated(self, item: QListWidgetItem) -> None:
        name = self._get_item_template_name(item)
        if name:
            self._open_template_location(name)

    def _show_list_context_menu(self, position: Any) -> None:
        item = self._list.itemAt(position) or self._list.currentItem()
        if item is None:
            return
        if item is not self._list.currentItem():
            self._list.setCurrentItem(item)

        menu = QMenu(self._list)
        template_name = self._get_item_template_name(item)
        if template_name and self._on_open_location_cb is not None:
            action = menu.addAction(self._open_location_button.toolTip())
            action.triggered.connect(  # type: ignore[attr-defined]
                lambda _checked: self._open_template_location(template_name or "")
            )
        menu.exec(self._list.viewport().mapToGlobal(position))

    # ── Location actions ──────────────────────────────────────────────────────

    def _open_selected_template_location(self) -> None:
        template_name = self.get_selected()
        if not template_name:
            return
        self._open_template_location(template_name)

    def _open_template_directory(self) -> None:
        if self._on_open_directory_cb is not None:
            self._on_open_directory_cb(self.template_type)
        else:
            self.template_error.emit(
                t("components.template_selector.unavailable"),
                t("components.template_selector.no_templates"),
            )

    def _open_template_location(self, template_name: str) -> None:
        if self._on_open_location_cb is not None:
            self._on_open_location_cb(self.template_type, template_name)
        else:
            self.template_error.emit(
                t("components.template_selector.unavailable"),
                t("components.template_selector.no_templates"),
            )

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _get_item_template_name(item: QListWidgetItem | None) -> str | None:
        if item is None:
            return None
        value = item.data(Qt.ItemDataRole.UserRole)
        return str(value) if value else item.text()

    def _sync_open_location_button_state(self, template_name: str | None) -> None:
        enabled = template_name is not None and self._on_open_location_cb is not None
        self._open_location_button.setEnabled(enabled)

    def _refresh_item_presentations(self) -> None:
        for i in range(self._list.count()):
            item = self._list.item(i)
            name = self._get_item_template_name(item)
            if item is not None and name:
                self._apply_item_presentation(item, name)

    def _apply_item_presentation(self, item: QListWidgetItem, name: str) -> None:
        item.setText(name)
        if self._on_open_location_cb is not None:
            item.setIcon(self._location_icon)
        else:
            item.setIcon(QIcon())
        item.setToolTip(self._build_item_tooltip(name))

    def _build_item_tooltip(self, name: str) -> str:
        lines = [name]
        details = self._template_details.get(name)
        if details:
            if details.usage_hint:
                lines.append(t("components.template_selector.usage_line", value=details.usage_hint))
            if details.source_label:
                lines.append(t("components.template_selector.source_line", value=details.source_label))
            if details.source_path:
                lines.append(t("components.template_selector.source_tooltip", value=details.source_path))
            if details.updated_label:
                lines.append(t("components.template_selector.updated_line", value=details.updated_label))
        if self._on_open_location_cb is not None:
            lines.append(t("components.template_selector.open_location"))
        return "\n".join(lines)

    def _update_selection_labels(self, name: str) -> None:
        details_text = self._build_selection_details_text(name)
        self._details_label.setText(details_text)
        self._details_label.setVisible(bool(details_text))
        tooltip = self._build_selection_details_tooltip(name)
        self._details_label.setToolTip(tooltip)
        self._footer_row.setVisible(bool(details_text) or self._open_location_button.isEnabled())

    def _build_selection_details_text(self, name: str) -> str:
        details = self._template_details.get(name)
        if not details:
            return ""
        lines: list[str] = []
        if details.source_label:
            lines.append(t("components.template_selector.source_line", value=details.source_label))
        if details.updated_label:
            lines.append(
                t(
                    "components.template_selector.updated_line",
                    value=details.updated_label,
                )
            )
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
