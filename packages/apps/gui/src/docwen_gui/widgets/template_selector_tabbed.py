"""
选项卡式模板选择组件 — 基于 PySide6 + qfluentwidgets。

管理多个模板类型（docx/xlsx）的选项卡式选择器，支持外部数据注入和刷新。

架构约定：
- 模板数据由外层通过 ``load_templates()`` 注入，组件内不直接访问文件系统。
- 刷新逻辑由外层控制（例如使用 QThread），组件仅提供接收数据的接口。
"""

from __future__ import annotations

import logging
import re
from collections.abc import Callable
from typing import Any

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QStackedWidget, QVBoxLayout, QWidget
from qfluentwidgets import Pivot

from ..i18n import t
from .template_selector import (
    TemplateItemDetails,
    TemplateSelectionFeedback,
    TemplateSelector,
)

logger = logging.getLogger(__name__)


# ── Sort helper ──────────────────────────────────────────────────────────────

_SORT_TOKEN_RE = re.compile(r"\d+|\D+")


def _template_name_sort_key(name: str) -> tuple[object, ...]:
    parts: list[tuple[int, object]] = []
    for token in _SORT_TOKEN_RE.findall(str(name)):
        if token.isdigit():
            parts.append((0, int(token)))
        else:
            parts.append((1, token.casefold()))
    return tuple(parts)


# ── Widget ───────────────────────────────────────────────────────────────────


class TabbedTemplateSelector(QWidget):
    """选项卡式模板选择组件。

    包含两个选项卡（docx/xlsx），每个选项卡内嵌一个
    :class:`TemplateSelector`。支持模板选中跟踪、手动选择记忆
    和外部模板数据注入。

    参数：
        parent: 父组件。
        on_template_selected: 选中回调 ``fn(template_type: str, template_name: str)``。
        on_tab_changed: 选项卡切换回调 ``fn(new_tab: str, old_tab: str)``。
        on_open_location: 打开模板文件位置回调 ``fn(template_type, template_name)``。
        on_open_directory: 打开模板目录回调 ``fn(template_type)``。
    """

    template_selected = Signal(str, str)
    """选中模板时发出: (template_type, template_name)。"""

    tab_changed = Signal(str, str)
    """选项卡切换时发出: (new_tab, old_tab)。"""

    def __init__(
        self,
        parent: QWidget | None = None,
        on_template_selected: Callable[[str, str], None] | None = None,
        on_tab_changed: Callable[[str, str], None] | None = None,
        on_open_location: Callable[[str, str], None] | None = None,
        on_open_directory: Callable[[str], None] | None = None,
    ):
        super().__init__(parent)
        self.setObjectName("tabbedTemplateSelector")
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self._on_template_selected_cb = on_template_selected
        self._on_tab_changed_cb = on_tab_changed
        self._on_open_location_cb = on_open_location
        self._on_open_directory_cb = on_open_directory
        self._current_tab: str = "docx"
        self._template_cache: dict[str, list[str]] = {}
        self._template_details_cache: dict[str, dict[str, TemplateItemDetails]] = {}
        self._last_selection_feedback: tuple[str, str, TemplateSelectionFeedback] | None = None
        self._selection_callback_contexts: list[tuple[str, str, TemplateSelectionFeedback] | None] = []
        self._manual_selection: tuple[str, str] | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # ── Tab bar ──────────────────────────────────────────────────────
        self._tab_titles: dict[str, str] = {}
        self._pivot = Pivot(self)
        self._pivot.setObjectName("templateSelectorPivot")
        layout.addWidget(self._pivot, 0)

        # ── Stacked widget ───────────────────────────────────────────────
        self._stack = QStackedWidget(self)
        self._stack.setObjectName("templateSelectorStack")
        layout.addWidget(self._stack, 1)

        # ── 创建两个选项卡 ────────────────────────────────────────────────
        self._selectors: dict[str, TemplateSelector] = {}
        for template_type, label_key in [
            ("docx", "components.template_selector_tabbed.document_templates"),
            ("xlsx", "components.template_selector_tabbed.spreadsheet_templates"),
        ]:
            title = t(label_key)
            self._tab_titles[template_type] = title
            selector = TemplateSelector(
                template_type=template_type,
                on_template_selected=lambda name, tt=template_type: self._on_selector_selected(tt, name),
                on_open_location=self._on_open_location_cb,
                on_open_directory=self._on_open_directory_cb,
            )
            # Forward the template_error signal
            selector.template_error.connect(self._forward_template_error)
            self._selectors[template_type] = selector
            self._stack.addWidget(selector)
            self._pivot.addItem(
                template_type,
                title,
                onClick=lambda _checked=False, tt=template_type: self._set_current_tab(tt, emit_signal=True),
            )
            item = getattr(self._pivot, "items", {}).get(template_type)
            if item is not None:
                item.setToolTip(title)

        self._set_current_tab("docx", emit_signal=False)
        self._refresh_accessibility()

    # ── Public API ────────────────────────────────────────────────────────────

    def activate_and_select(self, template_type: str) -> bool:
        """激活指定选项卡并选中第一个模板。

        Returns:
            True 表示成功选中。
        """
        resolved_type = (
            template_type if template_type in self._selectors else next(iter(self._selectors.keys()), "docx")
        )
        self._set_current_tab(resolved_type, emit_signal=False)
        selector = self._selectors.get(resolved_type)
        if selector and selector._list.count() > 0:
            selector.activate_first_template(
                selection_source="auto_default",
                explanation=self._build_auto_default_reason(resolved_type),
            )
            return True
        return False

    def ensure_preferred_selection(self, template_type: str) -> bool:
        """确保目标页有选中模板：保留现有选择，再恢复手选或激活默认。

        Returns:
            True 表示已有选中。
        """
        selector = self._selectors.get(template_type)
        selected = selector.get_selected() if selector is not None else None
        if selector is not None and selector.has_template(selected):
            self._set_current_tab(template_type, emit_signal=False)
            return True
        if self._restore_manual_selection(template_type):
            return True
        return self.activate_and_select(template_type)

    def set_selection_callback(self, callback: Callable[[str, str], None] | None) -> None:
        """设置模板选中回调。"""
        self._on_template_selected_cb = callback

    def get_selected_template(self) -> tuple[str, str] | None:
        """返回 ``(template_type, template_name)`` 或 ``None``。"""
        selector = self._selectors.get(self._current_tab)
        if selector:
            name = selector.get_selected()
            if name:
                return (self._current_tab, name)
        # Fallback: search across tabs
        for tt, sel in self._selectors.items():
            name = sel.get_selected()
            if name:
                return (tt, name)
        return None

    def get_selected_template_resource(self) -> tuple[str, str] | None:
        """Return ``(template_type, canonical_resource_id)`` for the selection."""
        selector = self._selectors.get(self._current_tab)
        if selector:
            resource_id = selector.get_selected_resource_id()
            if resource_id:
                return (self._current_tab, resource_id)
        for template_type, candidate in self._selectors.items():
            resource_id = candidate.get_selected_resource_id()
            if resource_id:
                return (template_type, resource_id)
        return None

    def consume_last_selection_feedback(
        self,
    ) -> tuple[str, str, TemplateSelectionFeedback] | None:
        """取出并清空最近一次选中回执。"""
        feedback = self._last_selection_feedback
        self._last_selection_feedback = None
        return feedback

    def peek_callback_selection_feedback(
        self,
    ) -> tuple[str, str, TemplateSelectionFeedback] | None:
        """查看当前 callback 的回执，不受公共 mailbox 消费影响。"""
        return self._selection_callback_contexts[-1] if self._selection_callback_contexts else None

    @property
    def current_tab(self) -> str:
        """当前激活的选项卡键（``"docx"`` 或 ``"xlsx"``）。"""
        return self._current_tab

    def restore_current_tab(self, template_type: str) -> str:
        """按持久化类型恢复活动页，不产生用户选择信号。"""
        resolved_type = template_type if template_type in self._selectors else next(iter(self._selectors), "docx")
        self._set_current_tab(resolved_type, emit_signal=False)
        return resolved_type

    def load_templates(
        self,
        template_type: str,
        names: list[str],
        *,
        details: dict[str, TemplateItemDetails] | None = None,
    ) -> None:
        """加载指定类型的模板列表（由外部数据源提供）。

        Args:
            template_type: 模板类型。
            names: 模板名称列表。
            details: 模板元数据。
        """
        selector = self._selectors.get(template_type)
        if selector is None:
            return
        sorted_names = sorted(names, key=_template_name_sort_key)
        normalized_details = dict(details or {})
        if sorted_names != self._template_cache.get(
            template_type
        ) or normalized_details != self._template_details_cache.get(template_type):
            self._template_cache[template_type] = sorted_names
            self._template_details_cache[template_type] = normalized_details
            selected = selector.get_selected()
            manual_name = None
            if self._manual_selection is not None and self._manual_selection[0] == template_type:
                manual_name = self._manual_selection[1]
            selector.add_templates(sorted_names, template_details=normalized_details)
            if manual_name and selector.has_template(manual_name):
                selector.select_template(manual_name, selection_source="restore")
            elif selected and selected in sorted_names:
                selector.select_template(selected, selection_source="restore")
            elif sorted_names:
                selector.activate_first_template(selection_source="restore")
        if self._manual_selection is not None and self._manual_selection[0] == template_type:
            self._restore_manual_selection()

    def load_all_templates(
        self,
        data: dict[str, list[str]],
        *,
        details: dict[str, dict[str, TemplateItemDetails]] | None = None,
    ) -> None:
        """加载全部类型的模板。

        Args:
            data: ``{template_type: [name, ...]}``。
            details: ``{template_type: {name: TemplateItemDetails}}``。
        """
        for tt in self._selectors:
            names = data.get(tt, [])
            tt_details = (details or {}).get(tt, {})
            self.load_templates(tt, names, details=tt_details)

    def show_load_error(self, summary: str, detail: str) -> None:
        """Project one registry failure into every template tab."""

        self._template_cache.clear()
        self._template_details_cache.clear()
        self._manual_selection = None
        for selector in self._selectors.values():
            selector.show_load_error(summary, detail)
        self._refresh_accessibility()

    def get_selector(self, template_type: str) -> TemplateSelector | None:
        """获取指定类型的内部 TemplateSelector 引用（用于测试/高级用法）。"""
        return self._selectors.get(template_type)

    # ── Focus ─────────────────────────────────────────────────────────────────

    def focusInEvent(self, event: Any) -> None:
        super().focusInEvent(event)
        selector = self._selectors.get(self._current_tab)
        if selector is not None:
            selector.setFocus(Qt.FocusReason.TabFocusReason)

    # ── Internal: tab switching ───────────────────────────────────────────────

    def _set_current_tab(self, route_key: str, *, emit_signal: bool) -> None:
        if route_key not in self._selectors:
            return
        old_tab = self._current_tab
        self._current_tab = route_key
        if getattr(self._pivot, "currentRouteKey", lambda: None)() != route_key:
            self._pivot.setCurrentItem(route_key)
        selector = self._selectors.get(route_key)
        if selector is not None:
            self._stack.setCurrentWidget(selector)
        self._refresh_accessibility()
        if emit_signal and old_tab != route_key:
            self.tab_changed.emit(route_key, old_tab)
            if self._on_tab_changed_cb:
                self._on_tab_changed_cb(route_key, old_tab)
            self.ensure_preferred_selection(route_key)

    # ── Internal: selection ───────────────────────────────────────────────────

    def _on_selector_selected(self, template_type: str, template_name: str) -> None:
        selector = self._selectors.get(template_type)
        callback_feedback: tuple[str, str, TemplateSelectionFeedback] | None = None
        if selector:
            feedback = selector.consume_selection_feedback()
            if feedback is not None:
                callback_feedback = (
                    template_type,
                    template_name,
                    feedback,
                )
                self._last_selection_feedback = callback_feedback
                if feedback.selection_source == "user":
                    self._manual_selection = (template_type, template_name)
        self._selection_callback_contexts.append(callback_feedback)
        try:
            self._refresh_accessibility()
            self.template_selected.emit(template_type, template_name)
            if self._on_template_selected_cb:
                self._on_template_selected_cb(template_type, template_name)
        finally:
            self._selection_callback_contexts.pop()

    def _restore_manual_selection(self, template_type: str | None = None) -> bool:
        if self._manual_selection is None:
            return False
        manual_type, template_name = self._manual_selection
        if template_type is not None and manual_type != template_type:
            return False
        selector = self._selectors.get(manual_type)
        if selector is None or not selector.has_template(template_name):
            return False
        self._set_current_tab(manual_type, emit_signal=False)
        selector.select_template(template_name, selection_source="restore")
        return True

    def _build_auto_default_reason(self, template_type: str) -> str:
        template_kind = t(
            "components.template_selector_tabbed.document_templates"
            if template_type == "docx"
            else "components.template_selector_tabbed.spreadsheet_templates"
        )
        return t(
            "components.template_selector.auto_selected_reason",
            template_kind=template_kind,
        )

    # ── Internal: helpers ─────────────────────────────────────────────────────

    def _forward_template_error(self, summary: str, detail: str) -> None:
        """Forward error from child TemplateSelector. Connected as signal handler."""
        # Just log — TemplateSelector's own signal is sufficient for UI binding.
        logger.debug("TemplateSelector error: %s — %s", summary, detail)

    def _refresh_accessibility(self) -> None:
        current_label = self._tab_titles.get(self._current_tab, "")
        selector = self._selectors.get(self._current_tab)
        selection_description = selector.get_selected() if selector is not None else ""
        description = selection_description or t("components.template_selector.empty_hint")
        self.setAccessibleName(current_label)
        self.setAccessibleDescription(description)
        self._pivot.setAccessibleName(current_label)
