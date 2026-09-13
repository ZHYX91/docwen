"""
选项卡式模板选择组件 — 基于 PySide6 + qfluentwidgets。

管理多个模板类型（docx/xlsx）的选项卡式选择器，支持外部数据注入和刷新。

架构约定：
- 模板数据由外层通过 ``load_templates()`` 注入；调用方顺序即展示顺序。
- 模板管理完成后可以通过 ``refresh_from_runtime_registry()`` 刷新两类模板。
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
from ..styles.design_tokens import Spacing
from .template_selector import (
    TemplateItemDetails,
    TemplateSelectionFeedback,
    TemplateSelector,
)

logger = logging.getLogger(__name__)


# ── Legacy natural-sort helper ───────────────────────────────────────────────
# Kept importable for compatibility.  Template ordering is now owned by the
# runtime template state and callers; the selector must not override it.
_SORT_TOKEN_RE = re.compile(r"\d+|\D+")


def _template_name_sort_key(name: str) -> tuple[object, ...]:
    parts: list[tuple[int, object]] = []
    for token in _SORT_TOKEN_RE.findall(str(name)):
        if token.isdigit():
            parts.append((0, int(token)))
        else:
            parts.append((1, token.casefold()))
    return tuple(parts)


class TabbedTemplateSelector(QWidget):
    """选项卡式模板选择组件。"""

    template_selected = Signal(str, str)
    tab_changed = Signal(str, str)

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
        layout.setSpacing(Spacing.GROUP_GAP)

        self._tab_titles: dict[str, str] = {}
        self._pivot = Pivot(self)
        self._pivot.setObjectName("templateSelectorPivot")
        layout.addWidget(self._pivot, 0)

        self._stack = QStackedWidget(self)
        self._stack.setObjectName("templateSelectorStack")
        layout.addWidget(self._stack, 1)

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
        resolved_type = (
            template_type if template_type in self._selectors else next(iter(self._selectors.keys()), "docx")
        )
        self._set_current_tab(resolved_type, emit_signal=False)
        return self._activate_default_template(resolved_type, selection_source="auto_default")

    def _activate_default_template(self, template_type: str, *, selection_source: str) -> bool:
        selector = self._selectors.get(template_type)
        if selector and selector._list.count() > 0:
            preferred_name = t("meta.template_name") if template_type == "docx" else ""
            explanation = self._build_auto_default_reason(template_type)
            if preferred_name and selector.has_template(preferred_name):
                selector.select_template(preferred_name, selection_source=selection_source, explanation=explanation)
            else:
                selector.activate_first_template(selection_source=selection_source, explanation=explanation)
            return True
        return False

    def ensure_preferred_selection(self, template_type: str) -> bool:
        selector = self._selectors.get(template_type)
        selected = selector.get_selected() if selector is not None else None
        if selector is not None and selector.has_template(selected):
            self._set_current_tab(template_type, emit_signal=False)
            return True
        if self._restore_manual_selection(template_type):
            return True
        return self.activate_and_select(template_type)

    def set_selection_callback(self, callback: Callable[[str, str], None] | None) -> None:
        self._on_template_selected_cb = callback

    def get_selected_template(self) -> tuple[str, str] | None:
        selector = self._selectors.get(self._current_tab)
        if selector:
            name = selector.get_selected()
            if name:
                return (self._current_tab, name)
        for tt, sel in self._selectors.items():
            name = sel.get_selected()
            if name:
                return (tt, name)
        return None

    def get_selected_template_resource(self) -> tuple[str, str] | None:
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

    def consume_last_selection_feedback(self) -> tuple[str, str, TemplateSelectionFeedback] | None:
        feedback = self._last_selection_feedback
        self._last_selection_feedback = None
        return feedback

    def peek_callback_selection_feedback(self) -> tuple[str, str, TemplateSelectionFeedback] | None:
        return self._selection_callback_contexts[-1] if self._selection_callback_contexts else None

    @property
    def current_tab(self) -> str:
        return self._current_tab

    def restore_current_tab(self, template_type: str) -> str:
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
        """Load templates while preserving the caller/runtime order."""

        selector = self._selectors.get(template_type)
        if selector is None:
            return
        ordered_names = list(dict.fromkeys(names))
        normalized_details = dict(details or {})
        if ordered_names != self._template_cache.get(
            template_type
        ) or normalized_details != self._template_details_cache.get(template_type):
            self._template_cache[template_type] = ordered_names
            self._template_details_cache[template_type] = normalized_details
            selected = selector.get_selected()
            manual_name = None
            if self._manual_selection is not None and self._manual_selection[0] == template_type:
                manual_name = self._manual_selection[1]
            selector.add_templates(ordered_names, template_details=normalized_details)
            if manual_name and selector.has_template(manual_name):
                selector.select_template(manual_name, selection_source="restore")
            elif selected and selected in ordered_names:
                selector.select_template(selected, selection_source="restore")
            elif ordered_names:
                self._activate_default_template(template_type, selection_source="restore")
        if self._manual_selection is not None and self._manual_selection[0] == template_type:
            self._restore_manual_selection()

    def load_all_templates(
        self,
        data: dict[str, list[str]],
        *,
        details: dict[str, dict[str, TemplateItemDetails]] | None = None,
    ) -> None:
        for tt in self._selectors:
            names = data.get(tt, [])
            tt_details = (details or {}).get(tt, {})
            self.load_templates(tt, names, details=tt_details)

    def refresh_from_runtime_registry(self) -> None:
        """Refresh both tabs from the managed runtime catalog."""

        try:
            from docwen_runtime.templates import TemplateManager

            manager = TemplateManager.default()
            data: dict[str, list[str]] = {"docx": [], "xlsx": []}
            details: dict[str, dict[str, TemplateItemDetails]] = {"docx": {}, "xlsx": {}}
            for target in ("docx", "xlsx"):
                for info in manager.list_templates(target, include_disabled=False):
                    data[target].append(info.name)
                    custom = manager.is_custom(info)
                    details[target][info.name] = TemplateItemDetails(
                        resource_id=info.id,
                        usage_hint=info.description,
                        source_label=(
                            t("settings.templates.custom", "自定义")
                            if custom
                            else t("settings.templates.builtin", "内置")
                        ),
                        source_path=str(info.path) if custom else None,
                        updated_label=TemplateSelector._format_modified_ns(info.modified_ns),
                    )
            self.load_all_templates(data, details=details)
        except Exception as exc:
            logger.exception("Unable to refresh managed template catalog")
            self.show_load_error(t("components.template_selector.unavailable"), str(exc))

    def show_load_error(self, summary: str, detail: str) -> None:
        self._template_cache.clear()
        self._template_details_cache.clear()
        self._manual_selection = None
        for selector in self._selectors.values():
            selector.show_load_error(summary, detail)
        self._refresh_accessibility()

    def get_selector(self, template_type: str) -> TemplateSelector | None:
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

    def _forward_template_error(self, summary: str, detail: str) -> None:
        logger.debug("TemplateSelector error: %s — %s", summary, detail)

    def _refresh_accessibility(self) -> None:
        current_label = self._tab_titles.get(self._current_tab, "")
        selector = self._selectors.get(self._current_tab)
        selection_description = selector.get_selected() if selector is not None else ""
        description = selection_description or t("components.template_selector.empty_hint")
        self.setAccessibleName(current_label)
        self.setAccessibleDescription(description)
        self._pivot.setAccessibleName(current_label)
