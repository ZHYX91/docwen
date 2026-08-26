"""模板选择器样式。

为 ``TemplateSelector`` 和 ``TabbedTemplateSelector`` 提供
QSS 样式表，统一控制列表、底部信息栏、空状态和选项卡的视觉风格。
"""

from __future__ import annotations

from .design_tokens import Border, Radius, Spacing


def build_template_selector_stylesheet() -> str:
    """返回模板选择器及其关联组件的完整样式表。"""
    return "\n".join(
        [
            "/* docwen-template-selector-foundation */",
            "QWidget#templateSelectorRoot {",
            "    background: transparent;",
            "}",
            "QListWidget#templateSelectorList {",
            f"    border: {Border.THIN}px solid palette(midlight);",
            f"    border-radius: {Radius.MEDIUM}px;",
            f"    padding: {Spacing.XS}px;",
            "    background: palette(base);",
            "}",
            "QListWidget#templateSelectorList:focus {",
            f"    border: {Border.THIN}px solid palette(highlight);",
            "}",
            "QListWidget#templateSelectorList::item {",
            f"    padding: {Spacing.XS}px 0;",
            "}",
            "QWidget#templateSelectorFooterRow {",
            f"    margin-top: {Spacing.XS}px;",
            f"    border: {Border.THIN}px solid palette(midlight);",
            f"    border-radius: {Radius.MEDIUM}px;",
            "    background: palette(alternate-base);",
            "}",
            "QLabel#templateSelectorMetaLabel {",
            "    color: palette(mid);",
            "    background: transparent;",
            "    padding: 2px 0;",
            "}",
            "QToolButton#templateSelectorOpenButton {",
            "    padding: 0;",
            "}",
            "QFrame#templateSelectorEmptyState {",
            f"    border: {Border.THIN}px dashed palette(midlight);",
            f"    border-radius: {Radius.MEDIUM}px;",
            f"    padding: {Spacing.SM}px;",
            "    background: palette(alternate-base);",
            "}",
            "QLabel#templateSelectorEmptyTitle {",
            "    font-weight: 600;",
            "}",
            "QLabel#templateSelectorEmptyHint {",
            "    color: palette(mid);",
            "}",
            "/* Tabbed selector */",
            "QWidget#tabbedTemplateSelector {",
            "    background: transparent;",
            "}",
            "QWidget#templateSelectorStack {",
            "    background: transparent;",
            "}",
        ]
    )
