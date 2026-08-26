"""主窗口壳层样式。"""

from __future__ import annotations


def build_main_window_stylesheet() -> str:
    """主窗口壳层相关的局部样式统一收口。"""
    return "\n".join(
        [
            "/* docwen-main-window-foundation */",
            "QWidget#bottomBar {",
            "    background: transparent;",
            "}",
            "QToolButton#fontSizeButton,",
            "QToolButton#aboutButton,",
            "QToolButton#settingsButton {",
            "    background: transparent;",
            "    border: 1px solid transparent;",
            "    border-radius: 4px;",
            "    padding: 0px;",
            "}",
            "QToolButton#fontSizeButton:hover,",
            "QToolButton#aboutButton:hover,",
            "QToolButton#settingsButton:hover {",
            "    background: rgba(127, 127, 127, 0.16);",
            "}",
            "QToolButton#fontSizeButton:focus,",
            "QToolButton#aboutButton:focus,",
            "QToolButton#settingsButton:focus {",
            "    border: 1px solid palette(highlight);",
            "}",
        ]
    )
