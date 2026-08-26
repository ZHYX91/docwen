"""Qt 主题语义与样式辅助函数。

本模块不依赖 ``theme_manager`` --- 需要当前主题名的函数通过参数接收，
从而避免循环导入。
"""

from __future__ import annotations

from typing import Literal

from PySide6.QtWidgets import QWidget

ThemeClass = Literal["primary", "success", "warning", "danger", "info", "secondary"]

DEFAULT_THEME = "light"

# 语义颜色 token
COLOR_PRIMARY = "#2F6FEB"
COLOR_SUCCESS = "#2DA44E"
COLOR_DANGER = "#D73A49"
COLOR_WARNING = "#BF8700"
COLOR_INFO = "#1F6FEB"
COLOR_SECONDARY = "#6E7781"

# 强调色交互态 token：必须与 theme_manager 中 QPalette Highlight (#3B82F6) 同源。
# QSS 不允许对 palette(highlight) 做明暗运算，hover/pressed 态在此显式定义。
COLOR_ACCENT = "#3B82F6"
COLOR_ACCENT_HOVER = "#2563EB"
COLOR_ACCENT_PRESSED = "#1D4ED8"

_BUTTON_CLASS_COLORS: dict[str, str] = {
    "primary": COLOR_PRIMARY,
    "success": COLOR_SUCCESS,
    "warning": COLOR_WARNING,
    "danger": COLOR_DANGER,
    "info": COLOR_INFO,
    "secondary": COLOR_SECONDARY,
}

_LIGHT_THEME_CLASS_COLORS: dict[str, str] = dict(_BUTTON_CLASS_COLORS)
_DARK_THEME_CLASS_COLORS: dict[str, str] = {
    "primary": "#58A6FF",
    "success": "#3FB950",
    "warning": "#D29922",
    "danger": "#F85149",
    "info": "#58A6FF",
    "secondary": "#8B949E",
}

_STATUS_THEME_CLASS: dict[str, ThemeClass] = {
    "pending": "secondary",
    "processing": "info",
    "completed": "success",
    "skipped": "warning",
    "failed": "danger",
    "cancelled": "warning",
}


def is_dark_theme(theme_name: str | None) -> bool:
    """判断给定主题名是否为暗色主题。"""
    return "dark" in str(theme_name or DEFAULT_THEME).lower()


def _resolve_theme_name(theme_name: str | None) -> str:
    """解析主题名；缺省时回退到默认主题。

    注意：本函数不再从 ``theme_manager`` 回退当前主题，以消除循环依赖。
    调用方应自行获取有效主题名后传入。
    """
    return theme_name if theme_name else DEFAULT_THEME


def get_theme_class_color(theme_class: ThemeClass | str, theme_name: str | None = None) -> str:
    """根据指定主题返回语义类的十六进制颜色。"""
    palette = _DARK_THEME_CLASS_COLORS if is_dark_theme(_resolve_theme_name(theme_name)) else _LIGHT_THEME_CLASS_COLORS
    return palette.get(str(theme_class), palette["secondary"])


def get_status_theme_class(status: str) -> ThemeClass:
    """根据状态返回统一的语义主题类。"""
    return _STATUS_THEME_CLASS.get(status, "secondary")


def get_status_color(status: str, theme_name: str | None = None) -> str:
    """根据状态和指定主题返回十六进制颜色。"""
    return get_theme_class_color(get_status_theme_class(status), theme_name)


def apply_theme_class(widget: QWidget | None, theme_class: ThemeClass | str) -> None:
    """统一处理语义类名设置与样式刷新。"""
    if widget is None:
        return
    widget.setProperty("class", theme_class)
    style = widget.style()
    if style is not None:
        style.unpolish(widget)
        style.polish(widget)
