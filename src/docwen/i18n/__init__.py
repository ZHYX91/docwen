"""
国际化（i18n）模块

提供多语言支持功能，包括：
- I18nManager: 语言管理器单例类
- StyleNameResolver: 多语言样式名解析器
- t(): 翻译函数快捷方式
- get_available_locales(): 获取可用语言列表
- t_locale(): 获取指定语言的翻译
- t_all_locales(): 获取所有语言版本的翻译
"""

from typing import TYPE_CHECKING

from .i18n_manager import I18nManager
from .style_resolver import StyleNameResolver, style_resolver

if TYPE_CHECKING:
    _i18n: I18nManager | None

_i18n = None


def _get_i18n() -> I18nManager:
    global _i18n
    if _i18n is None:
        _i18n = I18nManager()
    return _i18n


def t(key: str, default: str | None = None, **kwargs) -> str:
    return _get_i18n().t(key, default=default, **kwargs)


def t_locale(key: str, locale: str, **kwargs) -> str:
    return _get_i18n().t_locale(key, locale, **kwargs)


def t_all_locales(key: str) -> dict[str, str]:
    return _get_i18n().t_all_locales(key)


def get_available_locales() -> list[dict[str, str]]:
    return _get_i18n().get_available_locales()


def get_current_locale() -> str:
    return _get_i18n().get_current_locale()


def get_detection_priority() -> list[str]:
    return _get_i18n().get_detection_priority()


def set_locale(locale: str) -> bool:
    return _get_i18n().set_locale(locale)


def reload_translations() -> None:
    _get_i18n().reload_translations()


__all__ = [
    # 类
    "I18nManager",
    "StyleNameResolver",
    # 语言管理
    "get_available_locales",
    "get_current_locale",
    "get_detection_priority",
    "reload_translations",
    "set_locale",
    "style_resolver",
    # 翻译函数
    "t",
    "t_all_locales",
    "t_locale",
]
