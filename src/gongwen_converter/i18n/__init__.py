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

from .i18n_manager import I18nManager
from .style_resolver import StyleNameResolver, style_resolver

# 创建全局实例
_i18n = I18nManager()

# 导出翻译函数
t = _i18n.t
t_locale = _i18n.t_locale
t_all_locales = _i18n.t_all_locales
get_available_locales = _i18n.get_available_locales
get_current_locale = _i18n.get_current_locale
get_detection_priority = _i18n.get_detection_priority
set_locale = _i18n.set_locale
reload_translations = _i18n.reload_translations

__all__ = [
    # 类
    'I18nManager',
    'StyleNameResolver',
    'style_resolver',
    # 翻译函数
    't',
    't_locale',
    't_all_locales',
    # 语言管理
    'get_available_locales',
    'get_current_locale',
    'get_detection_priority',
    'set_locale',
    'reload_translations',
]
