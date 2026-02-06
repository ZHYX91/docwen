"""
CLI国际化辅助模块

为CLI提供独立的国际化支持：
- 支持 --lang 参数临时覆盖语言
- 封装 i18n_manager 调用
- 提供CLI专用翻译函数

使用方式：
    from docwen.cli.i18n import cli_t, init_cli_locale
    
    # 初始化（在main入口调用）
    init_cli_locale(args.lang)
    
    # 使用翻译
    print(cli_t("cli.messages.error_no_files"))
"""

import logging
from typing import Optional

logger = logging.getLogger(__name__)

# CLI临时语言设置（不保存到配置）
_cli_locale: Optional[str] = None


def init_cli_locale(lang: Optional[str] = None):
    """
    初始化CLI语言设置
    
    如果提供了 --lang 参数，临时覆盖语言设置。
    这不会修改配置文件，只影响当前CLI会话。
    
    Args:
        lang: 语言代码（如 zh_CN 或 en_US），None表示使用配置文件设置
    """
    global _cli_locale
    
    if lang:
        # 验证语言代码
        try:
            from docwen.i18n.i18n_manager import I18nManager
            valid_locales = [l.get("code") for l in I18nManager().get_available_locales() if l.get("code")]
        except Exception:
            valid_locales = ['zh_CN', 'en_US']

        if lang in valid_locales:
            _cli_locale = lang
            logger.debug(f"CLI临时语言设置: {lang}")
            
            # 重新加载I18nManager的翻译
            try:
                i18n = I18nManager()
                # 临时修改内部语言（不保存配置）
                i18n._locale = lang
                i18n._load_translations()
                logger.debug(f"已加载语言: {lang}")
            except Exception as e:
                logger.warning(f"切换语言失败: {e}")
        else:
            logger.warning(f"无效的语言代码: {lang}，将使用默认设置（可用：{', '.join(valid_locales)}）")
    else:
        _cli_locale = None


def get_cli_locale() -> str:
    """
    获取当前CLI语言设置
    
    Returns:
        str: 语言代码
    """
    if _cli_locale:
        return _cli_locale
    
    try:
        from docwen.i18n.i18n_manager import I18nManager
        return I18nManager().get_current_locale()
    except Exception:
        return 'zh_CN'


def cli_t(key: str, default: Optional[str] = None, **kwargs) -> str:
    """
    CLI翻译函数
    
    根据键名查找翻译文本，支持嵌套键和参数替换。
    优先使用 --lang 指定的语言，否则使用配置文件语言。
    
    Args:
        key: 翻译键，支持点号分隔的嵌套键（如 "cli.messages.error"）
        default: 默认值，当找不到翻译时返回
        **kwargs: 用于格式化字符串的参数
        
    Returns:
        str: 翻译后的文本
        
    示例：
        cli_t("cli.header")  # 返回 "DocWen - 交互模式"
        cli_t("cli.messages.files_found", count=5)  # 返回 "找到 5 个文件"
    """
    try:
        from docwen.i18n.i18n_manager import I18nManager
        i18n = I18nManager()
        return i18n.t(key, default=default, **kwargs)
    except Exception as e:
        logger.debug(f"翻译失败 [{key}]: {e}")
        if default is not None:
            return default
        return f"[{key}]"


# 便捷别名
t = cli_t
