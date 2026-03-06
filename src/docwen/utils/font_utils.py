"""
字体工具模块
用于处理字体相关的功能
提供获取系统字体和根据配置选择字体的功能
"""

import logging
import tkinter as tk
from tkinter import font as tkfont

# 配置日志
logger = logging.getLogger(__name__)

# 字体配置
DEFAULT_FONT_FAMILIES = ["Microsoft YaHei", "SimSun", "Arial", "Helvetica", "sans-serif"]
DEFAULT_FONT_SIZE = 12

TITLE_FONT_FAMILIES = ["Microsoft YaHei", "SimHei", "Arial", "Helvetica", "sans-serif"]
TITLE_FONT_SIZE = 16

SMALL_FONT_FAMILIES = ["SimSun", "Arial", "Helvetica", "sans-serif"]
SMALL_FONT_SIZE = 10

MICRO_FONT_FAMILIES = ["SimSun", "Arial", "Helvetica", "sans-serif"]
MICRO_FONT_SIZE = 8

LARGE_FONT_FAMILIES = ["Microsoft YaHei", "SimHei", "Arial Black", "Impact", "sans-serif"]
LARGE_FONT_SIZE = 14

# 缓存系统字体列表，避免重复获取
_system_fonts_cache = None


def get_system_fonts() -> list[str]:
    """
    获取系统已安装的字体列表

    返回:
        List[str]: 系统字体名称列表
    """
    global _system_fonts_cache

    # 如果已经缓存，直接返回缓存结果
    if _system_fonts_cache is not None:
        return _system_fonts_cache

    try:
        # 检查是否已经有活动的Tkinter实例
        if getattr(tk, "_default_root", None) is not None:
            # 使用现有的根窗口获取字体
            fonts = list(tkfont.families())
            logger.debug(f"使用现有根窗口获取系统字体列表，共 {len(fonts)} 种字体")
        else:
            # 没有活动的Tkinter实例，创建临时根窗口
            root = tk.Tk()
            root.withdraw()

            # 获取系统字体列表
            fonts = list(tkfont.families())
            root.destroy()

            logger.debug(f"创建临时根窗口获取系统字体列表，共 {len(fonts)} 种字体")

        # 缓存结果
        _system_fonts_cache = fonts
        return fonts
    except Exception as e:
        logger.error(f"获取系统字体失败: {e!s}")
        # 返回一个默认字体列表作为回退
        fallback_fonts = [
            "Arial",
            "Helvetica",
            "Times New Roman",
            "Courier New",
            "Microsoft Sans Serif",
            "Tahoma",
            "Verdana",
            "SimSun",
            "Microsoft YaHei",
            "SimHei",
            "KaiTi",
            "FangSong",
        ]
        _system_fonts_cache = fallback_fonts
        return fallback_fonts


def get_available_font(font_families: list[str]) -> str | None:
    """
    根据字体家族列表选择第一个可用的字体

    参数:
        font_families: 字体家族列表，按优先级排序

    返回:
        Optional[str]: 第一个可用的字体名称，如果没有可用字体则返回None
    """
    if not font_families:
        logger.warning("字体家族列表为空")
        return None

    system_fonts = get_system_fonts()

    # 创建字体名称映射表，处理常见的字体别名和变体
    font_aliases = {
        "Microsoft YaHei": ["Microsoft YaHei", "微软雅黑", "YaHei"],
        "PingFang SC": ["PingFang SC", "苹方", "PingFang"],
        "Noto Sans CJK SC": ["Noto Sans CJK SC", "思源黑体", "Source Han Sans", "Noto Sans"],
        "SimSun": ["SimSun", "宋体", "Sim Sun"],
        "SimHei": ["SimHei", "黑体", "Sim Hei"],
        "KaiTi": ["KaiTi", "楷体", "Kai Ti"],
        "FangSong": ["FangSong", "仿宋", "Fang Song"],
        "Arial": ["Arial", "Arial MT"],
        "Helvetica": ["Helvetica", "Helvetica Neue"],
        "sans-serif": ["sans-serif", "Sans Serif"],
    }

    for font in font_families:
        # 首先尝试精确匹配
        if font in system_fonts:
            logger.debug(f"找到精确匹配的字体: {font}")
            return font

        # 如果没有精确匹配，尝试通过别名查找
        if font in font_aliases:
            for alias in font_aliases[font]:
                if alias in system_fonts:
                    logger.debug(f"通过别名找到字体: {alias} (原名称: {font})")
                    return alias

        # 最后尝试模糊匹配（不区分大小写）
        font_lower = font.lower()
        for system_font in system_fonts:
            if font_lower in system_font.lower():
                logger.debug(f"通过模糊匹配找到字体: {system_font} (搜索: {font})")
                return system_font

    logger.warning(f"没有找到可用的字体，配置的字体家族: {font_families}")
    return None


def get_default_font() -> tuple[str, int]:
    """
    获取默认字体和大小

    返回:
        Tuple[str, int]: (字体名称, 字体大小)
    """
    # 尝试获取可用字体
    available_font = get_available_font(DEFAULT_FONT_FAMILIES)
    if available_font:
        return available_font, DEFAULT_FONT_SIZE

    # 如果配置的字体都不可用，尝试使用一些常见字体
    fallback_fonts = ["Microsoft YaHei", "SimSun", "Arial", "Helvetica", "sans-serif"]
    available_font = get_available_font(fallback_fonts)
    if available_font:
        return available_font, DEFAULT_FONT_SIZE

    # 如果还是找不到，返回Tk默认字体
    return "TkDefaultFont", DEFAULT_FONT_SIZE


def get_title_font() -> tuple[str, int]:
    """
    获取标题字体和大小

    返回:
        Tuple[str, int]: (字体名称, 字体大小)
    """
    # 尝试获取可用字体
    available_font = get_available_font(TITLE_FONT_FAMILIES)
    if available_font:
        return available_font, TITLE_FONT_SIZE

    # 如果配置的字体都不可用，尝试使用一些常见字体
    fallback_fonts = ["Microsoft YaHei", "SimHei", "Arial", "Helvetica", "sans-serif"]
    available_font = get_available_font(fallback_fonts)
    if available_font:
        return available_font, TITLE_FONT_SIZE

    # 如果还是找不到，返回Tk默认字体
    return "TkHeadingFont", TITLE_FONT_SIZE


def get_small_font() -> tuple[str, int]:
    """
    获取小字体和大小

    返回:
        Tuple[str, int]: (字体名称, 字体大小)
    """
    # 尝试获取可用字体
    available_font = get_available_font(SMALL_FONT_FAMILIES)
    if available_font:
        return available_font, SMALL_FONT_SIZE

    # 如果配置的字体都不可用，尝试使用一些常见字体
    fallback_fonts = ["SimSun", "Arial", "Helvetica", "sans-serif"]
    available_font = get_available_font(fallback_fonts)
    if available_font:
        return available_font, SMALL_FONT_SIZE

    # 如果还是找不到，返回Tk默认字体
    return "TkSmallFont", SMALL_FONT_SIZE


def get_micro_font() -> tuple[str, int]:
    """
    获取微字体和大小

    返回:
        Tuple[str, int]: (字体名称, 字体大小)
    """
    # 尝试获取可用字体
    available_font = get_available_font(MICRO_FONT_FAMILIES)
    if available_font:
        return available_font, MICRO_FONT_SIZE

    # 如果配置的字体都不可用，尝试使用一些常见字体
    fallback_fonts = ["SimSun", "Arial", "Helvetica", "sans-serif"]
    available_font = get_available_font(fallback_fonts)
    if available_font:
        return available_font, MICRO_FONT_SIZE

    # 如果还是找不到，返回Tk默认字体
    return "TkSmallFont", MICRO_FONT_SIZE


def get_large_font() -> tuple[str, int]:
    """
    获取大字体和大小

    返回:
        Tuple[str, int]: (字体名称, 字体大小)
    """
    # 尝试获取可用字体
    available_font = get_available_font(LARGE_FONT_FAMILIES)
    if available_font:
        return available_font, LARGE_FONT_SIZE

    # 如果配置的字体都不可用，尝试使用一些常见字体
    fallback_fonts = ["Microsoft YaHei", "SimHei", "Arial Black", "Impact", "sans-serif"]
    available_font = get_available_font(fallback_fonts)
    if available_font:
        return available_font, LARGE_FONT_SIZE

    # 如果还是找不到，返回Tk默认字体
    return "TkHeadingFont", LARGE_FONT_SIZE


def apply_font_styles(root: tk.Tk):
    """
    应用字体样式到根窗口

    参数:
        root: Tkinter根窗口
    """
    try:
        # 获取默认字体配置
        font_name, font_size = get_default_font()
        title_font_name, title_font_size = get_title_font()
        small_font_name, small_font_size = get_small_font()
        large_font_name, large_font_size = get_large_font()

        # 创建样式对象
        from tkinter import ttk

        style = ttk.Style()

        # 设置默认字体
        style.configure(".", font=(font_name, font_size))

        # 设置标题字体
        style.configure("Title.TLabel", font=(title_font_name, title_font_size))
        style.configure("Title.TButton", font=(title_font_name, title_font_size))

        # 设置小字体
        style.configure("Small.TLabel", font=(small_font_name, small_font_size))
        style.configure("Small.TButton", font=(small_font_name, small_font_size))

        # 设置大字体
        style.configure("Large.TLabel", font=(large_font_name, large_font_size))
        style.configure("Large.TButton", font=(large_font_name, large_font_size))

        logger.info(
            f"已应用字体样式: 默认={font_name} {font_size}pt, 标题={title_font_name} {title_font_size}pt, 小字体={small_font_name} {small_font_size}pt, 大字体={large_font_name} {large_font_size}pt"
        )

    except Exception as e:
        logger.error(f"应用字体样式失败: {e!s}")
