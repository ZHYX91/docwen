"""
核心验证工具
包含基本数据验证功能
"""

import logging
import re

# 配置日志
logger = logging.getLogger(__name__)


def is_value_empty(value) -> bool:
    """统一空值检查函数，支持多种数据类型"""
    # None值直接返回True
    if value is None:
        return True

    # 字符串类型检查
    if isinstance(value, str):
        return value.strip() == ""

    # 列表/元组/集合等可迭代类型
    if isinstance(value, (list, tuple, set)):
        # 空容器检查
        if len(value) == 0:
            return True

        # 容器内元素是否全为空
        return all(is_value_empty(item) for item in value)

    # 字典类型检查
    if isinstance(value, dict):
        # 空字典检查
        if len(value) == 0:
            return True

        # 字典值是否全为空
        return all(is_value_empty(v) for v in value.values())

    # 其他类型默认非空
    return False


def is_chinese(text: str) -> bool:
    """检查字符是否为中文字符"""
    if not text:
        return False
    # 对于多字符输入只检查第一个字符
    return "\u4e00" <= text[0] <= "\u9fff"


def contains_chinese(text: str) -> bool:
    """检查文本是否包含中文字符"""
    return any(is_chinese(c) for c in text)


def contains_east_asian(text: str) -> bool:
    """
    检查文本是否包含东亚语境字符

    用途：
        用于排版/字体决策（如设置 OpenXML 的 w:hint="eastAsia"），覆盖常见的
        中日韩文字与符号区间，而不仅限于汉字。

    参数:
        text: 待检测文本

    返回:
        bool: 包含东亚语境字符返回 True，否则返回 False
    """
    for ch in text:
        cp = ord(ch)
        if (
            0x3400 <= cp <= 0x4DBF  # CJK Unified Ideographs Extension A
            or 0x4E00 <= cp <= 0x9FFF  # CJK Unified Ideographs
            or 0x3040 <= cp <= 0x309F  # Hiragana
            or 0x30A0 <= cp <= 0x30FF  # Katakana
            or 0x31F0 <= cp <= 0x31FF  # Katakana Phonetic Extensions
            or 0xAC00 <= cp <= 0xD7AF  # Hangul Syllables
            or 0x1100 <= cp <= 0x11FF  # Hangul Jamo
            or 0x3130 <= cp <= 0x318F  # Hangul Compatibility Jamo
            or 0x3000 <= cp <= 0x303F  # CJK Symbols and Punctuation
            or 0xFF00 <= cp <= 0xFFEF  # Halfwidth and Fullwidth Forms
            or 0x3100 <= cp <= 0x312F  # Bopomofo
            or cp in {
                0x2018,  # ‘
                0x2019,  # ’
                0x201C,  # “
                0x201D,  # ”
                0x2026,  # …
                0x2013,  # –
                0x2014,  # —
                0x2015,  # ―
            }
        ):
            return True
    return False


def validate_date_format(date_str: str) -> bool:
    """
    验证日期格式是否符合公文规范
    支持格式: YYYY年MM月DD日, YYYY-MM-DD, YYYY/MM/DD

    参数:
        date_str: 日期字符串

    返回:
        bool: 格式正确返回True，否则返回False
    """
    logger.info(f"验证日期格式: {date_str}")

    # 定义合法日期格式正则
    valid_patterns = [
        r"^\d{4}年\d{1,2}月\d{1,2}日$",  # 2023年12月31日
        r"^\d{4}-\d{2}-\d{2}$",  # 2023-12-31
        r"^\d{4}/\d{2}/\d{2}$",  # 2023/12/31
        r"^\d{4}\.\d{2}\.\d{2}$",  # 2023.12.31
    ]

    for pattern in valid_patterns:
        if re.match(pattern, date_str):
            logger.info(f"日期格式验证通过: {date_str}")
            return True

    logger.warning(f"无效的日期格式: {date_str}")
    return False


def validate_ocr_requires_images(extract_images: bool, extract_ocr: bool) -> tuple[bool, str]:
    return True, ""
