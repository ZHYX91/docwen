"""
数字处理工具模块
包含数字转换和格式化工具函数，支持多种数字样式
"""

import logging

# 配置日志
logger = logging.getLogger(__name__)


def number_to_chinese(num: int) -> str:
    """
    转换为小写中文数字

    参数:
        num: 要转换的数字 (1-99)

    返回:
        str: 中文数字字符串，如 "一"、"十二"、"二十三"
    """
    logger.debug(f"转换数字到小写中文: {num}")

    if num <= 0:
        logger.warning(f"无效数字: {num}, 返回空字符串")
        return ""

    if num <= 10:
        chinese_nums = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
        result = chinese_nums[num - 1]
    elif num < 20:
        # 11-19: 十一、十二...
        result = "十" + ["一", "二", "三", "四", "五", "六", "七", "八", "九"][num - 11]
    elif num < 100:
        # 20-99: 二十、二十一...
        tens = ["二", "三", "四", "五", "六", "七", "八", "九"][num // 10 - 2]
        ones = num % 10
        if ones == 0:
            result = tens + "十"
        else:
            result = tens + "十" + ["一", "二", "三", "四", "五", "六", "七", "八", "九"][ones - 1]
    else:
        # 超出范围，返回阿拉伯数字
        result = str(num)
        logger.warning(f"数字 {num} 超出中文表示范围，使用阿拉伯数字")

    logger.debug(f"转换结果: {num} -> {result}")
    return result


def number_to_chinese_upper(num: int) -> str:
    """
    转换为大写中文数字（财务用）

    参数:
        num: 要转换的数字 (1-99)

    返回:
        str: 大写中文数字字符串，如 "壹"、"拾贰"、"贰拾叁"
    """
    logger.debug(f"转换数字到大写中文: {num}")

    if num <= 0:
        logger.warning(f"无效数字: {num}, 返回空字符串")
        return ""

    if num <= 10:
        upper_nums = ["壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖", "拾"]
        result = upper_nums[num - 1]
    elif num < 20:
        # 11-19: 拾壹、拾贰...
        result = "拾" + ["壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"][num - 11]
    elif num < 100:
        # 20-99: 贰拾、贰拾壹...
        tens = ["贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"][num // 10 - 2]
        ones = num % 10
        if ones == 0:
            result = tens + "拾"
        else:
            result = tens + "拾" + ["壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"][ones - 1]
    else:
        # 超出范围，返回阿拉伯数字
        result = str(num)
        logger.warning(f"数字 {num} 超出大写中文表示范围，使用阿拉伯数字")

    logger.debug(f"转换结果: {num} -> {result}")
    return result


def number_to_circled(num: int) -> str:
    """
    转换为带圈数字

    参数:
        num: 要转换的数字 (1-50)

    返回:
        str: 带圈数字字符串，如 "①"、"⑳"、"㊿"
    """
    logger.debug(f"转换数字到带圈: {num}")

    circled_numbers = [
        "①",
        "②",
        "③",
        "④",
        "⑤",
        "⑥",
        "⑦",
        "⑧",
        "⑨",
        "⑩",
        "⑪",
        "⑫",
        "⑬",
        "⑭",
        "⑮",
        "⑯",
        "⑰",
        "⑱",
        "⑲",
        "⑳",
        "㉑",
        "㉒",
        "㉓",
        "㉔",
        "㉕",
        "㉖",
        "㉗",
        "㉘",
        "㉙",
        "㉚",
        "㉛",
        "㉜",
        "㉝",
        "㉞",
        "㉟",
        "㊱",
        "㊲",
        "㊳",
        "㊴",
        "㊵",
        "㊶",
        "㊷",
        "㊸",
        "㊹",
        "㊺",
        "㊻",
        "㊼",
        "㊽",
        "㊾",
        "㊿",
    ]

    if 1 <= num <= 50:
        result = circled_numbers[num - 1]
    else:
        # 超出范围返回普通数字
        result = f"({num})"
        logger.warning(f"数字 {num} 超出带圈数字范围(1-50)，使用括号格式")

    logger.debug(f"转换结果: {num} -> {result}")
    return result


def number_to_arabic_full(num: int) -> str:
    """
    转换为全角阿拉伯数字

    参数:
        num: 要转换的数字

    返回:
        str: 全角数字字符串，如 "１"、"１０"、"９９"
    """
    logger.debug(f"转换数字到全角: {num}")

    # 全角数字映射表
    full_width_map = {
        "0": "０",
        "1": "１",
        "2": "２",
        "3": "３",
        "4": "４",
        "5": "５",
        "6": "６",
        "7": "７",
        "8": "８",
        "9": "９",
    }

    # 将每个数字字符转换为全角
    result = "".join(full_width_map.get(c, c) for c in str(num))
    logger.debug(f"转换结果: {num} -> {result}")
    return result


def number_to_letter_upper(num: int) -> str:
    """
    转换为大写拉丁字母

    参数:
        num: 要转换的数字 (1-26对应A-Z, 27=AA, 28=AB...)

    返回:
        str: 大写字母字符串，如 "A"、"Z"、"AA"
    """
    logger.debug(f"转换数字到大写字母: {num}")

    if num <= 0:
        logger.warning(f"无效数字: {num}, 返回空字符串")
        return ""

    original_num = num  # 保存原始值用于日志
    result = ""
    while num > 0:
        num -= 1  # 调整为0-based索引
        result = chr(65 + (num % 26)) + result
        num //= 26

    logger.debug(f"转换结果: {original_num} -> {result}")
    return result


def number_to_letter_lower(num: int) -> str:
    """
    转换为小写拉丁字母

    参数:
        num: 要转换的数字 (1-26对应a-z, 27=aa, 28=ab...)

    返回:
        str: 小写字母字符串，如 "a"、"z"、"aa"
    """
    logger.debug(f"转换数字到小写字母: {num}")

    if num <= 0:
        logger.warning(f"无效数字: {num}, 返回空字符串")
        return ""

    original_num = num  # 保存原始值用于日志
    result = ""
    while num > 0:
        num -= 1  # 调整为0-based索引
        result = chr(97 + (num % 26)) + result
        num //= 26

    logger.debug(f"转换结果: {original_num} -> {result}")
    return result


def number_to_roman_upper(num: int) -> str:
    """
    转换为大写罗马数字

    参数:
        num: 要转换的数字 (1-3999)

    返回:
        str: 大写罗马数字字符串，如 "I"、"IV"、"IX"、"MCMXCIV"
    """
    logger.debug(f"转换数字到大写罗马: {num}")

    if num <= 0 or num >= 4000:
        logger.warning(f"数字 {num} 超出罗马数字范围(1-3999)，使用阿拉伯数字")
        return str(num)

    original_num = num  # 保存原始值用于日志

    # 罗马数字映射表（从大到小）
    values = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    symbols = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]

    result = ""
    for i in range(len(values)):
        count = num // values[i]
        if count:
            result += symbols[i] * count
            num -= values[i] * count

    logger.debug(f"转换结果: {original_num} -> {result}")
    return result


def number_to_roman_lower(num: int) -> str:
    """
    转换为小写罗马数字

    参数:
        num: 要转换的数字 (1-3999)

    返回:
        str: 小写罗马数字字符串，如 "i"、"iv"、"ix"、"mcmxciv"
    """
    logger.debug(f"转换数字到小写罗马: {num}")

    # 复用大写转换，然后转小写
    result = number_to_roman_upper(num).lower()
    logger.debug(f"转换结果: {num} -> {result}")
    return result
