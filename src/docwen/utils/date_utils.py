"""
日期时间工具模块
包含日期格式转换等工具函数
"""

import datetime
import logging
import re

# 配置日志
logger = logging.getLogger(__name__)


def convert_date_format(date_str: str) -> str:
    """将日期转换为YYYY-MM-DD格式"""
    logger.debug(f"转换日期格式: {date_str}")

    # 尝试匹配常见日期格式
    patterns = [
        r"(\d{4})年(\d{1,2})月(\d{1,2})日",
        r"(\d{4})年(\d{1,2})月(\d{1,2})号",
        r"(\d{4})-(\d{1,2})-(\d{1,2})",
        r"(\d{4})\.(\d{1,2})\.(\d{1,2})",
    ]

    for pattern in patterns:
        match = re.match(pattern, date_str)
        if match:
            year = match.group(1)
            month = match.group(2).zfill(2)
            day = match.group(3).zfill(2)
            result = f"{year}-{month}-{day}"
            logger.debug(f"转换成功: {date_str} -> {result}")
            return result

    # 无法转换，返回原字符串
    logger.warning(f"无法识别的日期格式: {date_str}")
    return date_str


def generate_timestamp() -> str:
    """生成基于日期时间的版本号"""
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    logger.debug(f"生成时间戳: {timestamp}")
    return timestamp
