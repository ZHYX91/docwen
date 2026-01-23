"""
核心验证工具
包含基本数据验证功能
"""

import re
import logging

# 配置日志
logger = logging.getLogger(__name__)

def is_value_empty(value) -> bool:
    """统一空值检查函数，支持多种数据类型"""
    # None值直接返回True
    if value is None:
        return True
        
    # 字符串类型检查
    if isinstance(value, str):
        return value.strip() == ''
    
    # 列表/元组/集合等可迭代类型
    if isinstance(value, (list, tuple, set)):
        # 空容器检查
        if len(value) == 0:
            return True
            
        # 容器内元素是否全为空
        for item in value:
            if not is_value_empty(item):
                return False
        return True
        
    # 字典类型检查
    if isinstance(value, dict):
        # 空字典检查
        if len(value) == 0:
            return True
            
        # 字典值是否全为空
        for v in value.values():
            if not is_value_empty(v):
                return False
        return True
        
    # 其他类型默认非空
    return False

def is_chinese(text: str) -> bool:
    """检查字符是否为中文字符"""
    if not text:
        return False
    # 对于多字符输入只检查第一个字符
    return '\u4e00' <= text[0] <= '\u9fff'

def contains_chinese(text: str) -> bool:
    """检查文本是否包含中文字符"""
    return any(is_chinese(c) for c in text)

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
        r'^\d{4}年\d{1,2}月\d{1,2}日$',  # 2023年12月31日
        r'^\d{4}-\d{2}-\d{2}$',          # 2023-12-31
        r'^\d{4}/\d{2}/\d{2}$',           # 2023/12/31
        r'^\d{4}\.\d{2}\.\d{2}$'         # 2023.12.31
    ]
    
    for pattern in valid_patterns:
        if re.match(pattern, date_str):
            logger.info(f"日期格式验证通过: {date_str}")
            return True
            
    logger.warning(f"无效的日期格式: {date_str}")
    return False

# 模块测试
if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    logger.info("验证工具模块测试")
    
    # 空值测试
    assert is_value_empty("") is True
    assert is_value_empty("  ") is True
    assert is_value_empty([]) is True
    assert is_value_empty(["", None]) is True
    assert is_value_empty("内容") is False
    
    # 中文字符测试
    assert is_chinese("文") is True
    assert is_chinese("A") is False
    
    # 日期验证测试
    assert validate_date_format("2023年12月31日") is True
    assert validate_date_format("2023-12-31") is True
    assert validate_date_format("31/12/2023") is False
    
    logger.info("所有测试通过!")