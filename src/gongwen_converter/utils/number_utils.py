"""
数字处理工具模块
包含数字转换和格式化工具函数
"""

import logging

# 配置日志
logger = logging.getLogger(__name__)

def number_to_chinese(num: int) -> str:
    """数字转中文序号"""
    logger.debug(f"转换数字到中文: {num}")
    
    chinese_nums = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十']
    
    if num <= 0:
        logger.warning(f"无效数字: {num}, 返回空字符串")
        return ""
    elif num <= 10:
        result = chinese_nums[num - 1]
    elif num < 20:
        result = '十' + (chinese_nums[num - 11] if num > 10 else '')
    else:
        # 简化处理，实际公文通常不超过20级
        result = str(num)
    
    logger.debug(f"转换结果: {num} -> {result}")
    return result

def number_to_circled(num: int) -> str:
    """数字转带圈数字序号 (支持1-50)"""
    logger.debug(f"转换数字到带圈: {num}")
    
    circled_numbers = [
        '①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩',
        '⑪', '⑫', '⑬', '⑭', '⑮', '⑯', '⑰', '⑱', '⑲', '⑳',
        '㉑', '㉒', '㉓', '㉔', '㉕', '㉖', '㉗', '㉘', '㉙', '㉚',
        '㉛', '㉜', '㉝', '㉞', '㉟', '㊱', '㊲', '㊳', '㊴', '㊵',
        '㊶', '㊷', '㊸', '㊹', '㊺', '㊻', '㊼', '㊽', '㊾', '㊿'
    ]
    
    if 1 <= num <= 50:
        result = circled_numbers[num - 1]
    else:
        # 超出范围返回普通数字
        result = f"--{num}--"
        logger.warning(f"超出范围的数字: {num}, 使用普通格式")
    
    logger.debug(f"转换结果: {num} -> {result}")
    return result