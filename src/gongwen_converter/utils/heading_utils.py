"""
标题处理工具模块
包含标题级别检测和序号清理功能
支持全角数字和全角符号
"""

import re
import logging

# 配置日志
logger = logging.getLogger(__name__)

# 中文数字映射表
CHINESE_NUMS = '一二三四五六七八九十'

# 全角数字 (０-９)
FULL_WIDTH_NUMS = '０１２３４５６７８９'

def detect_heading_level(text: str) -> tuple:
    """
    检测小标题级别并清理序号
    自动识别1-5级小标题格式并清理小标题序号
    
    参数:
        text: 原始小标题文本
        
    返回:
        tuple: (清理后的文本, 小标题级别)
            小标题级别: 
                0 = 非小标题
                1 = 一级小标题
                2 = 二级小标题
                3 = 三级小标题
                4 = 四级小标题
                5 = 五级小标题
                
    支持的小标题格式:
    1. 一级小标题: 中文数字+顿号 (一、小标题内容)
    2. 二级小标题: 带括号中文数字 (（一）小标题内容)
    3. 三级小标题: 数字加点 (1. 小标题内容) - 支持半角和全角数字
    4. 四级小标题: 带括号数字 ((1) 小标题内容) - 支持半角和全角数字
    5. 五级小标题: 带圈数字 (① 小标题内容) - 支持半角和全角圈数字
    """
    # 记录输入和初始化
    logger.debug(f"检测标题级别 - 原始文本: '{text}'")
    cleaned_text = text.strip()
    original_text = cleaned_text  # 保存原始文本用于日志
    
    # 五级小标题: 带圈数字 (①, ②, ③) - 支持半角和全角
    if re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟㊱㊲㊳㊴㊵㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿]', cleaned_text):
        # 去除带圈数字序号 (兼容点号、顿号、逗号)
        cleaned = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟㊱㊲㊳㊴㊵㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿][\s\.、，,。]?', '', cleaned_text)
        logger.info(f"识别为五级小标题 | 原始: '{original_text}' -> 清理后: '{cleaned}'")
        return cleaned, 5
    
    # 四级小标题: 带括号数字 ((1), (2), (3)) - 支持半角和全角数字
    # 全角数字范围: ０１２３４５６７８９ (U+FF10-U+FF19)
    if re.match(r'^[（(][0-9\uFF10-\uFF19][)）]', cleaned_text) or re.match(r'^\(\d+\)', cleaned_text):
        # 处理中英文括号混用 (兼容点号、顿号、逗号)
        cleaned = re.sub(r'^[（(][0-9\uFF10-\uFF19]+[)）][\s\.、，,。]?', '', cleaned_text)
        logger.info(f"识别为四级小标题 | 原始: '{original_text}' -> 清理后: '{cleaned}'")
        return cleaned, 4
    
    # 三级小标题: 数字加点 (1., 2., 3.) - 支持半角和全角数字
    # 全角点号: ． (U+FF0E)
    if re.match(r'^[0-9\uFF10-\uFF19]+[\.．、，,。]', cleaned_text):
        # 处理点号/顿号混用 (支持全角点号)
        cleaned = re.sub(r'^[0-9\uFF10-\uFF19]+[\.．、，,。]\s*', '', cleaned_text)
        logger.info(f"识别为三级小标题 | 原始: '{original_text}' -> 清理后: '{cleaned}'")
        return cleaned, 3
    
    # 二级小标题: 带括号中文数字 ((一), (二))
    if re.match(r'^[（(][' + CHINESE_NUMS + '][)）]', cleaned_text):
        # 处理中英文括号混用 (兼容点号、顿号、逗号)
        cleaned = re.sub(r'^[（(][' + CHINESE_NUMS + r']+[)）][\s\.、，,。]?', '', cleaned_text)
        logger.info(f"识别为二级小标题 | 原始: '{original_text}' -> 清理后: '{cleaned}'")
        return cleaned, 2
    
    # 一级小标题: 中文数字 (一、, 二、)
    if re.match(r'^[' + CHINESE_NUMS + r']+[、,，\.。]', cleaned_text):
        # 处理顿号/逗号混用
        cleaned = re.sub(r'^[' + CHINESE_NUMS + r']+[、,，\.。]\s*', '', cleaned_text)
        logger.info(f"识别为一级小标题 | 原始: '{original_text}' -> 清理后: '{cleaned}'")
        return cleaned, 1
    
    # 非小标题内容
    logger.info(f"非小标题内容: '{original_text}'")
    return cleaned_text, 0

def split_content_by_delimiters(text):
    """
    按照特定符号分割文本内容
    查找顺序：中文冒号、英文冒号、中文句号、中文感叹号、英文感叹号
    返回: (内容1, 内容2)
    用于将docx正文中，第1、2层小标题和正文文本混合的段落，按标点符号分开。
    """
    # 查找分隔符位置 - 支持全角符号
    # 全角冒号: ： (U+FF1A), 全角句号: ． (U+FF0E), 全角感叹号: ！ (U+FF01)
    delimiters = ['：', ':', '。', '．', '！', '!']
    
    # 查找第一个出现的分隔符
    for delimiter in delimiters:
        pos = text.find(delimiter)
        if pos != -1:
            # 找到分隔符，分割文本
            content1 = text[:pos + len(delimiter)].strip()
            content2 = text[pos + len(delimiter):].strip()
            logger.debug(f"在位置 {pos} 处发现分隔符 '{delimiter}'，分割文本: '{content1}' | '{content2}'")
            return content1, content2
    
    # 未找到分隔符
    logger.debug("未找到有效分隔符")
    return text, ""

def add_markdown_heading(content, heading_level):
    """
    根据标题级别添加Markdown标题符号
    
    参数:
        content: 标题内容
        heading_level: 标题级别
        
    返回:
        str: 添加了Markdown标题符号的内容
    """
    if heading_level == 1:
        return f"# {content}"
    elif heading_level == 2:
        return f"## {content}"
    elif heading_level == 3:
        return f"### {content}"
    elif heading_level == 4:
        return f"#### {content}"
    elif heading_level == 5:
        return f"##### {content}"
    else:
        return content

def convert_to_halfwidth(text: str) -> str:
    """
    将全角数字转换为半角数字
    
    参数:
        text: 原始文本
        
    返回:
        str: 转换后的文本
        
    功能说明:
        仅转换全角数字（U+FF10 到 U+FF19），其他字符保持不变
    """
    # 全角数字范围: U+FF10 到 U+FF19
    full_to_half = {chr(0xFF10 + i): chr(0x30 + i) for i in range(10)}
    return ''.join(full_to_half.get(char, char) for char in text)

# 模块测试代码
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(level=logging.DEBUG)
    logger.info("标题工具模块测试")
    
    # 测试全角数字标题
    test_text = "１．全角数字标题"
    cleaned, level = detect_heading_level(test_text)
    print(f"测试: '{test_text}' -> 清理后: '{cleaned}', 级别: {level}")
    
    # 测试全角括号标题
    test_text = "（１）全角括号标题"
    cleaned, level = detect_heading_level(test_text)
    print(f"测试: '{test_text}' -> 清理后: '{cleaned}', 级别: {level}")
    
    # 测试分割功能
    test_text = "一级标题：正文内容"
    part1, part2 = split_content_by_delimiters(test_text)
    print(f"分割测试: '{part1}' | '{part2}'")
    
    logger.info("模块测试完成!")
