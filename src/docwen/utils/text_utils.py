"""
文本处理工具模块
包含文本清理、安全获取值等工具函数
"""

import re
import logging
from collections.abc import Mapping, Iterable

from docwen.utils.validation_utils import is_value_empty

# 配置日志
logger = logging.getLogger(__name__)

def remove_colon(text):
    """
    去除文本结尾的中文或英文冒号
    :param text: 原始文本
    :return: 处理后的文本
    """
    return re.sub(r'[：:]$', '', text)

def extract_after_colon(text):
    """
    提取冒号之后的内容
    :param text: 原始文本
    :return: 冒号之后的内容
    """
    parts = re.split(r'[：:]', text, 1)
    return parts[1].strip() if len(parts) > 1 else text

def remove_brackets(text):
    """
    去除文本头尾的中文或英文括号
    :param text: 原始文本
    :return: 处理后的文本
    """
    # 去除开头和结尾的括号
    text = re.sub(r'^[（(]', '', text)
    text = re.sub(r'[)）]$', '', text)
    return text

def process_copy_to(text):
    """
    处理抄送机关文本
    :param text: 原始文本
    :return: 处理后的列表
    """
    # 去除开头的"抄送："或"抄送."等
    text = re.sub(r'^抄送[：:.]\s*', '', text)
    
    # 去除结尾的句号（若有）
    text = re.sub(r'[。.]$', '', text)
    
    # 使用中文逗号、英文逗号、中文顿号分割
    parts = re.split(r'[，,、]', text)
    
    # 去除每个部分两端的空格
    return [part.strip() for part in parts if part.strip()]

def process_attachment_item(text):
    """
    处理附件项文本
    :param text: 原始文本
    :return: 处理后的附件名称
    """
    # 去除开头的"附件："或"附件: "
    text = re.sub(r'^附件[：:]\s*', '', text)
    
    # 去除开头的序号（如：1. 2. 等）
    text = re.sub(r'^\d+[\.．]\s*', '', text)
    
    # 去除结尾的句号（若有）
    text = re.sub(r'[。.]$', '', text)
    
    return text.strip()

def replace_placeholder(text: str, yaml_data: dict) -> tuple:
    """
    替换文本中的占位符并返回处理状态
    参数:
        text: 原始文本
        yaml_data: YAML数据字典
    返回:
        tuple: (处理后的文本, 是否需要删除, 是否需要清空)
    """
    logger.debug(f"处理占位符: {text[:50]}...")
    placeholder_found = False
    should_delete = False
    should_clear = False
    
    # 定义需要特殊处理的占位符
    DELETE_IF_EMPTY = [
        '密级和保密期限', '紧急程度', '发文字号', '公开方式',
        '主送机关', '附注', '单附件说明', '多附件说明'
    ]
    
    DELETE_ROW_IF_EMPTY = ['抄送机关']
    
    # 遍历所有YAML键
    for key in yaml_data:
        placeholder = f'{{{{{key}}}}}'
        
        # 检查占位符是否存在
        if placeholder in text:
            placeholder_found = True
            value = yaml_data[key]
            logger.debug(f"找到占位符: {placeholder} -> {value}")
            
            # 检查是否需要删除段落
            if key in DELETE_IF_EMPTY and is_value_empty(value):
                should_delete = True
                logger.info(f"标记删除段落 (占位符: {placeholder})")
            
            # 检查是否需要删除表格行
            elif key in DELETE_ROW_IF_EMPTY and is_value_empty(value):
                should_clear = True
                logger.info(f"标记清空单元格 (占位符: {placeholder})")
            
            # 执行替换
            display_value = format_display_value(value)
            text = text.replace(placeholder, display_value)
            logger.debug(f"替换: {placeholder} -> {display_value}")
    
    return text, should_delete, should_clear

def convert_html_br_to_newline(text: str) -> str:
    """
    将 HTML 换行标签转换为换行符
    
    支持的格式：<br>, <br/>, <br />
    
    参数:
        text: 要处理的文本
    
    返回:
        转换后的文本（<br> 替换为 \n）
    """
    return re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)


def clean_text(text: str) -> str:
    """
    清理文本中的各种标记格式和特殊字符
    
    功能：
    - 将 <br> 标签转换为换行符（保留换行效果）
    - 清理其他HTML标签：<tag>内容</tag> 等
    - 清理HTML注释：<!-- 注释 -->
    - 转换HTML实体：&nbsp;、&lt;、&#数字; 等
    - 移除零宽字符：\u200B、\u200C、\uFEFF 等
    - 移除控制字符（保留换行、制表符）
    
    注意：
    - 不处理链接！链接处理由 link_embedding.process_markdown_links() 负责
    
    参数:
        text: 要清理的文本
    
    返回:
        清理后的纯文本
    """
    logger.debug(f"清理文本: {text[:50]}{'...' if len(text) > 50 else ''}")
    
    if not text:
        return text
    
    cleaned = text
    
    # 1. 先将 <br> 标签转换为换行符（在删除HTML标签之前）
    cleaned = convert_html_br_to_newline(cleaned)
    
    # 3. 处理HTML注释
    # 格式：<!-- 注释内容 -->
    html_comment_pattern = r'<!--.*?-->'
    cleaned = re.sub(html_comment_pattern, '', cleaned, flags=re.DOTALL)
    
    # 4. 处理HTML标签
    # 移除所有HTML标签，保留标签内的文本
    # 包括：<div>文本</div>、<br/>、<img src="..."/> 等
    html_tag_pattern = r'<[^>]+>'
    cleaned = re.sub(html_tag_pattern, '', cleaned)
    
    # 5. 处理HTML实体
    # 常见字符实体
    html_entities = {
        '&nbsp;': ' ',      # 非断开空格
        '&ensp;': ' ',      # 半角空格
        '&emsp;': '　',     # 全角空格
        '&lt;': '<',        # 小于号
        '&gt;': '>',        # 大于号
        '&amp;': '&',       # 和号
        '&quot;': '"',      # 双引号
        '&#39;': "'",       # 单引号
        '&apos;': "'",      # 单引号
        '&mdash;': '—',     # em破折号
        '&ndash;': '–',     # en破折号
        '&hellip;': '…',    # 省略号
        '&copy;': '©',      # 版权符号
        '&reg;': '®',       # 注册商标
        '&trade;': '™',     # 商标
        '&times;': '×',     # 乘号
        '&divide;': '÷',    # 除号
        '&middot;': '·',    # 中点
        '&bull;': '•',      # 项目符号
        '&laquo;': '«',     # 左双尖括号
        '&raquo;': '»',     # 右双尖括号
    }
    for entity, char in html_entities.items():
        cleaned = cleaned.replace(entity, char)
    
    # 处理数字实体：&#数字; 和 &#x十六进制;
    # 十进制数字实体
    cleaned = re.sub(r'&#(\d+);', lambda m: chr(int(m.group(1))), cleaned)
    # 十六进制数字实体
    cleaned = re.sub(r'&#[xX]([0-9a-fA-F]+);', lambda m: chr(int(m.group(1), 16)), cleaned)
    
    # 6. 移除零宽字符
    zero_width_chars = [
        '\u200B',  # 零宽空格
        '\u200C',  # 零宽非连接符
        '\u200D',  # 零宽连接符
        '\uFEFF',  # 零宽非断开空格（BOM）
        '\u2060',  # 字符连接符
    ]
    for char in zero_width_chars:
        cleaned = cleaned.replace(char, '')
    
    # 7. 移除控制字符（保留换行\n、回车\r、制表符\t）
    # ASCII控制字符范围：0x00-0x1F（除了\n、\r、\t）
    cleaned = re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F]', '', cleaned)
    
    logger.debug(f"清理后: {cleaned[:50]}{'...' if len(cleaned) > 50 else ''}")
    return cleaned

def clean_text_in_data(data):
    """
    递归清理数据结构中的所有文本标记和特殊字符
    
    支持嵌套的字典、列表、元组等数据结构，对所有字符串执行清理。
    清理内容包括：HTML标签、HTML实体、零宽字符、控制字符等。
    
    注意：不处理链接！链接由 link_embedding.process_markdown_links() 处理。
    
    参数:
        data: 要处理的数据 (可以是dict, list, str或其他类型)
        
    返回:
        清理后的数据结构（与输入类型相同）
    """
    if isinstance(data, Mapping):  # 处理字典类型
        return {k: clean_text_in_data(v) for k, v in data.items()}
    elif isinstance(data, Iterable) and not isinstance(data, str):  # 处理列表、元组等
        return type(data)(clean_text_in_data(item) for item in data)
    elif isinstance(data, str):  # 处理字符串
        return clean_text(data)
    else:  # 其他类型保持不变
        return data

def safe_get(data: dict, key: str, default="") -> str:
    """安全获取字典值，处理空值和None"""
    logger.debug(f"安全获取字段: {key}")
    
    value = data.get(key, default)
    
    # 处理各种空值情况
    if is_value_empty(value):
        logger.debug(f"字段 {key} 为空，使用默认值: {default}")
        return default
    
    logger.debug(f"获取到字段值: {value}")
    # 确保返回字符串类型，兼容 Cython 类型检查
    return str(value) if value is not None else default

def is_pure_number(value) -> bool:
    """
    检查值是否为纯数字（包括整数、浮点数、分数，支持正负数）
    例如: "123", "45.67", "-123", "-45.67", "1/2", "-3/4" 返回 True;
          "123abc", "12.34.56", "1/2/3" 返回 False
    """
    if value is None:
        return False
    str_value = str(value).strip()
    if not str_value:
        return False
    # 匹配整数、浮点数或分数，包括负数
    return bool(re.match(r'^-?(\d+(\.\d+)?|\d+/\d+)$', str_value))


def format_yaml_value(value) -> str:
    """
    格式化值为安全的YAML字符串
    
    使用PyYAML自动处理特殊字符、引号、布尔字面量等。
    
    处理场景：
    - 特殊字符开头（如 [ ] { } : # & * ! | > ' " @ ` % ? -）
    - 包含冒号+空格（: ）
    - 包含行内注释（ #）
    - 纯数字、布尔值、空值字面量（true, false, null 等）
    - 单引号和双引号混合
    - 换行符
    
    参数:
        value: 要格式化的值（字符串、数字或None）
        
    返回:
        str: 格式化后的YAML安全字符串（不含键名，只有值部分）
        
    示例:
        >>> format_yaml_value("报告")
        '报告'
        >>> format_yaml_value("[2024]年报")
        "'[2024]年报'"
        >>> format_yaml_value('他说"你好"')
        '\'他说"你好"\''
        >>> format_yaml_value("It's ok")
        '"It\'s ok"'
    """
    import yaml
    
    # 处理空值
    if value is None or value == '':
        return ''
    
    # 转为字符串
    str_value = str(value)
    
    # 使用 yaml.safe_dump 自动选择最佳格式
    # default_flow_style=True 保证输出单行格式
    # allow_unicode=True 保留中文等 Unicode 字符
    result = yaml.safe_dump(str_value, default_flow_style=True, allow_unicode=True)
    
    # safe_dump 会在末尾添加换行符和可能的文档结束标记，需要清理
    result = result.strip()
    if result.endswith('...'):
        result = result[:-3].strip()
    
    return result


def format_display_value(value, separator: str = "") -> str:
    """
    格式化显示值，避免显示None字符串
    
    处理嵌套列表：递归展平后用分隔符连接
    
    参数:
        value: 原始值（可能是 None、字符串、数字、列表等）
        separator: 列表元素拼接符，默认空字符串（无分隔符直接拼接）
                   常见值: "、"（中文顿号）、"，"（中文逗号）、", "（英文逗号+空格）
    
    返回:
        str: 格式化后的字符串
    """
    logger.debug(f"格式化值: {value}，分隔符: '{separator}'")
    
    if value is None:
        return ""
    
    if isinstance(value, list):
        # 递归展平嵌套列表
        def flatten(lst):
            """递归展平列表"""
            result = []
            for item in lst:
                if isinstance(item, list):
                    # 递归展平子列表
                    result.extend(flatten(item))
                elif item not in [None, "", "null", "None"]:
                    # 添加非空值
                    result.append(str(item))
            return result
        
        non_empty_items = flatten(value)
        return separator.join(non_empty_items)
    
    return str(value)


