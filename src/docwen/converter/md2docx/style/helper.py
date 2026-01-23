"""
样式查找辅助模块

本模块提供在 Word 文档中查找各种样式的辅助函数。
用于在文档处理过程中智能查找标题样式、代码块样式、引用样式等。

主要功能：
- get_heading_style_name: 智能查找标题样式（支持中英文模板）
- get_code_block_style_name: 获取代码块段落样式名
- get_quote_style_name: 获取引用块段落样式名
- get_formula_block_style_name: 获取公式块段落样式名

样式获取策略：
样式已在 injector.py 中统一注入，本模块直接返回当前语言的样式名。

国际化说明：
所有自定义样式通过 i18n/style_resolver.py 获取当前语言的样式名。
"""

import logging

# 导入 style_resolver 用于获取国际化样式名
from docwen.i18n import style_resolver

# 配置日志
logger = logging.getLogger(__name__)


def get_heading_style_name(doc, level: int) -> str:
    """
    智能获取标题样式名称
    
    支持中英文模板和 Style ID 查找，确保在不同语言版本的 Word/WPS 中都能正确找到标题样式。
    
    查找策略：
    1. 按样式名称查找：Heading N, 标题 N, HeadingN
    2. 按 Style ID 查找：HeadingN（最可靠）
    3. 降级返回中文名称（尝试让 Word 自动激活）
    
    参数:
        doc: python-docx Document 对象
        level: 标题级别 (1-9)
        
    返回:
        str: 样式名称，如果完全找不到则返回默认的中文名称
    """
    # 候选样式名列表（优先级从高到低）
    candidates = [
        f"Heading {level}",  # 标准英文名称（带空格）
        f"标题 {level}",     # 标准中文名称（带空格）
        f"Heading{level}"    # 无空格版本
    ]
    
    # 1. 尝试直接按名称查找
    for style_name in candidates:
        try:
            if style_name in doc.styles:
                style = doc.styles[style_name]
                # 检查类型是否为段落样式 (1 = WD_STYLE_TYPE.PARAGRAPH)
                if style.type.value == 1:
                    logger.debug(f"找到标题样式(按名称): '{style_name}'")
                    return style_name
        except Exception:
            pass
            
    # 2. 遍历所有样式查找 Style ID（最可靠的方法）
    target_style_id = f"Heading{level}"
    
    try:
        for style in doc.styles:
            if style.type.value == 1:  # 只检查段落样式
                if style.style_id == target_style_id:
                    logger.debug(f"找到标题样式(按ID): ID='{style.style_id}', Name='{style.name}'")
                    return style.name
    except Exception as e:
        logger.warning(f"遍历样式查找ID时出错: {str(e)}")
        
    # 3. 降级返回中文名称（在中文Word中可能自动激活）
    logger.debug(f"未明确找到标题样式，尝试默认中文名称: '标题 {level}'")
    return f"标题 {level}"


def get_code_block_style_name(doc, config_mgr=None) -> str:
    """
    获取代码块段落样式名称（支持国际化）
    
    样式已在 injector.py 中注入，直接返回当前语言的样式名。
    
    参数:
        doc: python-docx Document 对象
        config_mgr: 配置管理器（可选，保留向后兼容，目前未使用）
        
    返回:
        str: 样式名称
    """
    # 直接返回当前语言的样式名（样式已在 injector.py 中注入）
    style_name = style_resolver.get_injection_name("code_block")
    logger.debug(f"代码块样式名: '{style_name}'")
    return style_name


def get_quote_style_name(doc, level: int, config_mgr=None) -> str:
    """
    获取引用块段落样式名称（支持国际化）
    
    样式已在 injector.py 中注入，直接返回当前语言的样式名。
    
    参数:
        doc: python-docx Document 对象
        level: 引用级别 (1-9)
        config_mgr: 配置管理器（可选，保留向后兼容，目前未使用）
        
    返回:
        str: 样式名称
    """
    # 直接返回当前语言的样式名（样式已在 injector.py 中注入）
    style_key = f"quote_{level}"
    style_name = style_resolver.get_injection_name(style_key)
    logger.debug(f"引用样式名（{level}级）: '{style_name}'")
    return style_name


def get_formula_block_style_name(doc, config_mgr=None) -> str:
    """
    获取公式块段落样式名称（支持国际化）
    
    样式已在 injector.py 中注入，直接返回当前语言的样式名。
    
    参数:
        doc: python-docx Document 对象
        config_mgr: 配置管理器（可选，保留向后兼容，目前未使用）
        
    返回:
        str: 样式名称
    """
    # 直接返回当前语言的样式名（样式已在 injector.py 中注入）
    style_name = style_resolver.get_injection_name("formula_block")
    logger.debug(f"公式块样式名: '{style_name}'")
    return style_name


def get_list_block_style_name(doc) -> str:
    """
    获取列表块段落样式名称（支持国际化）
    
    样式已在 injector.py 中注入，直接返回当前语言的样式名。
    
    参数:
        doc: python-docx Document 对象
        
    返回:
        str: 样式名称
    """
    # 直接返回当前语言的样式名（样式已在 injector.py 中注入）
    style_name = style_resolver.get_injection_name("list_block")
    logger.debug(f"列表块样式名: '{style_name}'")
    return style_name


def get_inline_code_style_name(doc) -> str:
    """
    获取行内代码字符样式名称（支持国际化）
    
    样式已在 injector.py 中注入，直接返回当前语言的样式名。
    
    参数:
        doc: python-docx Document 对象
        
    返回:
        str: 样式名称
    """
    # 直接返回当前语言的样式名（样式已在 injector.py 中注入）
    style_name = style_resolver.get_injection_name("inline_code")
    logger.debug(f"行内代码样式名: '{style_name}'")
    return style_name


def get_horizontal_rule_style_name(doc, hr_num: str) -> str:
    """
    获取分隔线段落样式名称（支持国际化）
    
    样式已在 injector.py 中注入，直接返回当前语言的样式名。
    
    参数:
        doc: python-docx Document 对象
        hr_num: 分隔线编号 ("1", "2", "3")
        
    返回:
        str: 样式名称
    """
    # 直接返回当前语言的样式名（样式已在 injector.py 中注入）
    style_key = f"horizontal_rule_{hr_num}"
    style_name = style_resolver.get_injection_name(style_key)
    logger.debug(f"分隔线样式名: '{style_name}'")
    return style_name
