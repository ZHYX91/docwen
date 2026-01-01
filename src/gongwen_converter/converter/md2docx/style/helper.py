"""
样式查找辅助模块

本模块提供在 Word 文档中查找各种样式的辅助函数。
用于在文档处理过程中智能查找标题样式、代码块样式、引用样式等。

主要功能：
- get_heading_style_name: 智能查找标题样式（支持中英文模板）
- get_code_block_style_name: 查找代码块段落样式
- get_quote_style_name: 查找引用块段落样式
- get_formula_block_style_name: 查找公式块段落样式

样式查找策略：
1. 使用 style_resolver 获取国际化样式名（首选）
2. 按 Style ID 查找（最可靠，不随语言变化）
3. 降级返回默认值或 None

国际化说明：
自定义样式（代码、引用、公式、列表、分隔线）通过 i18n/style_resolver.py 获取，
自动支持所有语言版本的样式名检测。
"""

import logging

# 导入 style_resolver 用于国际化样式名查找
from gongwen_converter.i18n import style_resolver

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
    
    查找策略：
    1. 使用 style_resolver 查找（自动支持所有语言版本）
    2. 按 Style ID 查找：CodeBlock（最可靠）
    3. 返回 None（调用方需要使用兼容方式处理）
    
    参数:
        doc: python-docx Document 对象
        config_mgr: 配置管理器（可选，保留向后兼容，目前未使用）
        
    返回:
        str: 样式名称，如果不存在则返回 None
    """
    # 1. 使用 style_resolver 获取可用的样式名（按语言优先级）
    usable_name = style_resolver.get_usable_name(doc, "code_block")
    if usable_name:
        logger.debug(f"通过 style_resolver 找到代码块样式: '{usable_name}'")
        return usable_name
    
    # 2. 尝试通过 Style ID 查找（最可靠，不随语言变化）
    try:
        for style in doc.styles:
            if style.type.value == 1 and style.style_id == 'CodeBlock':
                logger.debug(f"通过 Style ID 找到代码块样式: '{style.name}'")
                return style.name
    except Exception as e:
        logger.debug(f"通过 Style ID 查找代码块样式失败: {e}")
    
    logger.debug("未找到代码块样式，将使用底纹兼容方式")
    return None


def get_quote_style_name(doc, level: int, config_mgr=None) -> str:
    """
    获取引用块段落样式名称（支持国际化）
    
    查找策略：
    1. 使用 style_resolver 查找（自动支持所有语言版本）
    2. 按 Style ID 查找：QuoteN（最可靠）
    3. 返回 None（调用方使用基础样式）
    
    参数:
        doc: python-docx Document 对象
        level: 引用级别 (1-9)
        config_mgr: 配置管理器（可选，保留向后兼容，目前未使用）
        
    返回:
        str: 样式名称，如果不存在则返回 None
    """
    # 1. 使用 style_resolver 获取可用的样式名（按语言优先级）
    style_key = f"quote_{level}"
    usable_name = style_resolver.get_usable_name(doc, style_key)
    if usable_name:
        logger.debug(f"通过 style_resolver 找到引用样式（{level}级）: '{usable_name}'")
        return usable_name
    
    # 2. 尝试通过 Style ID 查找（最可靠，不随语言变化）
    try:
        target_style_id = f"Quote{level}"
        for style in doc.styles:
            if style.type.value == 1 and style.style_id == target_style_id:
                logger.debug(f"通过 Style ID 找到引用样式: '{style.name}'")
                return style.name
    except Exception as e:
        logger.debug(f"通过 Style ID 查找引用样式失败: {e}")
    
    logger.debug(f"未找到引用样式（{level}级），将使用基础样式")
    return None


def get_formula_block_style_name(doc, config_mgr=None) -> str:
    """
    获取公式块段落样式名称（支持国际化）
    
    查找策略：
    1. 使用 style_resolver 查找（自动支持所有语言版本）
    2. 按 Style ID 查找：FormulaBlock（最可靠）
    3. 返回 None（调用方使用基础样式）
    
    参数:
        doc: python-docx Document 对象
        config_mgr: 配置管理器（可选，保留向后兼容，目前未使用）
        
    返回:
        str: 样式名称，如果不存在则返回 None
    """
    # 1. 使用 style_resolver 获取可用的样式名（按语言优先级）
    usable_name = style_resolver.get_usable_name(doc, "formula_block")
    if usable_name:
        logger.debug(f"通过 style_resolver 找到公式块样式: '{usable_name}'")
        return usable_name
    
    # 2. 尝试通过 Style ID 查找（最可靠，不随语言变化）
    try:
        for style in doc.styles:
            if style.type.value == 1 and style.style_id == 'FormulaBlock':
                logger.debug(f"通过 Style ID 找到公式块样式: '{style.name}'")
                return style.name
    except Exception as e:
        logger.debug(f"通过 Style ID 查找公式块样式失败: {e}")
    
    logger.debug("未找到公式块样式，将使用基础样式")
    return None


def get_list_block_style_name(doc) -> str:
    """
    获取列表块段落样式名称（支持国际化）
    
    查找策略：
    1. 使用 style_resolver 查找（自动支持所有语言版本）
    2. 按 Style ID 查找：ListBlock（最可靠）
    3. 降级查找 List Paragraph 样式（Word 内置）
    4. 返回 None（调用方使用基础样式）
    
    参数:
        doc: python-docx Document 对象
        
    返回:
        str: 样式名称，如果不存在则返回 None
    """
    # 1. 使用 style_resolver 获取可用的样式名（按语言优先级）
    usable_name = style_resolver.get_usable_name(doc, "list_block")
    if usable_name:
        logger.debug(f"通过 style_resolver 找到列表块样式: '{usable_name}'")
        return usable_name
    
    # 2. 尝试通过 Style ID 查找
    try:
        for style_id in ['ListBlock', 'ListParagraph']:
            for style in doc.styles:
                if style.type.value == 1 and style.style_id == style_id:
                    logger.debug(f"通过 Style ID 找到列表块样式: '{style.name}'")
                    return style.name
    except Exception as e:
        logger.debug(f"通过 Style ID 查找列表块样式失败: {e}")
    
    # 3. 降级查找 Word 内置的 List Paragraph 样式
    for name in ["List Paragraph", "列表段落"]:
        try:
            if name in doc.styles:
                style = doc.styles[name]
                if style.type.value == 1:
                    logger.debug(f"找到列表块样式（Word 内置）: '{name}'")
                    return name
        except Exception:
            pass
    
    logger.debug("未找到列表块样式，将使用基础样式")
    return None


def get_inline_code_style_name(doc) -> str:
    """
    获取行内代码字符样式名称（支持国际化）
    
    查找策略：
    1. 使用 style_resolver 查找（自动支持所有语言版本）
    2. 按 Style ID 查找：InlineCode（最可靠）
    3. 返回 None（调用方需要使用兼容方式处理）
    
    参数:
        doc: python-docx Document 对象
        
    返回:
        str: 样式名称，如果不存在则返回 None
    """
    # 1. 使用 style_resolver 获取可用的样式名（按语言优先级）
    usable_name = style_resolver.get_usable_name(doc, "inline_code")
    if usable_name:
        logger.debug(f"通过 style_resolver 找到行内代码样式: '{usable_name}'")
        return usable_name
    
    # 2. 尝试通过 Style ID 查找（最可靠，不随语言变化）
    try:
        for style in doc.styles:
            # 字符样式的 type.value == 2
            if style.type.value == 2 and style.style_id == 'InlineCode':
                logger.debug(f"通过 Style ID 找到行内代码样式: '{style.name}'")
                return style.name
    except Exception as e:
        logger.debug(f"通过 Style ID 查找行内代码样式失败: {e}")
    
    logger.debug("未找到行内代码样式，将使用底纹兼容方式")
    return None


def get_horizontal_rule_style_name(doc, hr_num: str) -> str:
    """
    获取分隔线段落样式名称（支持国际化）
    
    参数:
        doc: python-docx Document 对象
        hr_num: 分隔线编号 ("1", "2", "3")
        
    返回:
        str: 样式名称，如果不存在则返回 None
    """
    # 1. 使用 style_resolver 获取可用的样式名（按语言优先级）
    style_key = f"horizontal_rule_{hr_num}"
    usable_name = style_resolver.get_usable_name(doc, style_key)
    if usable_name:
        logger.debug(f"通过 style_resolver 找到分隔线样式: '{usable_name}'")
        return usable_name
    
    # 2. 尝试通过 Style ID 查找（最可靠，不随语言变化）
    try:
        target_style_id = f"HorizontalRule{hr_num}"
        for style in doc.styles:
            if style.type.value == 1 and style.style_id == target_style_id:
                logger.debug(f"通过 Style ID 找到分隔线样式: '{style.name}'")
                return style.name
    except Exception as e:
        logger.debug(f"通过 Style ID 查找分隔线样式失败: {e}")
    
    logger.debug(f"未找到分隔线样式: HorizontalRule {hr_num}")
    return None
