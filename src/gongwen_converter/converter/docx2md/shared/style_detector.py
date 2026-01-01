"""
样式检测模块

负责检测段落和 Run 的样式类型（代码/引用），支持样式合并。

主要组件：
- detect_paragraph_style_type(): 检测段落样式类型
- detect_run_style_type(): 检测 Run 样式类型
- merge_consecutive_runs(): 合并连续同类型 Run
- is_full_paragraph_code_style(): 检测整段是否为代码样式

国际化说明：
样式检测通过 i18n/style_resolver.py 实现，自动合并以下数据源：
- 所有语言文件的 [styles] 节（如 "代码块"、"Code Block"）
- 配置文件的 *_aliases 列表（第三方软件内置样式名）
"""

import logging

logger = logging.getLogger(__name__)

# 导入 style_resolver 用于样式检测
from gongwen_converter.i18n import style_resolver


def detect_paragraph_style_type(para, config_manager):
    """
    检测段落样式类型（用于 DOCX→MD 转换）
    
    使用 style_resolver 进行样式匹配，自动合并：
    - 所有语言版本的样式名
    - 配置文件的第三方别名
    
    检测优先级：
    1. 代码块样式（code_block）
    2. 分级引用样式（quote_1 ~ quote_9）
    3. 模糊匹配（如果启用）
    
    参数:
        para: Word段落对象
        config_manager: 配置管理器实例
        
    返回:
        tuple: (style_type, style_value)
            - ('code_block', True) - 代码块样式
            - ('quote', level) - 引用样式，level 为引用级别 (1-9)
            - (None, None) - 普通段落
    """
    try:
        style_name = para.style.name if para.style else None
        if not style_name:
            return None, None
        
        style_name_lower = style_name.lower()
        
        # 1. 检测代码块段落样式
        if style_resolver.is_style_match(style_name, "code_block"):
            logger.debug("段落样式 '%s' 匹配代码块样式", style_name)
            return 'code_block', True
        
        # 2. 检测分级引用样式（使用 style_resolver 的专用方法）
        quote_level = style_resolver.get_quote_level(style_name)
        if quote_level > 0:
            logger.debug("段落样式 '%s' 匹配引用样式，级别 %d", style_name, quote_level)
            return 'quote', quote_level
        
        # 3. 模糊匹配
        if config_manager.get_code_fuzzy_match_enabled():
            code_keywords = config_manager.get_code_fuzzy_keywords()
            for keyword in code_keywords:
                if keyword.lower() in style_name_lower:
                    logger.debug("段落样式 '%s' 模糊匹配代码关键词 '%s'", style_name, keyword)
                    return 'code_block', True
        
        if config_manager.get_quote_fuzzy_match_enabled():
            quote_keywords = config_manager.get_quote_fuzzy_keywords()
            for keyword in quote_keywords:
                if keyword.lower() in style_name_lower:
                    logger.debug("段落样式 '%s' 模糊匹配引用关键词 '%s'", style_name, keyword)
                    return 'quote', 1
        
        return None, None
        
    except Exception as e:
        logger.debug("检测段落样式类型失败: %s", e)
        return None, None


def detect_run_style_type(run, config_manager):
    """
    检测 Run 样式类型（用于 DOCX→MD 转换）
    
    使用 style_resolver 进行样式匹配，自动合并：
    - 所有语言版本的样式名
    - 配置文件的第三方别名
    
    检测优先级：
    1. 行内代码字符样式（inline_code）
    2. 引用字符样式
    3. 关联字符样式检测（使用 startswith 匹配段落样式名）
    4. 模糊匹配（如果启用）
    
    参数:
        run: Word Run对象
        config_manager: 配置管理器实例
        
    返回:
        str: 样式类型
            - 'code' - 代码样式
            - 'quote' - 引用字符样式
            - None - 普通文本
    """
    try:
        style_name = run.style.name if run.style else None
        if not style_name:
            return None
        
        style_name_lower = style_name.lower()
        
        # 1. 检测行内代码字符样式
        if style_resolver.is_style_match(style_name, "inline_code"):
            logger.debug("Run样式 '%s' 匹配行内代码样式", style_name)
            return 'code'
        
        # 2. 检测引用字符样式
        # 检查是否是任何级别的引用样式
        quote_level = style_resolver.get_quote_level(style_name)
        if quote_level > 0:
            logger.debug("Run样式 '%s' 匹配引用样式，级别 %d", style_name, quote_level)
            return 'quote'
        
        # 3. 关联字符样式检测（使用 startswith）
        # WPS 会自动创建关联字符样式，如 "Code Block Char"、"代码块 Char"
        # 获取代码块和引用样式的所有可能名称
        code_block_names = style_resolver.get_all_detection_names("code_block")
        for para_style in code_block_names:
            if style_name_lower.startswith(para_style.lower()):
                logger.debug("Run样式 '%s' 是代码段落样式 '%s' 的关联字符样式", style_name, para_style)
                return 'code'
        
        # 检查引用段落样式的关联字符样式
        for level in range(1, 10):
            quote_names = style_resolver.get_all_detection_names(f"quote_{level}")
            for para_style in quote_names:
                if style_name_lower.startswith(para_style.lower()):
                    logger.debug("Run样式 '%s' 是引用段落样式 '%s' 的关联字符样式", style_name, para_style)
                    return 'quote'
        
        # 4. 模糊匹配
        if config_manager.get_code_fuzzy_match_enabled():
            code_keywords = config_manager.get_code_fuzzy_keywords()
            for keyword in code_keywords:
                if keyword.lower() in style_name_lower:
                    logger.debug("Run样式 '%s' 模糊匹配代码关键词 '%s'", style_name, keyword)
                    return 'code'
        
        if config_manager.get_quote_fuzzy_match_enabled():
            quote_keywords = config_manager.get_quote_fuzzy_keywords()
            for keyword in quote_keywords:
                if keyword.lower() in style_name_lower:
                    logger.debug("Run样式 '%s' 模糊匹配引用关键词 '%s'", style_name, keyword)
                    return 'quote'
        
        return None
        
    except Exception as e:
        logger.debug("检测Run样式类型失败: %s", e)
        return None


def is_full_paragraph_code_style(runs, config_manager, wps_shading_enabled=True, word_shading_enabled=True):
    """
    检测段落中所有 Run 是否都是代码样式
    
    用于实现 full_paragraph_as_block 配置：当整段文字都应用了代码字符样式时，
    应该输出代码块而不是多个连续的行内代码。
    
    参数:
        runs: Run 对象列表
        config_manager: 配置管理器实例
        wps_shading_enabled: 是否启用 WPS 底纹检测
        word_shading_enabled: 是否启用 Word 底纹检测
        
    返回:
        bool: 所有非空 Run 是否都是代码样式
    """
    from .markdown_converter import has_gray_shading
    
    # 过滤出有文本内容的 Run
    non_empty_runs = [run for run in runs if run.text and run.text.strip()]
    
    # 空段落不视为代码块
    if not non_empty_runs:
        return False
    
    # 检查每个 Run 是否都是代码样式
    for run in non_empty_runs:
        run_type = detect_run_style_type(run, config_manager)
        
        # 如果样式检测无结果，检查灰色底纹
        if run_type is None:
            if has_gray_shading(run, wps_shading_enabled, word_shading_enabled):
                run_type = 'code'
        
        # 如果有任何一个 Run 不是代码样式，返回 False
        if run_type != 'code':
            return False
    
    return True


def merge_consecutive_runs(runs, config_manager, wps_shading_enabled=None, word_shading_enabled=None):
    """
    合并连续同类型样式的 Run（用于 DOCX→MD 转换）
    
    合并规则：
    - 代码 + 代码 → 合并
    - 引用 + 引用 → 合并
    - 代码 + 引用 → 不合并
    - 任意 + 普通 → 不合并
    
    参数:
        runs: Run对象列表
        config_manager: 配置管理器实例
        wps_shading_enabled: 是否启用 WPS 底纹检测（None 时从配置读取）
        word_shading_enabled: 是否启用 Word 底纹检测（None 时从配置读取）
        
    返回:
        list: [{'text': '合并文本', 'type': 'code'/'quote'/None}, ...]
    """
    from .markdown_converter import has_gray_shading
    
    # 如果未指定，从配置读取
    if wps_shading_enabled is None:
        wps_shading_enabled = config_manager.is_wps_shading_enabled()
    if word_shading_enabled is None:
        word_shading_enabled = config_manager.is_word_shading_enabled()
    
    result = []
    current_text = ""
    current_type = None
    
    for run in runs:
        if not run.text:
            continue
        
        # 检测当前 Run 的样式类型
        run_type = detect_run_style_type(run, config_manager)
        
        # 如果样式检测无结果，检查灰色底纹（兼容旧文档）
        if run_type is None:
            if has_gray_shading(run, wps_shading_enabled, word_shading_enabled):
                run_type = 'code'
                logger.debug("Run '%s...' 通过底纹检测识别为代码", run.text[:20])
        
        # 判断是否需要合并
        if run_type == current_type and run_type is not None:
            # 同类型，合并文本
            current_text += run.text
        else:
            # 类型变化，保存当前块（如果有）
            if current_text:
                result.append({
                    'text': current_text,
                    'type': current_type
                })
            # 开始新块
            current_text = run.text
            current_type = run_type
    
    # 保存最后一个块
    if current_text:
        result.append({
            'text': current_text,
            'type': current_type
        })
    
    logger.debug("合并 %d 个Run为 %d 个块", len(runs), len(result))
    return result
