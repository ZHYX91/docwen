"""
Markdown 转换模块

负责将 Word 段落转换为 Markdown 文本，支持格式保留和脚注处理。

主要组件：
- convert_paragraph_to_markdown(): 基础段落转换
- convert_paragraph_to_markdown_skip_prefix(): 跳过序号前缀的转换
- convert_paragraph_to_markdown_with_styles(): 带样式检测的完整转换
- apply_format_markers(): 应用格式标记
- has_gray_shading(): 检测灰色字符底纹
- has_paragraph_gray_shading(): 检测灰色段落底纹
- detect_note_reference_in_run(): 检测脚注/尾注引用
"""

import logging
from docx.enum.text import WD_UNDERLINE

logger = logging.getLogger(__name__)


# WPS 字符底纹使用的灰色颜色值（硬编码）
_WPS_SHADING_GRAY_COLORS = [
    "D9D9D9", "E7E6E6", "F2F2F2", "CCCCCC",
    "C0C0C0", "A6A6A6", "BFBFBF", "D0CECE",
]


def detect_note_reference_in_run(run) -> tuple:
    """
    检测 Run 中是否包含脚注或尾注引用
    
    Word 文档中的脚注/尾注引用存储在独立的 Run 中，包含 <w:footnoteReference> 或
    <w:endnoteReference> 元素。该函数检测这些元素并返回引用类型和 ID。
    
    参数:
        run: python-docx Run 对象
        
    返回:
        tuple: (引用类型, ID)
            - 脚注: ('footnote', int)
            - 尾注: ('endnote', int)
            - 无引用: (None, None)
    """
    WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    NSMAP = {'w': WORD_NS}
    
    try:
        # 检查脚注引用
        footnote_refs = run._r.findall('.//w:footnoteReference', NSMAP)
        for ref in footnote_refs:
            id_str = ref.get(f'{{{WORD_NS}}}id')
            if id_str:
                try:
                    note_id = int(id_str)
                    logger.debug(f"检测到脚注引用: ID={note_id}")
                    return ('footnote', note_id)
                except ValueError:
                    logger.warning(f"无效的脚注引用ID: {id_str}")
        
        # 检查尾注引用
        endnote_refs = run._r.findall('.//w:endnoteReference', NSMAP)
        for ref in endnote_refs:
            id_str = ref.get(f'{{{WORD_NS}}}id')
            if id_str:
                try:
                    note_id = int(id_str)
                    logger.debug(f"检测到尾注引用: ID={note_id}")
                    return ('endnote', note_id)
                except ValueError:
                    logger.warning(f"无效的尾注引用ID: {id_str}")
    
    except Exception as e:
        logger.debug(f"检测脚注/尾注引用时出错: {e}")
    
    return (None, None)


def has_gray_shading(run, wps_enabled=True, word_enabled=True):
    """
    检测Run是否有灰色字符底纹（用于识别代码）
    
    支持两种底纹实现方式：
    1. WPS纯色填充：w:fill="D9D9D9"（灰色）
    2. Word图案填充：w:val="pct15" + w:fill="FFFFFF"
    
    参数:
        run: Word Run对象
        wps_enabled: 是否启用 WPS 底纹检测
        word_enabled: 是否启用 Word 底纹检测
    
    返回:
        bool: 是否有灰色底纹
    """
    if not wps_enabled and not word_enabled:
        return False
    
    try:
        rPr = run._element.rPr
        if rPr is None:
            return False
        
        shd = rPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd')
        if shd is None:
            return False
        
        val = shd.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
        fill = shd.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill')
        
        # WPS纯色填充
        if wps_enabled and fill:
            fill_upper = fill.upper()
            for gray in _WPS_SHADING_GRAY_COLORS:
                if fill_upper == gray.upper():
                    return True
        
        # Word图案填充
        if word_enabled and val and val.startswith('pct'):
            if fill is None or fill.upper() in ['FFFFFF', 'AUTO']:
                return True
        
        return False
    except Exception as e:
        logger.debug(f"检测灰色底纹失败: {e}")
        return False


def has_paragraph_gray_shading(para, wps_enabled=True, word_enabled=True):
    """
    检测段落是否有灰色段落底纹（用于识别代码块）
    
    参数:
        para: Word段落对象
        wps_enabled: 是否启用 WPS 底纹检测
        word_enabled: 是否启用 Word 底纹检测
    
    返回:
        bool: 是否有灰色段落底纹
    """
    if not wps_enabled and not word_enabled:
        return False
    
    try:
        pPr = para._element.pPr
        if pPr is None:
            return False
        
        shd = pPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd')
        if shd is None:
            return False
        
        val = shd.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
        fill = shd.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill')
        
        # WPS纯色填充
        if wps_enabled and fill:
            fill_upper = fill.upper()
            for gray in _WPS_SHADING_GRAY_COLORS:
                if fill_upper == gray.upper():
                    return True
        
        # Word图案填充
        if word_enabled and val and val.startswith('pct'):
            if fill is None or fill.upper() in ['FFFFFF', 'AUTO']:
                return True
        
        return False
    except Exception as e:
        logger.debug(f"检测段落底纹失败: {e}")
        return False


def apply_format_markers(run, text, syntax_config, wps_shading_enabled=True, word_shading_enabled=True):
    """
    根据Run的格式属性添加Markdown标记
    
    格式应用顺序（从内到外）：
    1. 上标/下标  2. 删除线  3. 下划线  4. 斜体  5. 粗体  6. 高亮  7. 代码（灰色背景）
    
    参数:
        run: Word Run对象
        text: 原始文本
        syntax_config: 语法配置字典
        wps_shading_enabled: 是否启用 WPS 底纹检测
        word_shading_enabled: 是否启用 Word 底纹检测
    
    返回:
        str: 添加了Markdown标记的文本
    """
    # 检测灰色背景（代码）- 最高优先级
    if has_gray_shading(run, wps_shading_enabled, word_shading_enabled):
        return f'`{text}`'
    
    # 上标
    if run.font.superscript:
        if syntax_config.get('superscript') == 'extended':
            text = f'^{text}^'
        else:
            text = f'<sup>{text}</sup>'
    
    # 下标
    if run.font.subscript:
        if syntax_config.get('subscript') == 'extended':
            text = f'~{text}~'
        else:
            text = f'<sub>{text}</sub>'
    
    # 删除线
    if run.font.strike:
        if syntax_config.get('strikethrough') == 'extended':
            text = f'~~{text}~~'
        else:
            text = f'<del>{text}</del>'
    
    # 下划线
    if run.underline and run.underline != WD_UNDERLINE.NONE:
        text = f'<u>{text}</u>'
    
    # 斜体
    if run.italic:
        if syntax_config.get('italic') == 'asterisk':
            text = f'*{text}*'
        else:
            text = f'_{text}_'
    
    # 粗体
    if run.bold:
        if syntax_config.get('bold') == 'asterisk':
            text = f'**{text}**'
        else:
            text = f'__{text}__'
    
    # 高亮
    if run.font.highlight_color:
        if syntax_config.get('highlight') == 'extended':
            text = f'=={text}=='
        else:
            text = f'<mark>{text}</mark>'
    
    return text


def convert_paragraph_to_markdown(para, preserve_formatting=True, syntax_config=None, 
                                   wps_shading_enabled=True, word_shading_enabled=True):
    """
    将Word段落转换为Markdown文本，保留字符格式
    
    参数:
        para: Word段落对象
        preserve_formatting: 是否保留格式
        syntax_config: 语法配置字典
        wps_shading_enabled: 是否启用 WPS 底纹检测
        word_shading_enabled: 是否启用 Word 底纹检测
        
    返回:
        str: Markdown格式的文本
    """
    if not preserve_formatting:
        return para.text.strip()
    
    md_parts = []
    for run in para.runs:
        text = run.text
        if not text:
            continue
        
        text = apply_format_markers(run, text, syntax_config, wps_shading_enabled, word_shading_enabled)
        md_parts.append(text)
    
    return ''.join(md_parts)


def convert_paragraph_to_markdown_skip_prefix(para, skip_chars, preserve_formatting=True, syntax_config=None,
                                               wps_shading_enabled=True, word_shading_enabled=True):
    """
    将Word段落转换为Markdown文本，跳过前 skip_chars 个字符（序号部分）
    
    参数:
        para: Word段落对象
        skip_chars: 要跳过的字符数（序号长度）
        preserve_formatting: 是否保留格式
        syntax_config: 语法配置字典
        wps_shading_enabled: 是否启用 WPS 底纹检测
        word_shading_enabled: 是否启用 Word 底纹检测
        
    返回:
        str: Markdown格式的文本（已跳过序号）
    """
    if skip_chars <= 0:
        return convert_paragraph_to_markdown(para, preserve_formatting, syntax_config, 
                                              wps_shading_enabled, word_shading_enabled)
    
    if not preserve_formatting:
        return para.text[skip_chars:].strip()
    
    md_parts = []
    skipped = 0
    
    for run in para.runs:
        run_text = run.text
        if not run_text:
            continue
        
        if skipped < skip_chars:
            remaining_to_skip = skip_chars - skipped
            
            if len(run_text) <= remaining_to_skip:
                skipped += len(run_text)
                continue
            else:
                run_text = run_text[remaining_to_skip:]
                skipped = skip_chars
        
        formatted = apply_format_markers(run, run_text, syntax_config, wps_shading_enabled, word_shading_enabled)
        md_parts.append(formatted)
    
    return ''.join(md_parts)


def _process_text_with_notes_for_code_block(para, note_extractor) -> str:
    """处理代码块/引用块中的文本，包含脚注/尾注引用"""
    if note_extractor is None:
        return para.text
    
    parts = []
    for run in para.runs:
        ref_type, ref_id = detect_note_reference_in_run(run)
        if ref_type is not None:
            if ref_type == 'footnote':
                md_ref = note_extractor.get_footnote_reference_text(ref_id)
            else:
                md_ref = note_extractor.get_endnote_reference_text(ref_id)
            if md_ref:
                parts.append(md_ref)
            continue
        
        if run.text:
            parts.append(run.text)
    
    return ''.join(parts)


def convert_paragraph_to_markdown_with_styles(para, config_manager, preserve_formatting=True, syntax_config=None,
                                               wps_shading_enabled=None, word_shading_enabled=None,
                                               note_extractor=None):
    """
    将Word段落转换为Markdown文本（支持样式检测、Run 合并和脚注/尾注处理）
    
    参数:
        para: Word段落对象
        config_manager: 配置管理器实例
        preserve_formatting: 是否保留格式
        syntax_config: 语法配置字典
        wps_shading_enabled: 是否启用 WPS 底纹检测（None 时从配置读取）
        word_shading_enabled: 是否启用 Word 底纹检测（None 时从配置读取）
        note_extractor: NoteExtractor 对象
        
    返回:
        tuple: (markdown_text, style_type, style_value)
    """
    from .style_detector import detect_paragraph_style_type, detect_run_style_type, is_full_paragraph_code_style
    
    if wps_shading_enabled is None:
        wps_shading_enabled = config_manager.is_wps_shading_enabled()
    if word_shading_enabled is None:
        word_shading_enabled = config_manager.is_word_shading_enabled()
    
    # 1. 检测段落级样式
    style_type, style_value = detect_paragraph_style_type(para, config_manager)
    
    if style_type == 'code_block':
        text = _process_text_with_notes_for_code_block(para, note_extractor)
        return text, 'code_block', True
    elif style_type == 'quote':
        text = _process_text_with_notes_for_code_block(para, note_extractor)
        return text, 'quote', style_value
    
    # 2. 检测段落级灰色底纹
    if has_paragraph_gray_shading(para, wps_shading_enabled, word_shading_enabled):
        text = _process_text_with_notes_for_code_block(para, note_extractor)
        return text, 'code_block', True
    
    # 3. 整段字符样式检测
    if config_manager.get_code_full_paragraph_as_block():
        if is_full_paragraph_code_style(para.runs, config_manager, wps_shading_enabled, word_shading_enabled):
            logger.debug(f"整段字符样式检测：所有 Run 都是代码样式，输出为代码块")
            text = _process_text_with_notes_for_code_block(para, note_extractor)
            return text, 'code_block', True
    
    # 4. Run级处理
    if not preserve_formatting:
        md_parts = []
        for run in para.runs:
            ref_type, ref_id = detect_note_reference_in_run(run)
            if ref_type is not None and note_extractor is not None:
                if ref_type == 'footnote':
                    md_ref = note_extractor.get_footnote_reference_text(ref_id)
                else:
                    md_ref = note_extractor.get_endnote_reference_text(ref_id)
                if md_ref:
                    md_parts.append(md_ref)
                continue
            
            if not run.text:
                continue
            
            run_style_type = detect_run_style_type(run, config_manager)
            if run_style_type is None:
                if has_gray_shading(run, wps_shading_enabled, word_shading_enabled):
                    run_style_type = 'code'
            
            if run_style_type in ('code', 'quote'):
                md_parts.append(f"`{run.text}`")
            else:
                md_parts.append(run.text)
        
        return ''.join(md_parts), None, None
    
    # 保留格式时
    md_parts = []
    current_code_text = ""
    current_code_type = None
    
    for run in para.runs:
        ref_type, ref_id = detect_note_reference_in_run(run)
        if ref_type is not None and note_extractor is not None:
            if current_code_text:
                md_parts.append(f"`{current_code_text}`")
                current_code_text = ""
                current_code_type = None
            
            if ref_type == 'footnote':
                md_ref = note_extractor.get_footnote_reference_text(ref_id)
            else:
                md_ref = note_extractor.get_endnote_reference_text(ref_id)
            
            if md_ref:
                md_parts.append(md_ref)
            continue
        
        if not run.text:
            continue
        
        run_style_type = detect_run_style_type(run, config_manager)
        
        if run_style_type is None:
            if has_gray_shading(run, wps_shading_enabled, word_shading_enabled):
                run_style_type = 'code'
        
        if run_style_type in ('code', 'quote'):
            if run_style_type == current_code_type:
                current_code_text += run.text
            else:
                if current_code_text:
                    md_parts.append(f"`{current_code_text}`")
                current_code_text = run.text
                current_code_type = run_style_type
        else:
            if current_code_text:
                md_parts.append(f"`{current_code_text}`")
                current_code_text = ""
                current_code_type = None
            
            text = apply_format_markers(run, run.text, syntax_config, wps_shading_enabled, word_shading_enabled)
            md_parts.append(text)
    
    if current_code_text:
        md_parts.append(f"`{current_code_text}`")
    
    return ''.join(md_parts), None, None
