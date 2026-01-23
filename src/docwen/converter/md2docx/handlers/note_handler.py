"""
MD脚注/尾注处理模块

提供Markdown脚注和尾注的提取，以及Word脚注和尾注的插入功能。

主要功能:
- extract_notes(): 从Markdown文本中提取脚注和尾注定义
- insert_word_footnote(): 在Word文档中插入脚注
- insert_word_endnote(): 在Word文档中插入尾注
- is_endnote_id(): 判断ID是否为尾注（以endnote-开头）

脚注语法支持:
- 数字ID: [^1], [^2]
- 命名ID: [^note], [^ref-1]

尾注语法支持:
- [^endnote-1], [^endnote-2]
- [^endnote-note], [^endnote-ref]

技术实现:
- 使用正则表达式匹配脚注/尾注定义和引用
- 通过python-docx的底层XML API创建Word脚注/尾注
"""

import re
import logging
from typing import Dict, Tuple, List, Optional
from lxml import etree

logger = logging.getLogger(__name__)

# 尾注ID前缀
ENDNOTE_PREFIX = 'endnote-'

# OOXML 命名空间
WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NSMAP = {'w': WORD_NS}

# 脚注/尾注引用正则：匹配 [^id]（不在定义行首）
NOTE_REF_REGEX = re.compile(r'\[\^([^\]]+)\](?!:)')

# 脚注/尾注定义正则：匹配 [^id]: 内容
# 改进版：使用更简单可靠的逐行匹配方式
# 单行定义：[^id]: 内容（到行尾）
NOTE_DEF_LINE_PATTERN = re.compile(r'^\[\^([^\]]+)\]:\s*(.*)$', re.MULTILINE)


def is_endnote_id(note_id: str) -> bool:
    """
    判断ID是否为尾注ID
    
    参数:
        note_id: 脚注/尾注ID
    
    返回:
        bool: 是否为尾注（以endnote-开头）
    """
    return note_id.startswith(ENDNOTE_PREFIX)


def get_clean_endnote_id(note_id: str) -> str:
    """
    获取去除前缀的尾注ID
    
    参数:
        note_id: 完整的尾注ID（如 endnote-1）
    
    返回:
        str: 去除前缀后的ID（如 1）
    """
    if is_endnote_id(note_id):
        return note_id[len(ENDNOTE_PREFIX):]
    return note_id


def extract_notes(md_body: str) -> Tuple[Dict[str, str], Dict[str, str], str]:
    """
    从Markdown文本中提取脚注和尾注定义
    
    使用改进的逐行匹配方式，确保能正确提取所有定义。
    
    参数:
        md_body: Markdown正文内容
    
    返回:
        Tuple[Dict[str, str], Dict[str, str], str]: (脚注字典, 尾注字典, 清理后的正文)
        - 脚注字典: {脚注ID: 脚注内容}
        - 尾注字典: {尾注ID: 尾注内容}（ID不含endnote-前缀）
        - 清理后的正文: 移除脚注/尾注定义后的Markdown文本
    """
    logger.debug("开始提取Markdown脚注/尾注定义...")
    
    footnotes = {}
    endnotes = {}
    lines_to_remove = set()  # 记录要移除的行号
    
    lines = md_body.split('\n')
    i = 0
    n = len(lines)
    
    while i < n:
        line = lines[i]
        match = NOTE_DEF_LINE_PATTERN.match(line)
        
        if match:
            note_id = match.group(1).strip()
            note_content = match.group(2).strip()
            lines_to_remove.add(i)
            
            # 检查后续行是否是多行内容（以4空格或制表符开头）
            content_lines = [note_content] if note_content else []
            j = i + 1
            while j < n:
                next_line = lines[j]
                # 多行内容的续行：以4空格或制表符开头
                if next_line.startswith('    ') or next_line.startswith('\t'):
                    # 移除缩进并添加
                    if next_line.startswith('    '):
                        content_lines.append(next_line[4:])
                    else:
                        content_lines.append(next_line[1:])  # 移除制表符
                    lines_to_remove.add(j)
                    j += 1
                elif next_line.strip() == '':
                    # 空行可能是多行内容的一部分，检查下一行
                    # 如果下一行仍是缩进内容，则这是内容中的空行
                    if j + 1 < n and (lines[j + 1].startswith('    ') or lines[j + 1].startswith('\t')):
                        content_lines.append('')
                        lines_to_remove.add(j)
                        j += 1
                    else:
                        # 空行后不是缩进内容，定义结束
                        break
                else:
                    # 非缩进的非空行，定义结束
                    break
            
            # 合并内容
            note_content = '\n'.join(content_lines).strip()
            
            if is_endnote_id(note_id):
                # 尾注：去掉前缀存储
                clean_id = get_clean_endnote_id(note_id)
                endnotes[clean_id] = note_content
                logger.debug(f"提取尾注 [{note_id}]: {note_content[:50] if note_content else '(空)'}...")
            else:
                # 脚注
                footnotes[note_id] = note_content
                logger.debug(f"提取脚注 [{note_id}]: {note_content[:50] if note_content else '(空)'}...")
            
            i = j  # 跳到下一个未处理的行
        else:
            i += 1
    
    # 从正文中移除定义行
    cleaned_lines = [line for idx, line in enumerate(lines) if idx not in lines_to_remove]
    cleaned_body = '\n'.join(cleaned_lines)
    
    # 清理多余空行
    cleaned_body = re.sub(r'\n{3,}', '\n\n', cleaned_body)
    cleaned_body = cleaned_body.strip()
    
    logger.info(f"提取完成 | 脚注: {len(footnotes)} 个, 尾注: {len(endnotes)} 个")
    
    return footnotes, endnotes, cleaned_body


def find_note_references(text: str) -> List[Tuple[str, int, int, bool]]:
    """
    查找文本中的所有脚注/尾注引用
    
    参数:
        text: 要搜索的文本
    
    返回:
        List[Tuple[str, int, int, bool]]: [(ID, 起始位置, 结束位置, 是否尾注), ...]
    """
    references = []
    for match in NOTE_REF_REGEX.finditer(text):
        note_id = match.group(1)
        start = match.start()
        end = match.end()
        is_endnote = is_endnote_id(note_id)
        references.append((note_id, start, end, is_endnote))
    return references


def has_note_references(text: str) -> bool:
    """
    检查文本中是否包含脚注/尾注引用
    
    参数:
        text: 要检查的文本
    
    返回:
        bool: 是否包含引用
    """
    return bool(NOTE_REF_REGEX.search(text))


def replace_note_reference(text: str, note_id: str, replacement: str) -> str:
    """
    替换文本中的指定脚注/尾注引用
    
    参数:
        text: 原文本
        note_id: 要替换的脚注/尾注ID
        replacement: 替换内容
    
    返回:
        str: 替换后的文本
    """
    pattern = re.compile(r'\[\^' + re.escape(note_id) + r'\](?!:)')
    return pattern.sub(replacement, text)


# ==============================================================================
# Word 脚注/尾注 XML 创建函数
# ==============================================================================

def _create_formatted_runs(p: etree._Element, text: str) -> None:
    """
    解析Markdown格式并创建带格式的run元素
    
    参数:
        p: 段落XML元素
        text: 包含Markdown格式的文本
    """
    from .text_handler import parse_markdown_formatting
    
    # 解析Markdown格式
    segments = parse_markdown_formatting(text, mode="apply")
    
    for segment in segments:
        seg_text = segment['text']
        seg_formats = segment['formats']
        
        if not seg_text:
            continue
        
        # 创建run元素
        r = etree.SubElement(p, f'{{{WORD_NS}}}r')
        
        # 如果有格式，创建rPr
        if seg_formats:
            rPr = etree.SubElement(r, f'{{{WORD_NS}}}rPr')
            
            for fmt in seg_formats:
                if fmt == 'bold' or fmt == 'bold_italic':
                    b = etree.SubElement(rPr, f'{{{WORD_NS}}}b')
                if fmt == 'italic' or fmt == 'bold_italic':
                    i = etree.SubElement(rPr, f'{{{WORD_NS}}}i')
                if fmt == 'strikethrough':
                    strike = etree.SubElement(rPr, f'{{{WORD_NS}}}strike')
                if fmt == 'underline':
                    u = etree.SubElement(rPr, f'{{{WORD_NS}}}u')
                    u.set(f'{{{WORD_NS}}}val', 'single')
                if fmt == 'superscript':
                    vertAlign = etree.SubElement(rPr, f'{{{WORD_NS}}}vertAlign')
                    vertAlign.set(f'{{{WORD_NS}}}val', 'superscript')
                if fmt == 'subscript':
                    vertAlign = etree.SubElement(rPr, f'{{{WORD_NS}}}vertAlign')
                    vertAlign.set(f'{{{WORD_NS}}}val', 'subscript')
                if fmt == 'highlight':
                    highlight = etree.SubElement(rPr, f'{{{WORD_NS}}}highlight')
                    highlight.set(f'{{{WORD_NS}}}val', 'yellow')
                if fmt == 'code':
                    # 代码格式：使用等宽字体和灰色底纹
                    rFonts = etree.SubElement(rPr, f'{{{WORD_NS}}}rFonts')
                    rFonts.set(f'{{{WORD_NS}}}ascii', 'Consolas')
                    rFonts.set(f'{{{WORD_NS}}}hAnsi', 'Consolas')
                    rFonts.set(f'{{{WORD_NS}}}eastAsia', 'Consolas')
                    # 底纹
                    shd = etree.SubElement(rPr, f'{{{WORD_NS}}}shd')
                    shd.set(f'{{{WORD_NS}}}val', 'clear')
                    shd.set(f'{{{WORD_NS}}}fill', 'D9D9D9')
        
        # 创建文本元素
        t = etree.SubElement(r, f'{{{WORD_NS}}}t')
        # 如果文本有前导或尾随空格，需要保留
        if seg_text.startswith(' ') or seg_text.endswith(' '):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = seg_text


def create_footnote_element(
    footnote_id: int, 
    content: str,
    text_style_id: str = 'FootnoteText',
    ref_style_id: str = 'FootnoteReference'
) -> etree._Element:
    """
    创建脚注内容XML元素
    
    参数:
        footnote_id: 脚注ID（从1开始）
        content: 脚注内容（支持多行，用\\n分隔；支持Markdown格式）
        text_style_id: 脚注文本的段落样式ID
        ref_style_id: 脚注引用的字符样式ID
    
    返回:
        etree._Element: 脚注XML元素
    
    注意:
        - styleId 在不同Word版本中不同（如中文版为ae/af0，英文版为FootnoteText/FootnoteReference）
        - 调用时应先通过 get_style_id_by_name() 获取实际的 styleId
        - 内容支持Markdown格式（粗体、斜体、代码等）
    """
    footnote = etree.Element(f'{{{WORD_NS}}}footnote')
    footnote.set(f'{{{WORD_NS}}}id', str(footnote_id))
    
    # 支持多段落内容
    paragraphs = content.split('\n') if '\n' in content else [content]
    
    for idx, para_text in enumerate(paragraphs):
        p = etree.SubElement(footnote, f'{{{WORD_NS}}}p')
        
        # 段落属性 - 必须设置段落样式
        pPr = etree.SubElement(p, f'{{{WORD_NS}}}pPr')
        pStyle = etree.SubElement(pPr, f'{{{WORD_NS}}}pStyle')
        pStyle.set(f'{{{WORD_NS}}}val', text_style_id)
        
        # 仅第一段包含脚注引用符号
        if idx == 0:
            # 脚注引用run
            r1 = etree.SubElement(p, f'{{{WORD_NS}}}r')
            rPr1 = etree.SubElement(r1, f'{{{WORD_NS}}}rPr')
            rStyle = etree.SubElement(rPr1, f'{{{WORD_NS}}}rStyle')
            rStyle.set(f'{{{WORD_NS}}}val', ref_style_id)
            etree.SubElement(r1, f'{{{WORD_NS}}}footnoteRef')
            
            # 空格run
            r2 = etree.SubElement(p, f'{{{WORD_NS}}}r')
            t2 = etree.SubElement(r2, f'{{{WORD_NS}}}t')
            t2.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t2.text = ' '
        
        # 内容：解析Markdown格式并创建带格式的run
        if para_text.strip():
            _create_formatted_runs(p, para_text)
    
    return footnote


def create_endnote_element(
    endnote_id: int, 
    content: str,
    text_style_id: str = 'EndnoteText',
    ref_style_id: str = 'EndnoteReference'
) -> etree._Element:
    """
    创建尾注内容XML元素
    
    参数:
        endnote_id: 尾注ID（从1开始）
        content: 尾注内容（支持多行，用\\n分隔；支持Markdown格式）
        text_style_id: 尾注文本的段落样式ID
        ref_style_id: 尾注引用的字符样式ID
    
    返回:
        etree._Element: 尾注XML元素
    
    注意:
        - styleId 在不同Word版本中不同（如中文版为af1/af3，WPS为3/7）
        - 调用时应先通过 get_style_id_by_name() 获取实际的 styleId
        - 内容支持Markdown格式（粗体、斜体、代码等）
    """
    endnote = etree.Element(f'{{{WORD_NS}}}endnote')
    endnote.set(f'{{{WORD_NS}}}id', str(endnote_id))
    
    # 支持多段落内容
    paragraphs = content.split('\n') if '\n' in content else [content]
    
    for idx, para_text in enumerate(paragraphs):
        p = etree.SubElement(endnote, f'{{{WORD_NS}}}p')
        
        # 段落属性 - 必须设置段落样式
        pPr = etree.SubElement(p, f'{{{WORD_NS}}}pPr')
        pStyle = etree.SubElement(pPr, f'{{{WORD_NS}}}pStyle')
        pStyle.set(f'{{{WORD_NS}}}val', text_style_id)
        
        # 仅第一段包含尾注引用符号
        if idx == 0:
            # 尾注引用run
            r1 = etree.SubElement(p, f'{{{WORD_NS}}}r')
            rPr1 = etree.SubElement(r1, f'{{{WORD_NS}}}rPr')
            rStyle = etree.SubElement(rPr1, f'{{{WORD_NS}}}rStyle')
            rStyle.set(f'{{{WORD_NS}}}val', ref_style_id)
            etree.SubElement(r1, f'{{{WORD_NS}}}endnoteRef')
            
            # 空格run
            r2 = etree.SubElement(p, f'{{{WORD_NS}}}r')
            t2 = etree.SubElement(r2, f'{{{WORD_NS}}}t')
            t2.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t2.text = ' '
        
        # 内容：解析Markdown格式并创建带格式的run
        if para_text.strip():
            _create_formatted_runs(p, para_text)
    
    return endnote


def create_footnote_reference_run(footnote_id: int, ref_style_id: str = 'FootnoteReference') -> etree._Element:
    """
    创建正文中的脚注引用元素（一个run）
    
    参数:
        footnote_id: 脚注ID
        ref_style_id: 脚注引用样式ID
    
    返回:
        etree._Element: run 元素
    """
    r = etree.Element(f'{{{WORD_NS}}}r')
    
    rPr = etree.SubElement(r, f'{{{WORD_NS}}}rPr')
    rStyle = etree.SubElement(rPr, f'{{{WORD_NS}}}rStyle')
    rStyle.set(f'{{{WORD_NS}}}val', ref_style_id)
    
    footnoteRef = etree.SubElement(r, f'{{{WORD_NS}}}footnoteReference')
    footnoteRef.set(f'{{{WORD_NS}}}id', str(footnote_id))
    
    return r


def create_endnote_reference_run(endnote_id: int, ref_style_id: str = 'EndnoteReference') -> etree._Element:
    """
    创建正文中的尾注引用元素（一个run）
    
    参数:
        endnote_id: 尾注ID
        ref_style_id: 尾注引用样式ID
    
    返回:
        etree._Element: run 元素
    """
    r = etree.Element(f'{{{WORD_NS}}}r')
    
    rPr = etree.SubElement(r, f'{{{WORD_NS}}}rPr')
    rStyle = etree.SubElement(rPr, f'{{{WORD_NS}}}rStyle')
    rStyle.set(f'{{{WORD_NS}}}val', ref_style_id)
    
    endnoteRef = etree.SubElement(r, f'{{{WORD_NS}}}endnoteReference')
    endnoteRef.set(f'{{{WORD_NS}}}id', str(endnote_id))
    
    return r


# ==============================================================================
# 样式ID辅助函数
# ==============================================================================

# 样式名到默认styleId的映射
DEFAULT_STYLE_IDS = {
    'footnote reference': 'FootnoteReference',
    'footnote text': 'FootnoteText',
    'endnote reference': 'EndnoteReference',
    'endnote text': 'EndnoteText',
}


def get_style_id_by_name(doc, style_name: str) -> str:
    """
    通过样式名称获取真实的 styleId
    
    由于不同语言版本的 Word 和 WPS 可能使用不同的 styleId
    （如中文 Word 可能使用 'a5' 而非 'FootnoteText'），
    因此需要通过 name 属性查找真实的 styleId。
    
    参数:
        doc: Document 对象
        style_name: 样式名称（如 'footnote reference'）
    
    返回:
        str: 对应的 styleId，如果未找到则返回默认值
    """
    target_name = style_name.lower()
    
    try:
        # 方法1：通过 python-docx 的 styles 遍历
        # 这是最可靠的方式，支持所有 python-docx 版本
        for style in doc.styles:
            # 获取样式的 XML 元素
            style_element = style._element
            name_elem = style_element.find('.//w:name', NSMAP)
            if name_elem is not None:
                name_val = name_elem.get(f'{{{WORD_NS}}}val', '')
                if name_val.lower() == target_name:
                    style_id = style_element.get(f'{{{WORD_NS}}}styleId')
                    if style_id:
                        logger.debug(f"找到样式 '{style_name}' -> styleId='{style_id}'")
                        return style_id
        
        # 方法2：如果 styles 迭代失败，尝试从关联部件获取
        # 某些模板可能有隐藏样式不在 doc.styles 中
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        try:
            styles_part = doc.part.part_related_by(RT.STYLES)
            if styles_part is not None:
                styles_element = styles_part._element
                for style in styles_element.findall('.//w:style', NSMAP):
                    name_elem = style.find('w:name', NSMAP)
                    if name_elem is not None:
                        name_val = name_elem.get(f'{{{WORD_NS}}}val', '')
                        if name_val.lower() == target_name:
                            style_id = style.get(f'{{{WORD_NS}}}styleId')
                            if style_id:
                                logger.debug(f"找到样式 '{style_name}' -> styleId='{style_id}' (从关联部件)")
                                return style_id
        except Exception as e:
            logger.debug(f"从关联部件获取样式失败: {e}")
        
        # 未找到，返回默认值
        default_id = DEFAULT_STYLE_IDS.get(target_name, style_name)
        logger.debug(f"未找到样式 '{style_name}'，使用默认 styleId='{default_id}'")
        return default_id
        
    except Exception as e:
        logger.warning(f"获取样式ID失败: {e}")
        return DEFAULT_STYLE_IDS.get(target_name, style_name)


def get_all_note_style_ids(doc) -> dict:
    """
    获取文档中所有脚注/尾注相关样式的ID
    
    参数:
        doc: Document 对象
    
    返回:
        dict: {
            'footnote_text': str,
            'footnote_ref': str,
            'endnote_text': str,
            'endnote_ref': str
        }
    """
    return {
        'footnote_text': get_style_id_by_name(doc, 'footnote text'),
        'footnote_ref': get_style_id_by_name(doc, 'footnote reference'),
        'endnote_text': get_style_id_by_name(doc, 'endnote text'),
        'endnote_ref': get_style_id_by_name(doc, 'endnote reference'),
    }


# ==============================================================================
# 脚注/尾注上下文管理
# ==============================================================================

class NoteContext:
    """
    脚注/尾注处理上下文
    
    管理脚注和尾注的ID分配、内容存储和样式ID
    """
    
    def __init__(self, footnotes: Dict[str, str] = None, endnotes: Dict[str, str] = None):
        """
        初始化脚注/尾注上下文
        
        参数:
            footnotes: 脚注字典 {md_id: content}
            endnotes: 尾注字典 {md_id: content}（ID已去除endnote-前缀）
        """
        self.footnotes = footnotes or {}
        self.endnotes = endnotes or {}
        
        # MD ID 到 Word ID 的映射
        self._footnote_id_map: Dict[str, int] = {}
        self._endnote_id_map: Dict[str, int] = {}
        
        # Word ID 计数器（从1开始）
        self._footnote_counter = 1
        self._endnote_counter = 1
        
        # 样式ID（需要通过文档初始化）
        self.footnote_text_style = 'FootnoteText'
        self.footnote_ref_style = 'FootnoteReference'
        self.endnote_text_style = 'EndnoteText'
        self.endnote_ref_style = 'EndnoteReference'
        
        # 已创建的脚注/尾注元素列表
        self.footnote_elements: List[etree._Element] = []
        self.endnote_elements: List[etree._Element] = []
    
    def init_styles_from_doc(self, doc):
        """
        从文档初始化样式ID
        
        参数:
            doc: Document 对象
        """
        style_ids = get_all_note_style_ids(doc)
        self.footnote_text_style = style_ids['footnote_text']
        self.footnote_ref_style = style_ids['footnote_ref']
        self.endnote_text_style = style_ids['endnote_text']
        self.endnote_ref_style = style_ids['endnote_ref']
        
        logger.debug(f"样式ID初始化完成: {style_ids}")
    
    def get_footnote_word_id(self, md_id: str) -> Optional[int]:
        """
        获取脚注的Word ID，如果未创建则创建新的
        
        参数:
            md_id: Markdown脚注ID
        
        返回:
            int: Word脚注ID，如果MD ID不存在则返回None
        """
        if md_id not in self.footnotes:
            logger.warning(f"脚注 [{md_id}] 未定义")
            return None
        
        if md_id not in self._footnote_id_map:
            word_id = self._footnote_counter
            self._footnote_id_map[md_id] = word_id
            self._footnote_counter += 1
            
            # 创建脚注元素
            content = self.footnotes[md_id]
            footnote_elem = create_footnote_element(
                word_id, content,
                self.footnote_text_style,
                self.footnote_ref_style
            )
            self.footnote_elements.append(footnote_elem)
            
            logger.debug(f"创建脚注 [{md_id}] -> Word ID {word_id}")
        
        return self._footnote_id_map[md_id]
    
    def get_endnote_word_id(self, md_id: str) -> Optional[int]:
        """
        获取尾注的Word ID，如果未创建则创建新的
        
        参数:
            md_id: Markdown尾注ID（不含endnote-前缀）
        
        返回:
            int: Word尾注ID，如果MD ID不存在则返回None
        """
        # 兼容处理：如果传入的是完整ID，去除前缀
        clean_id = get_clean_endnote_id(md_id)
        
        if clean_id not in self.endnotes:
            logger.warning(f"尾注 [endnote-{clean_id}] 未定义")
            return None
        
        if clean_id not in self._endnote_id_map:
            word_id = self._endnote_counter
            self._endnote_id_map[clean_id] = word_id
            self._endnote_counter += 1
            
            # 创建尾注元素
            content = self.endnotes[clean_id]
            endnote_elem = create_endnote_element(
                word_id, content,
                self.endnote_text_style,
                self.endnote_ref_style
            )
            self.endnote_elements.append(endnote_elem)
            
            logger.debug(f"创建尾注 [endnote-{clean_id}] -> Word ID {word_id}")
        
        return self._endnote_id_map[clean_id]
    
    def create_footnote_ref_run(self, md_id: str) -> Optional[etree._Element]:
        """
        创建脚注引用run
        
        参数:
            md_id: Markdown脚注ID
        
        返回:
            etree._Element: 脚注引用run，如果ID不存在则返回None
        """
        word_id = self.get_footnote_word_id(md_id)
        if word_id is None:
            return None
        return create_footnote_reference_run(word_id, self.footnote_ref_style)
    
    def create_endnote_ref_run(self, md_id: str) -> Optional[etree._Element]:
        """
        创建尾注引用run
        
        参数:
            md_id: Markdown尾注ID（可以包含或不包含endnote-前缀）
        
        返回:
            etree._Element: 尾注引用run，如果ID不存在则返回None
        """
        clean_id = get_clean_endnote_id(md_id)
        word_id = self.get_endnote_word_id(clean_id)
        if word_id is None:
            return None
        return create_endnote_reference_run(word_id, self.endnote_ref_style)
    
    @property
    def has_footnotes(self) -> bool:
        """是否有脚注"""
        return len(self.footnotes) > 0
    
    @property
    def has_endnotes(self) -> bool:
        """是否有尾注"""
        return len(self.endnotes) > 0
    
    @property
    def has_notes(self) -> bool:
        """是否有脚注或尾注"""
        return self.has_footnotes or self.has_endnotes


# ==============================================================================
# DOCX 脚注/尾注写入函数
# ==============================================================================

def write_notes_to_docx(docx_path: str, note_ctx: 'NoteContext') -> None:
    """
    将脚注和尾注内容写入DOCX文件
    
    此函数打开已保存的DOCX文件，向 footnotes.xml 和 endnotes.xml 
    追加已创建的脚注/尾注元素。
    
    参数:
        docx_path: DOCX文件路径
        note_ctx: NoteContext对象，包含要写入的脚注/尾注元素
    
    注意:
        - 必须在 xml_processor 处理之前调用
        - xml_processor 会处理所有 XML 文件，如果先执行可能导致脚注丢失
    """
    import zipfile
    import os
    import tempfile
    import shutil
    
    # 检查是否有需要写入的内容
    if not note_ctx.footnote_elements and not note_ctx.endnote_elements:
        logger.debug("没有脚注/尾注需要写入")
        return
    
    logger.info(f"写入脚注/尾注 | 脚注: {len(note_ctx.footnote_elements)} 个, "
                f"尾注: {len(note_ctx.endnote_elements)} 个")
    
    try:
        # 创建临时文件
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, "temp_notes.docx")
        
        with zipfile.ZipFile(docx_path, 'r') as zf_in:
            with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                for item in zf_in.infolist():
                    content = zf_in.read(item.filename)
                    
                    # 处理 footnotes.xml
                    if item.filename == 'word/footnotes.xml' and note_ctx.footnote_elements:
                        content = _append_notes_to_xml(
                            content, note_ctx.footnote_elements, 'footnote'
                        )
                        logger.debug(f"已向 footnotes.xml 添加 "
                                    f"{len(note_ctx.footnote_elements)} 个脚注")
                    
                    # 处理 endnotes.xml
                    elif item.filename == 'word/endnotes.xml' and note_ctx.endnote_elements:
                        content = _append_notes_to_xml(
                            content, note_ctx.endnote_elements, 'endnote'
                        )
                        logger.debug(f"已向 endnotes.xml 添加 "
                                    f"{len(note_ctx.endnote_elements)} 个尾注")
                    
                    zf_out.writestr(item, content)
        
        # 替换原文件
        shutil.move(temp_path, docx_path)
        
        # 清理临时目录
        try:
            os.rmdir(temp_dir)
        except OSError:
            pass
        
        logger.info("脚注/尾注写入完成")
        
    except Exception as e:
        logger.error(f"写入脚注/尾注失败: {e}", exc_info=True)
        raise


def _append_notes_to_xml(
    xml_content: bytes, 
    note_elements: List[etree._Element], 
    note_type: str
) -> bytes:
    """
    向 footnotes.xml 或 endnotes.xml 追加脚注/尾注元素
    
    参数:
        xml_content: 原始XML内容（bytes）
        note_elements: 要追加的脚注/尾注元素列表
        note_type: 'footnote' 或 'endnote'（用于日志）
    
    返回:
        bytes: 修改后的XML内容
    """
    root = etree.fromstring(xml_content)
    
    # 追加新的脚注/尾注元素
    for elem in note_elements:
        root.append(elem)
    
    logger.debug(f"向 {note_type}s.xml 追加了 {len(note_elements)} 个元素")
    
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone='yes')


def process_text_with_note_references(
    text: str, 
    paragraph, 
    fonts: dict, 
    formatting_mode: str, 
    doc, 
    note_ctx: 'NoteContext'
) -> None:
    """
    处理文本中的脚注/尾注引用，将 [^id] 替换为Word脚注/尾注引用
    
    此函数解析文本中的脚注/尾注引用语法（如 [^1]、[^endnote-1]），
    将其转换为Word原生的脚注/尾注引用，同时保留其他文本的Markdown格式。
    
    参数:
        text: 包含脚注/尾注引用的文本
        paragraph: Word段落对象（python-docx Paragraph）
        fonts: 字体格式信息字典（从模板提取）
        formatting_mode: Markdown格式处理模式（'apply', 'remove', 'keep'）
        doc: Document对象
        note_ctx: NoteContext对象，管理脚注/尾注的创建
    
    处理流程:
        1. 查找文本中所有 [^id] 格式的引用
        2. 将引用前的文本作为普通格式化文本添加
        3. 将引用转换为Word脚注/尾注引用元素
        4. 将引用后的文本作为普通格式化文本添加
    """
    from .text_handler import add_formatted_text_to_paragraph
    
    # 查找所有脚注/尾注引用
    references = list(NOTE_REF_REGEX.finditer(text))
    
    if not references:
        # 没有引用，直接添加文本
        add_formatted_text_to_paragraph(paragraph, text, fonts, formatting_mode, doc=doc)
        return
    
    logger.debug(f"处理文本中的 {len(references)} 个脚注/尾注引用")
    
    # 有引用，需要分段处理
    last_end = 0
    
    for match in references:
        note_id = match.group(1)
        start = match.start()
        end = match.end()
        
        # 添加引用前的文本
        if start > last_end:
            before_text = text[last_end:start]
            add_formatted_text_to_paragraph(
                paragraph, before_text, fonts, formatting_mode, doc=doc
            )
        
        # 创建并插入脚注/尾注引用
        if is_endnote_id(note_id):
            # 尾注引用
            clean_id = get_clean_endnote_id(note_id)
            ref_run = note_ctx.create_endnote_ref_run(clean_id)
            if ref_run is not None:
                paragraph._p.append(ref_run)
                logger.debug(f"插入尾注引用: [^{note_id}]")
            else:
                # 尾注未定义，保留原始文本
                paragraph.add_run(f"[^{note_id}]")
                logger.warning(f"尾注未定义: [^{note_id}]")
        else:
            # 脚注引用
            ref_run = note_ctx.create_footnote_ref_run(note_id)
            if ref_run is not None:
                paragraph._p.append(ref_run)
                logger.debug(f"插入脚注引用: [^{note_id}]")
            else:
                # 脚注未定义，保留原始文本
                paragraph.add_run(f"[^{note_id}]")
                logger.warning(f"脚注未定义: [^{note_id}]")
        
        last_end = end
    
    # 添加最后一个引用后的文本
    if last_end < len(text):
        after_text = text[last_end:]
        add_formatted_text_to_paragraph(
            paragraph, after_text, fonts, formatting_mode, doc=doc
        )
