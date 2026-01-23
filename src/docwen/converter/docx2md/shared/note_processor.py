"""
Word脚注/尾注提取模块

从Word文档中提取脚注和尾注并转换为Markdown格式。

主要功能:
- extract_footnotes_from_docx(): 从DOCX提取所有脚注
- extract_endnotes_from_docx(): 从DOCX提取所有尾注
- build_footnote_definitions(): 生成Markdown脚注定义块
- build_endnote_definitions(): 生成Markdown尾注定义块

技术实现:
- 优先通过python-docx访问文档的footnotes_part和endnotes_part
- 备用方案：直接从DOCX ZIP包读取word/footnotes.xml和word/endnotes.xml
- 使用w:type属性判断系统保留项（兼容Word和WPS）
"""

import logging
import zipfile
from typing import Dict, List, Tuple, Optional
from lxml import etree

logger = logging.getLogger(__name__)

# OOXML 命名空间
WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NSMAP = {'w': WORD_NS}

# 尾注ID前缀（用于Markdown）
ENDNOTE_PREFIX = 'endnote-'


def is_system_note(note_element) -> bool:
    """
    判断是否为系统保留的脚注/尾注
    
    通过w:type属性判断，而非固定ID值，以兼容Word和WPS。
    Word 使用 ID -1 和 0，WPS 使用 ID 4 和 5，但都有 type 属性。
    
    参数:
        note_element: w:footnote 或 w:endnote 元素
        
    返回:
        bool: 是否为系统保留项
    """
    note_type = note_element.get(f'{{{WORD_NS}}}type')
    return note_type in ('separator', 'continuationSeparator')


# ==============================================================================
# 直接从 ZIP 包读取脚注/尾注的备用方案
# ==============================================================================

def _extract_notes_from_zip(docx_path: str, note_type: str) -> Dict[int, str]:
    """
    直接从 DOCX ZIP 包中读取脚注或尾注
    
    当 python-docx 的 footnotes_part/endnotes_part 不可用时使用此备用方案。
    
    参数:
        docx_path: DOCX 文件路径
        note_type: 'footnotes' 或 'endnotes'
        
    返回:
        Dict[int, str]: {ID: 内容}
    """
    notes = {}
    xml_path = f'word/{note_type}.xml'
    element_tag = 'footnote' if note_type == 'footnotes' else 'endnote'
    ref_tag = 'footnoteRef' if note_type == 'footnotes' else 'endnoteRef'
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            if xml_path not in zf.namelist():
                logger.debug(f"ZIP 包中不存在 {xml_path}")
                return notes
            
            xml_content = zf.read(xml_path)
            root = etree.fromstring(xml_content)
            
            for note_elem in root.findall(f'.//w:{element_tag}', NSMAP):
                # 跳过系统保留项
                if is_system_note(note_elem):
                    continue
                
                note_id_str = note_elem.get(f'{{{WORD_NS}}}id')
                if note_id_str is None:
                    continue
                
                try:
                    note_id = int(note_id_str)
                except ValueError:
                    logger.warning(f"无效的{element_tag} ID: {note_id_str}")
                    continue
                
                content = _extract_note_content(note_elem, ref_tag)
                
                if content:
                    notes[note_id] = content
                    logger.debug(f"从ZIP提取{element_tag} #{note_id}: {content[:50]}...")
            
            logger.info(f"从ZIP提取{note_type}完成 | 共 {len(notes)} 个")
            
    except zipfile.BadZipFile:
        logger.warning(f"文件不是有效的ZIP格式: {docx_path}")
    except Exception as e:
        logger.error(f"从ZIP提取{note_type}失败: {e}", exc_info=True)
    
    return notes


def extract_footnotes_from_docx(doc, docx_path: str = None) -> Dict[int, str]:
    """
    从Word文档中提取所有脚注
    
    优先通过 python-docx 访问，如果失败则使用 ZIP 直接读取备用方案。
    
    参数:
        doc: python-docx Document对象
        docx_path: DOCX 文件路径（用于备用方案）
    
    返回:
        Dict[int, str]: {脚注ID: 脚注内容}
    """
    logger.debug("开始提取Word文档脚注...")
    
    footnotes = {}
    use_fallback = False
    
    try:
        footnotes_part = doc.part.footnotes_part
        if footnotes_part is None:
            logger.info("文档无脚注部分（python-docx返回None）")
            use_fallback = True
        else:
            footnotes_xml = footnotes_part._element
            
            for footnote in footnotes_xml.findall('.//w:footnote', NSMAP):
                # 跳过系统保留项
                if is_system_note(footnote):
                    continue
                
                footnote_id_str = footnote.get(f'{{{WORD_NS}}}id')
                if footnote_id_str is None:
                    continue
                
                try:
                    footnote_id = int(footnote_id_str)
                except ValueError:
                    logger.warning(f"无效的脚注ID: {footnote_id_str}")
                    continue
                
                content = _extract_note_content(footnote, 'footnoteRef')
                
                if content:
                    footnotes[footnote_id] = content
                    logger.debug(f"提取脚注 #{footnote_id}: {content[:50]}...")
            
            logger.info(f"Word脚注提取完成 | 共 {len(footnotes)} 个")
        
    except AttributeError:
        logger.debug("python-docx 不支持 footnotes_part 属性，尝试备用方案")
        use_fallback = True
    except Exception as e:
        logger.error(f"提取Word脚注失败: {e}", exc_info=True)
        use_fallback = True
    
    # 备用方案：直接从 ZIP 包读取
    if use_fallback and docx_path:
        logger.info(f"使用备用方案从ZIP包读取脚注: {docx_path}")
        footnotes = _extract_notes_from_zip(docx_path, 'footnotes')
    
    return footnotes


def extract_endnotes_from_docx(doc, docx_path: str = None) -> Dict[int, str]:
    """
    从Word文档中提取所有尾注
    
    优先通过 python-docx 访问，如果失败则使用 ZIP 直接读取备用方案。
    
    参数:
        doc: python-docx Document对象
        docx_path: DOCX 文件路径（用于备用方案）
    
    返回:
        Dict[int, str]: {尾注ID: 尾注内容}
    """
    logger.debug("开始提取Word文档尾注...")
    
    endnotes = {}
    use_fallback = False
    
    try:
        endnotes_part = doc.part.endnotes_part
        if endnotes_part is None:
            logger.info("文档无尾注部分（python-docx返回None）")
            use_fallback = True
        else:
            endnotes_xml = endnotes_part._element
            
            for endnote in endnotes_xml.findall('.//w:endnote', NSMAP):
                # 跳过系统保留项
                if is_system_note(endnote):
                    continue
                
                endnote_id_str = endnote.get(f'{{{WORD_NS}}}id')
                if endnote_id_str is None:
                    continue
                
                try:
                    endnote_id = int(endnote_id_str)
                except ValueError:
                    logger.warning(f"无效的尾注ID: {endnote_id_str}")
                    continue
                
                content = _extract_note_content(endnote, 'endnoteRef')
                
                if content:
                    endnotes[endnote_id] = content
                    logger.debug(f"提取尾注 #{endnote_id}: {content[:50]}...")
            
            logger.info(f"Word尾注提取完成 | 共 {len(endnotes)} 个")
        
    except AttributeError:
        logger.debug("python-docx 不支持 endnotes_part 属性，尝试备用方案")
        use_fallback = True
    except Exception as e:
        logger.error(f"提取Word尾注失败: {e}", exc_info=True)
        use_fallback = True
    
    # 备用方案：直接从 ZIP 包读取
    if use_fallback and docx_path:
        logger.info(f"使用备用方案从ZIP包读取尾注: {docx_path}")
        endnotes = _extract_notes_from_zip(docx_path, 'endnotes')
    
    return endnotes


def _extract_note_content(note_element, ref_tag: str) -> str:
    """
    从脚注/尾注XML元素中提取文本内容
    
    参数:
        note_element: w:footnote 或 w:endnote XML元素
        ref_tag: 引用标签名 ('footnoteRef' 或 'endnoteRef')，不带命名空间前缀
    
    返回:
        str: 脚注/尾注的文本内容
    """
    texts = []
    
    for para in note_element.findall('.//w:p', NSMAP):
        para_texts = []
        
        for run in para.findall('.//w:r', NSMAP):
            # 跳过包含引用符号的 run
            ref_elem = run.find(f'.//w:{ref_tag}', NSMAP)
            if ref_elem is not None:
                continue
            
            for t in run.findall('.//w:t', NSMAP):
                if t.text:
                    para_texts.append(t.text)
        
        if para_texts:
            texts.append(''.join(para_texts))
    
    return '\n'.join(texts).strip()


def find_footnote_references_in_paragraph(paragraph) -> List[Tuple[int, int]]:
    """
    查找段落中的所有脚注引用
    
    参数:
        paragraph: python-docx Paragraph 对象
        
    返回:
        List[Tuple[int, int]]: [(脚注ID, run索引), ...]
    """
    references = []
    
    for run_idx, run in enumerate(paragraph.runs):
        footnote_refs = run._r.findall('.//w:footnoteReference', NSMAP)
        for ref in footnote_refs:
            footnote_id_str = ref.get(f'{{{WORD_NS}}}id')
            if footnote_id_str:
                try:
                    references.append((int(footnote_id_str), run_idx))
                except ValueError:
                    pass
    
    return references


def find_endnote_references_in_paragraph(paragraph) -> List[Tuple[int, int]]:
    """
    查找段落中的所有尾注引用
    
    参数:
        paragraph: python-docx Paragraph 对象
        
    返回:
        List[Tuple[int, int]]: [(尾注ID, run索引), ...]
    """
    references = []
    
    for run_idx, run in enumerate(paragraph.runs):
        endnote_refs = run._r.findall('.//w:endnoteReference', NSMAP)
        for ref in endnote_refs:
            endnote_id_str = ref.get(f'{{{WORD_NS}}}id')
            if endnote_id_str:
                try:
                    references.append((int(endnote_id_str), run_idx))
                except ValueError:
                    pass
    
    return references


def find_all_note_references_in_paragraph(paragraph) -> List[Tuple[str, int, int]]:
    """
    查找段落中的所有脚注和尾注引用
    
    参数:
        paragraph: python-docx Paragraph 对象
        
    返回:
        List[Tuple[str, int, int]]: [('footnote'/'endnote', 原始ID, run索引), ...]
    """
    references = []
    
    for run_idx, run in enumerate(paragraph.runs):
        # 查找脚注引用
        for ref in run._r.findall('.//w:footnoteReference', NSMAP):
            id_str = ref.get(f'{{{WORD_NS}}}id')
            if id_str:
                try:
                    references.append(('footnote', int(id_str), run_idx))
                except ValueError:
                    pass
        
        # 查找尾注引用
        for ref in run._r.findall('.//w:endnoteReference', NSMAP):
            id_str = ref.get(f'{{{WORD_NS}}}id')
            if id_str:
                try:
                    references.append(('endnote', int(id_str), run_idx))
                except ValueError:
                    pass
    
    return references


def build_footnote_definitions(footnotes: Dict[int, str], id_map: Dict[int, int] = None) -> str:
    """
    生成Markdown格式的脚注定义块
    
    参数:
        footnotes: {Word脚注ID: 脚注内容}
        id_map: {Word脚注ID: Markdown脚注ID}，如果为None则使用顺序编号
    
    返回:
        str: Markdown脚注定义块
    """
    if not footnotes:
        return ""
    
    lines = []
    sorted_ids = sorted(footnotes.keys())
    
    if id_map is None:
        id_map = {old_id: new_id for new_id, old_id in enumerate(sorted_ids, start=1)}
    
    for old_id in sorted_ids:
        new_id = id_map.get(old_id, old_id)
        content = footnotes[old_id]
        content = _format_multiline_content(content)
        lines.append(f"[^{new_id}]: {content}")
    
    return '\n'.join(lines)


def build_endnote_definitions(endnotes: Dict[int, str], id_map: Dict[int, int] = None) -> str:
    """
    生成Markdown格式的尾注定义块
    
    参数:
        endnotes: {Word尾注ID: 尾注内容}
        id_map: {Word尾注ID: Markdown尾注ID}，如果为None则使用顺序编号
    
    返回:
        str: Markdown尾注定义块（使用endnote-前缀）
    """
    if not endnotes:
        return ""
    
    lines = []
    sorted_ids = sorted(endnotes.keys())
    
    if id_map is None:
        id_map = {old_id: new_id for new_id, old_id in enumerate(sorted_ids, start=1)}
    
    for old_id in sorted_ids:
        new_id = id_map.get(old_id, old_id)
        content = endnotes[old_id]
        content = _format_multiline_content(content)
        lines.append(f"[^{ENDNOTE_PREFIX}{new_id}]: {content}")
    
    return '\n'.join(lines)


def _format_multiline_content(content: str) -> str:
    """
    格式化多行内容：第二行起添加4空格缩进
    
    参数:
        content: 原始内容
        
    返回:
        str: 格式化后的内容
    """
    if '\n' in content:
        lines = content.split('\n')
        formatted = lines[0]
        for line in lines[1:]:
            formatted += '\n    ' + line
        return formatted
    return content


def create_footnote_id_mapping(footnotes: Dict[int, str]) -> Dict[int, int]:
    """
    创建Word脚注ID到Markdown脚注ID的映射
    
    按照原始ID排序后重新编号为1, 2, 3...
    
    参数:
        footnotes: {Word脚注ID: 脚注内容}
        
    返回:
        Dict[int, int]: {Word脚注ID: Markdown脚注ID}
    """
    sorted_ids = sorted(footnotes.keys())
    return {old_id: new_id for new_id, old_id in enumerate(sorted_ids, start=1)}


def create_endnote_id_mapping(endnotes: Dict[int, str]) -> Dict[int, int]:
    """
    创建Word尾注ID到Markdown尾注ID的映射
    
    按照原始ID排序后重新编号为1, 2, 3...
    
    参数:
        endnotes: {Word尾注ID: 尾注内容}
        
    返回:
        Dict[int, int]: {Word尾注ID: Markdown尾注ID}
    """
    sorted_ids = sorted(endnotes.keys())
    return {old_id: new_id for new_id, old_id in enumerate(sorted_ids, start=1)}


# ==============================================================================
# 脚注/尾注处理上下文类
# ==============================================================================

class NoteExtractor:
    """
    脚注/尾注提取器
    
    封装脚注和尾注的提取、ID映射和Markdown生成功能。
    支持从 python-docx Document 对象提取，也支持直接从 DOCX 文件路径提取（备用方案）。
    """
    
    def __init__(self, doc, docx_path: str = None):
        """
        初始化提取器
        
        参数:
            doc: python-docx Document 对象
            docx_path: DOCX 文件路径（可选，用于备用方案）
        """
        self.doc = doc
        self.docx_path = docx_path
        
        # 提取脚注和尾注（传入文件路径以支持备用方案）
        self.footnotes = extract_footnotes_from_docx(doc, docx_path)
        self.endnotes = extract_endnotes_from_docx(doc, docx_path)
        
        # 创建ID映射
        self.footnote_id_map = create_footnote_id_mapping(self.footnotes)
        self.endnote_id_map = create_endnote_id_mapping(self.endnotes)
        
        logger.info(f"NoteExtractor初始化 | 脚注: {len(self.footnotes)} 个, 尾注: {len(self.endnotes)} 个")
    
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
    
    def get_footnote_md_id(self, word_id: int) -> Optional[int]:
        """
        获取Word脚注ID对应的Markdown ID
        
        参数:
            word_id: Word脚注ID
            
        返回:
            int: Markdown脚注ID，如果不存在则返回None
        """
        return self.footnote_id_map.get(word_id)
    
    def get_endnote_md_id(self, word_id: int) -> Optional[int]:
        """
        获取Word尾注ID对应的Markdown ID
        
        参数:
            word_id: Word尾注ID
            
        返回:
            int: Markdown尾注ID，如果不存在则返回None
        """
        return self.endnote_id_map.get(word_id)
    
    def get_footnote_reference_text(self, word_id: int) -> str:
        """
        获取脚注的Markdown引用文本
        
        参数:
            word_id: Word脚注ID
            
        返回:
            str: Markdown引用文本（如 [^1]），如果ID不存在则返回空字符串
        """
        md_id = self.get_footnote_md_id(word_id)
        if md_id is None:
            return ''
        return f'[^{md_id}]'
    
    def get_endnote_reference_text(self, word_id: int) -> str:
        """
        获取尾注的Markdown引用文本
        
        参数:
            word_id: Word尾注ID
            
        返回:
            str: Markdown引用文本（如 [^endnote-1]），如果ID不存在则返回空字符串
        """
        md_id = self.get_endnote_md_id(word_id)
        if md_id is None:
            return ''
        return f'[^{ENDNOTE_PREFIX}{md_id}]'
    
    def build_definitions_block(self) -> str:
        """
        生成完整的脚注/尾注定义块
        
        返回:
            str: Markdown定义块（脚注在前，尾注在后）
        """
        blocks = []
        
        if self.has_footnotes:
            footnote_defs = build_footnote_definitions(self.footnotes, self.footnote_id_map)
            blocks.append(footnote_defs)
        
        if self.has_endnotes:
            endnote_defs = build_endnote_definitions(self.endnotes, self.endnote_id_map)
            blocks.append(endnote_defs)
        
        return '\n\n'.join(blocks)
