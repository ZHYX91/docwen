"""
文本框处理器模块
处理DOCX文档中的文本框，提取内容并转换为段落格式
"""

import logging
from gongwen_converter.utils.docx_utils import extract_format_from_paragraph, NAMESPACES

# 配置日志
logger = logging.getLogger(__name__)

def extract_textboxes_from_document(doc):
    """
    从文档中提取所有文本框内容
    
    参数:
        doc: Document对象
        
    返回:
        list: 文本框内容列表，每个元素为字典包含文本、格式和位置信息
    """
    logger.info("开始提取文档中的文本框内容")
    textbox_contents = []
    
    try:
        # 获取文档的XML body元素
        body = doc.element.body
        
        # 用于跟踪已经处理过的段落元素，避免重复提取
        processed_para_elements = set()
        # 用于跟踪已经处理过的文本内容，进行内容级别的去重
        processed_text_contents = set()
        
        # 按优先级处理不同类型的文本框
        # 1. 首先处理DrawingML文本框 (wps:txbx) - 最高优先级
        drawingml_textboxes = extract_drawingml_textboxes(body, doc, processed_para_elements, processed_text_contents)
        textbox_contents.extend(drawingml_textboxes)
        
        # 2. 然后处理AlternateContent结构中的文本框
        alternate_content_textboxes = extract_alternate_content_textboxes(body, doc, processed_para_elements, processed_text_contents)
        textbox_contents.extend(alternate_content_textboxes)
        
        # 3. 最后处理VML文本框 (v:textbox) - 最低优先级
        vml_textboxes = extract_vml_textboxes(body, doc, processed_para_elements, processed_text_contents)
        textbox_contents.extend(vml_textboxes)
        
        logger.info(f"共提取到 {len(textbox_contents)} 个文本框内容段落")
        return textbox_contents
        
    except Exception as e:
        logger.error(f"提取文本框内容失败: {str(e)}", exc_info=True)
        return []

def extract_drawingml_textboxes(body, doc, processed_para_elements=None, processed_text_contents=None):
    """
    提取DrawingML格式的文本框内容
    
    参数:
        body: 文档body元素
        doc: Document对象
        processed_para_elements: 已处理的段落元素集合，用于去重
        processed_text_contents: 已处理的文本内容集合，用于内容级别去重
        
    返回:
        list: DrawingML文本框内容列表
    """
    textbox_contents = []
    
    try:
        # 查找所有DrawingML文本框元素 - 使用findall替代xpath
        drawingml_textboxes = body.findall('.//wps:txbx', NAMESPACES)
        logger.debug(f"找到 {len(drawingml_textboxes)} 个DrawingML文本框")
        
        for textbox_idx, textbox in enumerate(drawingml_textboxes):
            # 获取文本框的位置信息
            position_info = get_textbox_position(textbox, textbox_idx, 'DrawingML')
            
            # 提取文本框中的段落 - 使用findall替代xpath
            paragraphs = textbox.findall('.//w:p', NAMESPACES)
            logger.debug(f"DrawingML文本框 {textbox_idx+1} 包含 {len(paragraphs)} 个段落")
            
            for para_idx, para_element in enumerate(paragraphs):
                # 检查段落元素是否已经处理过
                if processed_para_elements is not None:
                    para_id = id(para_element)
                    if para_id in processed_para_elements:
                        logger.debug(f"跳过已处理的DrawingML段落: 文本框{textbox_idx+1}段落{para_idx+1}")
                        continue
                    processed_para_elements.add(para_id)
                
                # 创建Paragraph对象来获取文本和格式 - 使用doc.part替代body
                from docx.text.paragraph import Paragraph
                paragraph = Paragraph(para_element, doc.part)
                
                para_text = paragraph.text.strip()
                if para_text:
                    # 检查文本内容是否已经处理过（内容级别去重）
                    if processed_text_contents is not None:
                        text_key = f"{para_text}_{position_info.get('index', -1)}"
                        if text_key in processed_text_contents:
                            logger.debug(f"跳过已处理的DrawingML文本内容: 文本框{textbox_idx+1}段落{para_idx+1} - '{para_text[:50]}...'")
                            continue
                        processed_text_contents.add(text_key)
                    
                    # 提取段落格式
                    para_fonts = extract_format_from_paragraph(paragraph)
                    
                    # 提取run级别的格式
                    runs_data = []
                    for run in paragraph.runs:
                        if run.text.strip():
                            runs_data.append({
                                'text': run.text,
                                'fonts': {}  # 可以扩展为提取run格式
                            })
                    
                    # 构建文本框数据
                    textbox_data = {
                        'text': para_text,
                        'fonts': para_fonts,
                        'runs': runs_data,
                        'position': position_info,
                        'textbox_index': textbox_idx,
                        'para_index': para_idx,
                        'source_type': 'DrawingML_textbox',
                        'element': para_element
                    }
                    
                    textbox_contents.append(textbox_data)
                    logger.debug(f"从DrawingML文本框提取段落: 文本框{textbox_idx+1}段落{para_idx+1} - '{para_text[:50]}...'")
        
        return textbox_contents
        
    except Exception as e:
        logger.error(f"提取DrawingML文本框内容失败: {str(e)}")
        return []

def extract_vml_textboxes(body, doc, processed_para_elements=None, processed_text_contents=None):
    """
    提取VML格式的文本框内容
    
    参数:
        body: 文档body元素
        doc: Document对象
        processed_para_elements: 已处理的段落元素集合，用于去重
        processed_text_contents: 已处理的文本内容集合，用于内容级别去重
        
    返回:
        list: VML文本框内容列表
    """
    textbox_contents = []
    
    try:
        # 查找所有VML文本框元素 - 使用findall替代xpath
        vml_textboxes = body.findall('.//v:textbox', NAMESPACES)
        logger.debug(f"找到 {len(vml_textboxes)} 个VML文本框")
        
        for textbox_idx, textbox in enumerate(vml_textboxes):
            # 获取文本框的位置信息
            position_info = get_textbox_position(textbox, textbox_idx, 'VML')
            
            # 提取文本框中的段落 - 使用findall替代xpath
            paragraphs = textbox.findall('.//w:p', NAMESPACES)
            logger.debug(f"VML文本框 {textbox_idx+1} 包含 {len(paragraphs)} 个段落")
            
            for para_idx, para_element in enumerate(paragraphs):
                # 检查段落元素是否已经处理过
                if processed_para_elements is not None:
                    para_id = id(para_element)
                    if para_id in processed_para_elements:
                        logger.debug(f"跳过已处理的VML段落: 文本框{textbox_idx+1}段落{para_idx+1}")
                        continue
                    processed_para_elements.add(para_id)
                
                # 创建Paragraph对象来获取文本和格式 - 使用doc.part替代body
                from docx.text.paragraph import Paragraph
                paragraph = Paragraph(para_element, doc.part)
                
                para_text = paragraph.text.strip()
                if para_text:
                    # 检查文本内容是否已经处理过（内容级别去重）
                    if processed_text_contents is not None:
                        text_key = f"{para_text}_{position_info.get('index', -1)}"
                        if text_key in processed_text_contents:
                            logger.debug(f"跳过已处理的VML文本内容: 文本框{textbox_idx+1}段落{para_idx+1} - '{para_text[:50]}...'")
                            continue
                        processed_text_contents.add(text_key)
                    
                    # 提取段落格式
                    para_fonts = extract_format_from_paragraph(paragraph)
                    
                    # 提取run级别的格式
                    runs_data = []
                    for run in paragraph.runs:
                        if run.text.strip():
                            runs_data.append({
                                'text': run.text,
                                'fonts': {}  # 可以扩展为提取run格式
                            })
                    
                    # 构建文本框数据
                    textbox_data = {
                        'text': para_text,
                        'fonts': para_fonts,
                        'runs': runs_data,
                        'position': position_info,
                        'textbox_index': textbox_idx,
                        'para_index': para_idx,
                        'source_type': 'VML_textbox',
                        'element': para_element
                    }
                    
                    textbox_contents.append(textbox_data)
                    logger.debug(f"从VML文本框提取段落: 文本框{textbox_idx+1}段落{para_idx+1} - '{para_text[:50]}...'")
        
        return textbox_contents
        
    except Exception as e:
        logger.error(f"提取VML文本框内容失败: {str(e)}")
        return []

def extract_alternate_content_textboxes(body, doc, processed_para_elements=None, processed_text_contents=None):
    """
    提取AlternateContent结构中的文本框内容
    
    AlternateContent结构包含：
    - mc:Choice: DrawingML格式 (wps:txbx) - Word 2010+
    - mc:Fallback: VML格式 (v:textbox) - Word 2007及更早版本兼容
    
    两者内容相同，只处理Choice部分，但需要记录Fallback中的元素ID以防止
    extract_vml_textboxes重复提取
    
    参数:
        body: 文档body元素
        doc: Document对象
        processed_para_elements: 已处理的段落元素集合，用于去重
        processed_text_contents: 已处理的文本内容集合，用于内容级别去重
        
    返回:
        list: AlternateContent文本框内容列表
    """
    textbox_contents = []
    
    try:
        # 查找所有AlternateContent元素 - 使用findall替代xpath
        alt_content_nodes = body.findall('.//mc:AlternateContent', NAMESPACES)
        logger.debug(f"找到 {len(alt_content_nodes)} 个AlternateContent元素")
        
        textbox_idx = 0
        
        for alt_content in alt_content_nodes:
            # 处理Choice部分 (wps:txbx) - 使用findall替代xpath
            choice_textboxes = alt_content.findall('.//mc:Choice//wps:txbx', NAMESPACES)
            for textbox in choice_textboxes:
                position_info = get_textbox_position(textbox, textbox_idx, 'AlternateContent_Choice')
                
                paragraphs = textbox.findall('.//w:p', NAMESPACES)
                logger.debug(f"AlternateContent Choice文本框 {textbox_idx+1} 包含 {len(paragraphs)} 个段落")
                
                for para_idx, para_element in enumerate(paragraphs):
                    # 检查段落元素是否已经处理过
                    if processed_para_elements is not None:
                        para_id = id(para_element)
                        if para_id in processed_para_elements:
                            logger.debug(f"跳过已处理的AlternateContent Choice段落: 文本框{textbox_idx+1}段落{para_idx+1}")
                            continue
                        processed_para_elements.add(para_id)
                    
                    from docx.text.paragraph import Paragraph
                    paragraph = Paragraph(para_element, doc.part)
                    
                    para_text = paragraph.text.strip()
                    if para_text:
                        # 检查文本内容是否已经处理过（内容级别去重）
                        if processed_text_contents is not None:
                            text_key = f"{para_text}_{position_info.get('index', -1)}"
                            if text_key in processed_text_contents:
                                logger.debug(f"跳过已处理的AlternateContent Choice文本内容: 文本框{textbox_idx+1}段落{para_idx+1} - '{para_text[:50]}...'")
                                continue
                            processed_text_contents.add(text_key)
                        
                        para_fonts = extract_format_from_paragraph(paragraph)
                        
                        runs_data = []
                        for run in paragraph.runs:
                            if run.text.strip():
                                runs_data.append({
                                    'text': run.text,
                                    'fonts': {}
                                })
                        
                        textbox_data = {
                            'text': para_text,
                            'fonts': para_fonts,
                            'runs': runs_data,
                            'position': position_info,
                            'textbox_index': textbox_idx,
                            'para_index': para_idx,
                            'source_type': 'AlternateContent_Choice_textbox',
                            'element': para_element
                        }
                        
                        textbox_contents.append(textbox_data)
                        logger.debug(f"从AlternateContent Choice文本框提取段落: 文本框{textbox_idx+1}段落{para_idx+1} - '{para_text[:50]}...'")
                
                textbox_idx += 1
            
            # 处理Fallback部分：不提取内容，但记录段落元素ID以防止extract_vml_textboxes重复提取
            # Fallback (v:textbox) 是为了兼容Word 2007及更早版本，内容与Choice相同
            fallback_textboxes = alt_content.findall('.//mc:Fallback//v:textbox', NAMESPACES)
            for textbox in fallback_textboxes:
                paragraphs = textbox.findall('.//w:p', NAMESPACES)
                logger.debug(f"标记AlternateContent Fallback文本框中的 {len(paragraphs)} 个段落（不提取，仅用于去重）")
                
                for para_element in paragraphs:
                    # 只记录段落元素ID，不提取内容
                    if processed_para_elements is not None:
                        para_id = id(para_element)
                        processed_para_elements.add(para_id)
        
        return textbox_contents
        
    except Exception as e:
        logger.error(f"提取AlternateContent文本框内容失败: {str(e)}")
        return []

def get_textbox_position(textbox_element, textbox_index, textbox_type):
    """
    获取文本框在文档中的位置信息
    
    参数:
        textbox_element: 文本框XML元素
        textbox_index: 文本框索引
        textbox_type: 文本框类型
        
    返回:
        dict: 位置信息
    """
    try:
        # 获取文本框的父元素（通常是w:p或w:r）
        parent = textbox_element.getparent()
        
        # 向上查找包含文本框的段落
        paragraph_element = None
        current = parent
        while current is not None:
            if current.tag.endswith('}p'):  # 段落元素
                paragraph_element = current
                break
            current = current.getparent()
        
        if paragraph_element is not None:
            # 获取body的所有子元素
            body = paragraph_element.getparent()
            all_elements = list(body)
            
            # 找到段落元素的位置
            if paragraph_element in all_elements:
                para_pos = all_elements.index(paragraph_element)
                
                # 返回文本框应该插入的位置（在包含文本框的段落之后）
                insert_position = para_pos + 1
                logger.debug(f"文本框{textbox_index+1}位置计算: 类型={textbox_type}, 段落位置{para_pos}, 插入位置{insert_position}")
                
                return {
                    'type': 'paragraph',
                    'index': insert_position,  # 在包含文本框的段落之后插入
                    'element': paragraph_element,
                    'absolute_position': para_pos,
                    'textbox_type': textbox_type
                }
            else:
                logger.warning(f"文本框{textbox_index+1}段落元素不在body子元素列表中")
                return {'type': 'paragraph', 'index': -1, 'element': None, 'textbox_type': textbox_type}
        
        logger.warning(f"文本框{textbox_index+1}无法找到包含段落")
        return {'type': 'paragraph', 'index': -1, 'element': None, 'textbox_type': textbox_type}
        
    except Exception as e:
        logger.warning(f"获取文本框位置失败: {str(e)}")
        return {'type': 'paragraph', 'index': -1, 'element': None, 'textbox_type': textbox_type}
