"""
内容插入器模块
将文本框和表格内容插入到文档的正确位置
"""

import logging
from .textbox_processor import extract_textboxes_from_document
from .table_processor import extract_tables_from_document

# 配置日志
logger = logging.getLogger(__name__)

# 硬编码开关：控制是否输出提取的文档（将表格和文本框内容转为普通段落的版本）
OUTPUT_EXTRACTED_DOCX = False  # 设置为False不输出提取文档，True则输出

def inject_special_content(doc):
    """
    将文本框和表格内容插入到文档的paragraphs中
    
    参数:
        doc: Document对象
        
    返回:
        Document: 处理后的Document对象
    """
    logger.info("开始将特殊内容插入到文档段落中")
    
    try:
        # 提取文本框内容
        textbox_contents = extract_textboxes_from_document(doc)
        logger.info(f"提取到 {len(textbox_contents)} 个文本框内容")
        
        # 提取表格内容
        table_contents = extract_tables_from_document(doc)
        logger.info(f"提取到 {len(table_contents)} 个表格内容")
        
        # 合并所有特殊内容
        all_special_contents = textbox_contents + table_contents
        logger.info(f"总共 {len(all_special_contents)} 个特殊内容需要插入")
        
        if not all_special_contents:
            logger.info("未发现需要插入的特殊内容")
            return doc
        
        # 按位置排序内容
        sorted_contents = sort_contents_by_position(all_special_contents)
        
        # 插入内容到文档
        modified_doc = insert_contents_into_document(doc, sorted_contents)
        
        logger.info("特殊内容插入完成")
        return modified_doc
        
    except Exception as e:
        logger.error(f"插入特殊内容失败: {str(e)}", exc_info=True)
        return doc

def sort_contents_by_position(contents):
    """
    按位置信息对内容进行排序
    
    参数:
        contents: 内容列表
        
    返回:
        list: 排序后的内容列表
    """
    try:
        # 为每个内容计算插入位置
        for content in contents:
            content['insert_position'] = calculate_insert_position(content)
        
        # 按插入位置排序，当插入位置相同时使用二级排序
        # 二级排序规则：表格索引 -> 行索引 -> 列索引 -> 段落索引
        sorted_contents = sorted(contents, key=lambda x: (
            x['insert_position'],
            x.get('table_index', 0),
            x.get('row_index', 0),
            x.get('cell_index', 0),
            x.get('para_index', 0)
        ))
        
        logger.debug(f"已对 {len(sorted_contents)} 个内容按位置排序")
        return sorted_contents
        
    except Exception as e:
        logger.error(f"排序内容失败: {str(e)}")
        return contents

def calculate_insert_position(content):
    """
    计算内容的插入位置
    
    参数:
        content: 内容数据
        
    返回:
        int: 插入位置索引（段落索引），-1表示插入到末尾
    """
    try:
        position_info = content.get('position', {})
        position_type = position_info.get('type', 'unknown')
        source_type = content.get('source_type', '')
        
        logger.debug(f"计算插入位置: 类型={position_type}, 来源={source_type}, 位置信息={position_info}")
        
        # 优先使用位置信息中的段落索引
        if position_info and position_info.get('index', -1) >= 0:
            para_index = position_info['index']
            logger.debug(f"使用位置信息中的段落索引: {para_index}")
            return para_index
        
        # 对于表格内容，尝试使用表格索引
        if source_type.startswith('table'):
            table_index = content.get('table_index', -1)
            if table_index >= 0:
                # 表格通常位于特定段落位置，使用表格索引作为相对位置
                logger.debug(f"使用表格索引作为插入位置: {table_index}")
                return table_index
        
        # 对于文本框内容，使用默认位置（插入到文档开头）
        elif source_type.startswith('DrawingML') or source_type.startswith('VML') or source_type.startswith('AlternateContent'):
            logger.debug("使用默认位置0插入文本框内容")
            return 0
        
        # 无法确定位置，插入到末尾
        logger.warning(f"无法确定插入位置，将插入到文档末尾: 类型={position_type}, 来源={source_type}")
        return -1
        
    except Exception as e:
        logger.warning(f"计算插入位置失败: {str(e)}")
        return -1

def insert_contents_into_document(doc, sorted_contents):
    """
    将排序后的内容插入到文档中
    
    参数:
        doc: Document对象
        sorted_contents: 排序后的内容列表
        
    返回:
        Document: 插入内容后的Document对象
    """
    try:
        # 获取所有现有段落
        existing_paragraphs = list(doc.paragraphs)
        total_paragraphs = len(existing_paragraphs)
        
        logger.debug(f"文档当前有 {total_paragraphs} 个段落")
        
        # 先在文档末尾添加一个空段落作为锚点，用于统一使用insert_paragraph_before方法
        anchor_paragraph = doc.add_paragraph("")
        logger.debug("已添加锚点段落用于统一插入操作")
        
        # 记录插入操作
        insert_operations = []
        
        for content in sorted_contents:
            insert_position = content['insert_position']
            source_type = content.get('source_type', 'unknown')
            text_content = content.get('text', '')
            
            # 验证插入位置
            if insert_position < 0:
                # 插入到文档末尾（锚点段落之前）
                logger.debug(f"将在文档末尾（锚点段落之前）插入{source_type}内容")
                insert_operations.append({
                    'position': total_paragraphs,  # 锚点段落的位置
                    'text': text_content,
                    'source_type': source_type,
                    'content': content
                })
            elif insert_position < total_paragraphs:
                # 在指定段落前插入
                logger.debug(f"将在段落 {insert_position+1} 前插入{source_type}内容")
                insert_operations.append({
                    'position': insert_position,
                    'text': text_content,
                    'source_type': source_type,
                    'content': content
                })
            else:
                # 位置超出范围，插入到文档末尾（锚点段落之前）
                logger.warning(f"插入位置 {insert_position} 超出范围，将插入到文档末尾（锚点段落之前）")
                insert_operations.append({
                    'position': total_paragraphs,  # 锚点段落的位置
                    'text': text_content,
                    'source_type': source_type,
                    'content': content
                })
        
        # 按位置降序执行插入，避免索引变化
        # 对于同一位置的多个内容，按表格索引、行索引、列索引、段落索引降序排序
        # 确保表格单元格内容按正确顺序插入（从上到下，从左到右）
        insert_operations.sort(key=lambda x: (
            x['position'],
            x['content'].get('table_index', 0),
            x['content'].get('row_index', 0),
            x['content'].get('cell_index', 0),
            x['content'].get('para_index', 0)
        ), reverse=True)
        
        # 执行插入操作 - 统一使用insert_paragraph_before方法
        for operation in insert_operations:
            insert_position = operation['position']
            text_content = operation['text']
            source_type = operation['source_type']
            content = operation['content']
            
            # 统一使用insert_paragraph_before方法插入段落
            if insert_position < len(doc.paragraphs):
                target_paragraph = doc.paragraphs[insert_position]
                new_para = target_paragraph.insert_paragraph_before(text_content)
                logger.debug(f"在段落 {insert_position+1} 前插入{source_type}内容: '{text_content[:50]}...'")
            else:
                # 如果位置超出当前段落范围，插入到锚点段落之前
                anchor_paragraph.insert_paragraph_before(text_content)
                logger.debug(f"在锚点段落前插入{source_type}内容: '{text_content[:50]}...'")
            
            # 设置段落样式（如果需要）
            if new_para and content.get('style'):
                try:
                    new_para.style = content['style']
                except Exception as style_error:
                    logger.warning(f"设置段落样式失败: {style_error}")
        
        # 删除锚点段落
        if anchor_paragraph._element is not None:
            p = anchor_paragraph._element
            p.getparent().remove(p)
            logger.debug("已删除锚点段落")
        
        logger.info(f"成功插入 {len(insert_operations)} 个特殊内容段落")
        return doc
        
    except Exception as e:
        logger.error(f"插入内容到文档失败: {str(e)}", exc_info=True)
        return doc

def save_extracted_document(doc, original_docx_path, output_dir):
    """
    保存提取版本的文档（将表格和文本框内容转为普通段落的版本）
    
    参数:
        doc: Document对象
        original_docx_path: 原始DOCX文件路径
        output_dir: 输出目录
        
    返回:
        str: 提取文档路径
    """
    try:
        from docwen.utils.path_utils import generate_output_path
        
        # 生成提取文档路径
        extracted_output_path = generate_output_path(
            original_docx_path,
            output_dir=output_dir,
            section="",
            add_timestamp=True,
            description="extracted",
            file_type="docx"
        )
        
        # 保存文档
        doc.save(extracted_output_path)
        logger.info(f"提取文档已保存: {extracted_output_path}")
        
        return extracted_output_path
        
    except Exception as e:
        logger.error(f"保存提取文档失败: {str(e)}")
        return None

def inject_textboxes_only(doc):
    """
    只将文本框内容插入到文档的paragraphs中（不处理表格）
    用于非公文模式，保留表格原始结构
    
    参数:
        doc: Document对象
        
    返回:
        Document: 处理后的Document对象
    """
    logger.info("开始将文本框内容插入到文档段落中（非公文模式）")
    
    try:
        # 只提取文本框内容
        textbox_contents = extract_textboxes_from_document(doc)
        logger.info(f"提取到 {len(textbox_contents)} 个文本框内容")
        
        if not textbox_contents:
            logger.info("未发现需要插入的文本框内容")
            return doc
        
        # 按位置排序内容
        sorted_contents = sort_contents_by_position(textbox_contents)
        
        # 插入内容到文档
        modified_doc = insert_contents_into_document(doc, sorted_contents)
        
        logger.info("文本框内容插入完成")
        return modified_doc
        
    except Exception as e:
        logger.error(f"插入文本框内容失败: {str(e)}", exc_info=True)
        return doc


def process_document_with_special_content(doc, original_docx_path, output_dir, mode='gongwen'):
    """
    处理文档的特殊内容并保存提取版本（根据开关控制）
    
    参数:
        doc: Document对象
        original_docx_path: 原始DOCX文件路径
        output_dir: 输出目录
        mode: 处理模式，'gongwen'（公文模式）或 'simple'（非公文模式）
              - 'gongwen': 表格转段落，文本框转段落（默认）
              - 'simple': 只处理文本框，保留表格结构
        
    返回:
        tuple: (处理后的Document对象, 提取文档路径)
    """
    logger.info(f"开始处理文档的特殊内容（模式: {mode}）")
    
    try:
        # 根据模式选择处理方式
        if mode == 'simple':
            # 非公文模式：只处理文本框，保留表格结构
            logger.info("非公文模式：只处理文本框，保留表格结构")
            modified_doc = inject_textboxes_only(doc)
        else:
            # 公文模式：表格和文本框都转为段落
            logger.info("公文模式：表格和文本框都转为段落")
            modified_doc = inject_special_content(doc)
        
        # 根据开关控制是否保存提取文档
        extracted_path = None
        if OUTPUT_EXTRACTED_DOCX:
            extracted_path = save_extracted_document(modified_doc, original_docx_path, output_dir)
            if extracted_path:
                logger.info(f"提取表格、文本框内容的文档已保存: {extracted_path}")
            else:
                logger.warning("提取表格、文本框内容的文档保存失败")
        else:
            logger.info("提取表格、文本框内容的文档输出已禁用（OUTPUT_EXTRACTED_DOCX=False）")
        
        logger.info("文档特殊内容处理完成")
        return modified_doc, extracted_path
        
    except Exception as e:
        logger.error(f"处理文档表格、文本框内容失败: {str(e)}")
        return doc, None
