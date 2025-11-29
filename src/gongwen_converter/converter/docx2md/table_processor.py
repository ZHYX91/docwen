"""
表格处理器模块
处理DOCX文档中的表格，提取内容并转换为段落格式
"""

import logging
from docx.oxml.ns import qn
from gongwen_converter.utils.docx_utils import extract_format_from_paragraph, NAMESPACES

# 配置日志
logger = logging.getLogger(__name__)

def extract_tables_from_document(doc):
    """
    从文档中提取所有表格内容
    
    参数:
        doc: Document对象
        
    返回:
        list: 表格内容列表，每个元素为 (位置索引, 段落对象, 格式信息)
    """
    logger.info("开始提取文档中的表格内容")
    table_contents = []
    
    try:
        # 遍历所有表格
        for table_idx, table in enumerate(doc.tables):
            logger.debug(f"处理表格 {table_idx + 1}, 行数: {len(table.rows)}, 列数: {len(table.columns)}")
            
            # 提取表格内容
            table_data = extract_table_content(table, table_idx)
            if table_data:
                table_contents.extend(table_data)
        
        logger.info(f"共提取到 {len(table_contents)} 个表格内容段落")
        return table_contents
        
    except Exception as e:
        logger.error(f"提取表格内容失败: {str(e)}", exc_info=True)
        return []

def extract_table_content(table, table_index):
    """
    提取单个表格的内容
    
    参数:
        table: Table对象
        table_index: 表格索引
        
    返回:
        list: 表格内容数据列表
    """
    table_data = []
    
    try:
        # 获取表格在文档中的位置
        position_info = get_table_position(table, table_index)
        
        # 遍历表格的所有行和单元格
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                # 获取单元格的XML元素
                tc = cell._tc
                
                # 检查垂直合并：跳过被合并的单元格
                vMerge = tc.find('.//w:vMerge', namespaces=NAMESPACES)
                if vMerge is not None and vMerge.get(qn('w:val')) == 'continue':
                    # 跳过被垂直合并的单元格
                    logger.debug(f"跳过被垂直合并的单元格: 表{table_index+1}行{row_idx+1}列{cell_idx+1}")
                    continue
                
                # 检查水平合并：跳过被跨越的单元格
                gridSpan = tc.find('.//w:gridSpan', namespaces=NAMESPACES)
                if gridSpan is not None:
                    # 有水平合并，检查是否是被跨越的单元格
                    # 被跨越的单元格通常没有内容或内容为空
                    if cell_idx > 0:
                        # 检查前一个单元格是否有gridSpan，如果有则当前单元格是被跨越的
                        prev_cell = row.cells[cell_idx - 1]
                        prev_tc = prev_cell._tc
                        prev_gridSpan = prev_tc.find('.//w:gridSpan', namespaces=NAMESPACES)
                        if prev_gridSpan is not None:
                            # 当前单元格是被水平跨越的单元格，跳过
                            logger.debug(f"跳过被水平跨越的单元格: 表{table_index+1}行{row_idx+1}列{cell_idx+1}")
                            continue
                
                # 处理单元格中的所有段落，合并为一个字符串
                para_texts = []
                first_para = None
                first_para_fonts = None
                all_runs_data = []
                
                for para_idx, paragraph in enumerate(cell.paragraphs):
                    para_text = paragraph.text.strip()
                    if para_text:
                        # 替换段落内部的换行符为 <br>
                        para_text = para_text.replace('\r\n', '<br>').replace('\n', '<br>').replace('\r', '<br>')
                        para_texts.append(para_text)
                        
                        # 保存第一个段落的格式信息
                        if first_para is None:
                            first_para = paragraph
                            first_para_fonts = extract_format_from_paragraph(paragraph)
                        
                        # 收集所有run的数据
                        for run in paragraph.runs:
                            if run.text.strip():
                                all_runs_data.append({
                                    'text': run.text,
                                    'fonts': {}  # 可以扩展为提取run格式
                                })
                
                # 如果单元格有内容，合并所有段落并创建一条记录
                if para_texts:
                    # 用 <br> 连接多个段落
                    cell_text = '<br>'.join(para_texts)
                    
                    # 转义管道符，避免破坏Markdown表格结构
                    cell_text = cell_text.replace('|', '\\|')
                    
                    # 构建表格数据（只为整个单元格创建一条记录）
                    cell_data = {
                        'text': cell_text,
                        'fonts': first_para_fonts if first_para_fonts else {},
                        'runs': all_runs_data,
                        'position': position_info,
                        'table_index': table_index,
                        'row_index': row_idx,
                        'cell_index': cell_idx,
                        'para_index': 0,  # 单元格级别，不再区分段落
                        'source_type': 'table',
                        'element': first_para._element if first_para else None
                    }
                    
                    table_data.append(cell_data)
                    logger.debug(f"从表格提取单元格: 表{table_index+1}行{row_idx+1}列{cell_idx+1} - '{cell_text[:50]}...'")
        
        return table_data
        
    except Exception as e:
        logger.error(f"提取表格内容失败: {str(e)}")
        return []

def get_table_position(table, table_index):
    """
    获取表格在文档中的位置信息
    
    参数:
        table: Table对象
        table_index: 表格索引
        
    返回:
        dict: 位置信息
    """
    try:
        # 获取表格元素
        table_element = table._element
        
        # 获取表格在文档中的位置
        body = table_element.getparent()
        if body is not None:
            # 获取body的所有子元素
            all_elements = list(body)
            
            # 找到表格元素的位置
            if table_element in all_elements:
                table_pos = all_elements.index(table_element)
                
                # 计算表格之前的段落数量
                paragraph_count_before_table = 0
                for i in range(table_pos):
                    elem = all_elements[i]
                    if elem.tag.endswith('}p'):  # 段落元素
                        paragraph_count_before_table += 1
                
                # 返回表格应该插入的位置（在表格之前的最后一个段落之后）
                # 注意：这里返回的是表格应该被插入的位置，而不是表格本身的位置
                insert_position = paragraph_count_before_table
                logger.debug(f"表格{table_index+1}位置计算: 表格在元素位置{table_pos}, 之前段落数{paragraph_count_before_table}, 插入位置{insert_position}")
                
                return {
                    'type': 'paragraph',
                    'index': insert_position,  # 在表格原本的位置插入
                    'element': table_element,
                    'absolute_position': table_pos
                }
            else:
                logger.warning(f"表格{table_index+1}元素不在body子元素列表中")
                return {'type': 'paragraph', 'index': -1, 'element': None}
        
        logger.warning(f"表格{table_index+1}无法获取父元素body")
        return {'type': 'paragraph', 'index': -1, 'element': None}
        
    except Exception as e:
        logger.warning(f"获取表格位置失败: {str(e)}")
        return {'type': 'paragraph', 'index': -1, 'element': None}

def convert_spreadsheet_to_md(table_data_list):
    """
    将表格数据转换为Markdown格式（用于调试或特定输出）
    
    参数:
        table_data_list: 表格数据列表
        
    返回:
        str: Markdown格式的表格
    """
    if not table_data_list:
        return ""
    
    try:
        # 按表格分组
        tables_dict = {}
        for data in table_data_list:
            table_idx = data['table_index']
            if table_idx not in tables_dict:
                tables_dict[table_idx] = []
            tables_dict[table_idx].append(data)
        
        # 为每个表格生成Markdown
        markdown_tables = []
        for table_idx, table_data in tables_dict.items():
            # 确定表格的行列数
            max_row = max(data['row_index'] for data in table_data)
            max_col = max(data['cell_index'] for data in table_data)
            
            # 创建二维数组存储表格内容
            table_grid = [['' for _ in range(max_col + 1)] for _ in range(max_row + 1)]
            
            # 填充表格内容
            for data in table_data:
                row = data['row_index']
                col = data['cell_index']
                table_grid[row][col] = data['text']
            
            # 生成Markdown表格
            markdown_lines = []
            
            # 表头
            header_line = '| ' + ' | '.join(table_grid[0]) + ' |'
            markdown_lines.append(header_line)
            
            # 分隔线
            separator_line = '| ' + ' | '.join(['---'] * (max_col + 1)) + ' |'
            markdown_lines.append(separator_line)
            
            # 数据行
            for row_idx in range(1, max_row + 1):
                row_line = '| ' + ' | '.join(table_grid[row_idx]) + ' |'
                markdown_lines.append(row_line)
            
            markdown_tables.append('\n'.join(markdown_lines))
        
        return '\n\n'.join(markdown_tables)
        
    except Exception as e:
        logger.error(f"转换表格为Markdown失败: {str(e)}")
        return "表格内容提取失败"
