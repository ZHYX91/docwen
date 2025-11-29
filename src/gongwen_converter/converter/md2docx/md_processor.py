"""
Markdown处理器模块
处理MD文件的正文部分，转换为结构化的段落数据
"""

import logging
import hashlib
from gongwen_converter.utils.number_utils import number_to_chinese, number_to_circled
from gongwen_converter.utils.heading_utils import detect_heading_level
from gongwen_converter.utils.markdown_utils import extract_markdown_tables

# 配置日志
logger = logging.getLogger(__name__)

# 定义符号列表，用于判断标题末尾是否有符号
PUNCTUATION_SET = {'。', '，', '；', '：', '！', '？', '.', ',', ';', ':', '!', '?'}

def process_md_body(md_body):
    """
    处理MD正文部分 - 精确控制标题和正文之间的分段，支持表格
    
    参数:
        md_body: Markdown正文内容
        
    返回:
        list: 处理后的段落列表 [{
            'text': 文本内容,
            'level': 标题级别 (0表示正文),
            'type': 段落类型 ('heading'/'heading_with_content'/'content'/'table'),
            'table_data': 表格数据 (仅type='table'时存在)
        }]
    """
    logger.info("开始处理MD正文...")
    
    # 记录正文初始状态
    body_hash = hashlib.md5(md_body.encode('utf-8')).hexdigest()[:8]
    logger.debug(f"MD正文初始状态 | 长度: {len(md_body)} 字符 | 哈希: {body_hash}")
    
    if not md_body:
        logger.warning("MD正文为空")
        return []
    
    # 提取所有表格
    tables = extract_markdown_tables(md_body)
    logger.info(f"提取到 {len(tables)} 个表格")
    
    # 创建表格行号映射（用于跳过表格行）
    table_line_ranges = {}
    for table_idx, table in enumerate(tables):
        start_line = table['start_line']
        end_line = table['end_line']
        for line_no in range(start_line, end_line):
            table_line_ranges[line_no] = table_idx
        logger.debug(f"表格 {table_idx+1} 占据行 {start_line}-{end_line-1}")
    
    # 初始化各级标题计数器 - 每次处理独立
    heading_levels = [0, 0, 0, 0, 0]  # 一至五级标题计数器
    processed = []  # 处理后的段落列表
    
    # 分割为行，保留空行
    lines = md_body.split('\n')
    logger.debug(f"正文总行数: {len(lines)}")
    
    i = 0
    n = len(lines)
    
    # 状态跟踪变量
    total_paragraphs = 0
    heading_count = 0
    content_count = 0
    combined_count = 0
    
    while i < n:
        line = lines[i].strip()
        
        # 检查当前行是否是表格的一部分
        if i in table_line_ranges:
            table_idx = table_line_ranges[i]
            table = tables[table_idx]
            
            # 如果是表格的第一行，插入表格段落
            if i == table['start_line']:
                processed.append({
                    'type': 'table',
                    'level': 0,
                    'table_data': table
                })
                logger.info(f"在行 {i} 插入表格 {table_idx+1}: {len(table['headers'])}列 x {len(table['rows'])}行")
            
            # 跳过表格行
            i += 1
            continue
        
        # 跳过空行
        if not line:
            i += 1
            continue
            
        # 处理标题行
        if line.startswith('# '):  # 一级标题
            heading_levels[0] += 1
            heading_levels[1:] = [0, 0, 0, 0]  # 重置下级计数器
            text = clean_heading(line[2:])
            
            # 检查标题末尾是否有符号
            if text and text[-1] in PUNCTUATION_SET:
                # 有符号，检查下一行是否是正文（没有空行）
                if i + 1 < n and lines[i + 1].strip() and not lines[i + 1].startswith('#'):
                    processed.append({
                        'text': f"{number_to_chinese(heading_levels[0])}、{text}",
                        'level': 1,
                        'type': 'heading_with_content',
                        'content': lines[i + 1].strip()
                    })
                    i += 1  # 跳过正文行
                    combined_count += 1
                    logger.debug(f"创建组合标题段落 (一级): {text[:20]}... + 正文")
                else:
                    processed.append({
                        'text': f"{number_to_chinese(heading_levels[0])}、{text}",
                        'level': 1,
                        'type': 'heading'
                    })
                    logger.debug(f"创建独立标题段落 (一级): {text[:20]}...")
            else:
                # 没有符号，即使下一行是正文也不合并
                processed.append({
                    'text': f"{number_to_chinese(heading_levels[0])}、{text}",
                    'level': 1,
                    'type': 'heading'
                })
                logger.debug(f"创建独立标题段落 (一级): {text[:20]}...")
                
                # 下一行是正文时，单独处理
                if i + 1 < n and lines[i + 1].strip() and not lines[i + 1].startswith('#'):
                    processed.append({
                        'text': lines[i + 1].strip(),
                        'level': 0,
                        'type': 'content'
                    })
                    i += 1
                    content_count += 1
                    logger.debug(f"标题后添加正文段落: {lines[i + 1].strip()[:20]}...")
                    
        elif line.startswith('## '):  # 二级标题
            heading_levels[1] += 1
            heading_levels[2:] = [0, 0, 0]  # 重置下级计数器
            text = clean_heading(line[3:])
            
            if text and text[-1] in PUNCTUATION_SET:
                if i + 1 < n and lines[i + 1].strip() and not lines[i + 1].startswith('#'):
                    processed.append({
                        'text': f"（{number_to_chinese(heading_levels[1])}）{text}",
                        'level': 2,
                        'type': 'heading_with_content',
                        'content': lines[i + 1].strip()
                    })
                    i += 1
                    combined_count += 1
                    logger.debug(f"创建组合标题段落 (二级): {text[:20]}... + 正文")
                else:
                    processed.append({
                        'text': f"（{number_to_chinese(heading_levels[1])}）{text}",
                        'level': 2,
                        'type': 'heading'
                    })
                    logger.debug(f"创建独立标题段落 (二级): {text[:20]}...")
            else:
                processed.append({
                    'text': f"（{number_to_chinese(heading_levels[1])}）{text}",
                    'level': 2,
                    'type': 'heading'
                })
                logger.debug(f"创建独立标题段落 (二级): {text[:20]}...")
                
                # 下一行是正文时，单独处理
                if i + 1 < n and lines[i + 1].strip() and not lines[i + 1].startswith('#'):
                    processed.append({
                        'text': lines[i + 1].strip(),
                        'level': 0,
                        'type': 'content'
                    })
                    i += 1
                    content_count += 1
                    logger.debug(f"标题后添加正文段落: {lines[i + 1].strip()[:20]}...")
                
        elif line.startswith('### '):  # 三级标题
            heading_levels[2] += 1
            heading_levels[3:] = [0, 0]  # 重置下级计数器
            text = clean_heading(line[4:])
            
            if text and text[-1] in PUNCTUATION_SET:
                if i + 1 < n and lines[i + 1].strip() and not lines[i + 1].startswith('#'):
                    processed.append({
                        'text': f"{heading_levels[2]}. {text}",
                        'level': 3,
                        'type': 'heading_with_content',
                        'content': lines[i + 1].strip()
                    })
                    i += 1
                    combined_count += 1
                    logger.debug(f"创建组合标题段落 (三级): {text[:20]}... + 正文")
                else:
                    processed.append({
                        'text': f"{heading_levels[2]}. {text}",
                        'level': 3,
                        'type': 'heading'
                    })
                    logger.debug(f"创建独立标题段落 (三级): {text[:20]}...")
            else:
                processed.append({
                    'text': f"{heading_levels[2]}. {text}",
                    'level': 3,
                    'type': 'heading'
                })
                logger.debug(f"创建独立标题段落 (三级): {text[:20]}...")
                
                # 下一行是正文时，单独处理
                if i + 1 < n and lines[i + 1].strip() and not lines[i + 1].startswith('#'):
                    processed.append({
                        'text': lines[i + 1].strip(),
                        'level': 0,
                        'type': 'content'
                    })
                    i += 1
                    content_count += 1
                    logger.debug(f"标题后添加正文段落: {lines[i + 1].strip()[:20]}...")
                
        elif line.startswith('#### '):  # 四级标题
            heading_levels[3] += 1
            heading_levels[4] = 0  # 重置下级计数器
            text = clean_heading(line[5:])
            
            if text and text[-1] in PUNCTUATION_SET:
                if i + 1 < n and lines[i + 1].strip() and not lines[i + 1].startswith('#'):
                    processed.append({
                        'text': f"（{heading_levels[3]}）{text}",
                        'level': 4,
                        'type': 'heading_with_content',
                        'content': lines[i + 1].strip()
                    })
                    i += 1
                    combined_count += 1
                    logger.debug(f"创建组合标题段落 (四级): {text[:20]}... + 正文")
                else:
                    processed.append({
                        'text': f"（{heading_levels[3]}）{text}",
                        'level': 4,
                        'type': 'heading'
                    })
                    logger.debug(f"创建独立标题段落 (四级): {text[:20]}...")
            else:
                processed.append({
                    'text': f"（{heading_levels[3]}）{text}",
                    'level': 4,
                    'type': 'heading'
                })
                logger.debug(f"创建独立标题段落 (四级): {text[:20]}...")
                
                # 下一行是正文时，单独处理
                if i + 1 < n and lines[i + 1].strip() and not lines[i + 1].startswith('#'):
                    processed.append({
                        'text': lines[i + 1].strip(),
                        'level': 0,
                        'type': 'content'
                    })
                    i += 1
                    content_count += 1
                    logger.debug(f"标题后添加正文段落: {lines[i + 1].strip()[:20]}...")
                
        elif line.startswith('##### '):  # 五级标题
            heading_levels[4] += 1
            text = clean_heading(line[6:])
            
            if text and text[-1] in PUNCTUATION_SET:
                if i + 1 < n and lines[i + 1].strip() and not lines[i + 1].startswith('#'):
                    processed.append({
                        'text': f"{number_to_circled(heading_levels[4])}{text}",
                        'level': 5,
                        'type': 'heading_with_content',
                        'content': lines[i + 1].strip()
                    })
                    i += 1
                    combined_count += 1
                    logger.debug(f"创建组合标题段落 (五级): {text[:20]}... + 正文")
                else:
                    processed.append({
                        'text': f"{number_to_circled(heading_levels[4])}{text}",
                        'level': 5,
                        'type': 'heading'
                    })
                    logger.debug(f"创建独立标题段落 (五级): {text[:20]}...")
            else:
                processed.append({
                    'text': f"{number_to_circled(heading_levels[4])}{text}",
                    'level': 5,
                    'type': 'heading'
                })
                logger.debug(f"创建独立标题段落 (五级): {text[:20]}...")
                
                # 下一行是正文时，单独处理
                if i + 1 < n and lines[i + 1].strip() and not lines[i + 1].startswith('#'):
                    processed.append({
                        'text': lines[i + 1].strip(),
                        'level': 0,
                        'type': 'content'
                    })
                    i += 1
                    content_count += 1
                    logger.debug(f"标题后添加正文段落: {lines[i + 1].strip()[:20]}...")
                
        else:  # 正文段落
            # 修改点：不再合并连续正文行，每行都作为一个独立段落
            if line.strip():  # 确保不是空行
                processed.append({
                    'text': line.strip(),
                    'level': 0,
                    'type': 'content'
                })
                content_count += 1
                logger.debug(f"添加正文段落: {line.strip()[:50]}...")
                
        i += 1
        total_paragraphs = len(processed)
    
    # 统计处理结果
    heading_count = sum(1 for p in processed if p['type'].startswith('heading'))
    logger.info(f"处理完成 | 总段落数: {total_paragraphs}")
    logger.info(f"段落统计 | 标题: {heading_count} | 正文: {content_count} | 组合段落: {combined_count}")
    logger.info(f"标题级别统计 | 一级: {heading_levels[0]} | 二级: {heading_levels[1]} | 三级: {heading_levels[2]} | 四级: {heading_levels[3]} | 五级: {heading_levels[4]}")
    
    return processed


def clean_heading(text):
    """
    清理标题中的序号（兼容中英文符号混用）
    使用统一的标题处理工具
    
    参数:
        text: 原始标题文本
        
    返回:
        str: 清理后的标题文本
    """
    # 使用标题工具模块的功能
    cleaned, level = detect_heading_level(text)
    logger.debug(f"清理标题: '{text}' -> 级别: {level}, 清理后: '{cleaned}'")
    return cleaned

# 模块测试代码
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(level=logging.DEBUG)
    logger.info("Markdown处理器模块测试")
    
    # 测试MD正文处理
    test_content = """
# 一级标题
这是正文内容

## 二级标题
（一）二级标题带序号
这是二级标题的正文

### 三级标题
1. 三级标题带序号
这是三级标题的正文

#### 四级标题
（1）四级标题带序号
这是四级标题的正文

##### 五级标题
①五级标题带序号
这是五级标题的正文

普通正文段落1
普通正文段落2
    """
    
    # 处理MD正文
    paragraphs = process_md_body(test_content)
    
    # 输出处理结果
    logger.info("\n处理结果:")
    for i, para in enumerate(paragraphs):
        level = para['level']
        prefix = "  " * level if level > 0 else ""
        if para['type'] == 'heading_with_content':
            logger.info(f"{prefix}[{i}] 标题({level}): {para['text']}")
            logger.info(f"{prefix}    -> 正文: {para['content']}")
        else:
            logger.info(f"{prefix}[{i}] {para['type']}({level}): {para['text']}")
    
    logger.info("模块测试完成!")
