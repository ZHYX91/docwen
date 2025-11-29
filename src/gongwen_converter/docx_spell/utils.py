"""
docx_spell 工具模块
包含段落处理相关的实用函数
"""

import logging
from gongwen_converter.utils.docx_utils import apply_run_style, apply_paragraph_format

# 配置日志
logger = logging.getLogger(__name__)

def find_run_at_position(paragraph, position):
    """
    在段落中查找包含指定文本位置的run
    
    参数:
        paragraph: 段落对象
        position: 文本位置（从段落开头开始的字符索引）
        
    返回:
        Run对象: 包含指定位置的run
        None: 如果找不到对应的run
        
    详细说明:
        DOCX文档中的段落由多个run组成，每个run有自己的文本和样式。
        此函数用于定位包含特定文本位置的run，以便在该位置添加批注。
    """
    logger.debug(f"在段落中查找位置 {position} 对应的run")
    
    # 检查段落是否为空
    if not paragraph.runs:
        logger.warning("段落中没有run对象")
        return None
        
    current_pos = 0  # 当前累计位置
    
    # 遍历段落中的所有run
    for run in paragraph.runs:
        run_text = run.text
        run_length = len(run_text)
        logger.debug(f"检查run: 文本='{run_text}' (长度={run_length}), 当前累计位置={current_pos}")
        
        # 检查目标位置是否在当前run的范围内
        if current_pos <= position < current_pos + run_length:
            # 计算在run内的相对位置
            relative_pos = position - current_pos
            logger.info(f"找到run: 位置 {position} 在run内相对位置 {relative_pos}, run文本: '{run_text}'")
            return run
            
        # 更新累计位置
        current_pos += run_length
        
    # 如果找不到匹配的run，返回第一个run
    logger.warning(f"未找到包含位置 {position} 的run，返回第一个run")
    return paragraph.runs[0]

def copy_run_formatting(source_run, target_run):
    """
    复制run的所有格式属性
    
    参数:
        source_run: 源run对象
        target_run: 目标run对象
        
    详细说明:
        使用项目已有的 apply_run_style 函数，通过深拷贝 rPr 元素来完整复制所有格式属性。
        这比逐个属性复制更可靠，能确保所有格式（包括样式、字体、颜色等）都被完整保留。
    """
    try:
        # 使用项目已有的 apply_run_style 函数
        # 它会深拷贝整个 rPr 元素，确保所有格式属性都被复制
        source_rPr = source_run._element.rPr
        apply_run_style(target_run, source_rPr)
        logger.debug("成功复制run格式（使用apply_run_style）")
    except Exception as e:
        logger.warning(f"复制run格式时出现警告: {str(e)}")


def copy_paragraph_format(source_para, target_para):
    """
    复制段落的所有格式属性（使用项目通用函数）
    
    参数:
        source_para: 源段落对象
        target_para: 目标段落对象
        
    详细说明:
        使用 docx_utils 中的 apply_paragraph_format 函数来复制段落格式。
        与 copy_run_formatting 类似，复用项目已有的通用工具。
    """
    apply_paragraph_format(target_para, source_para)


def plan_run_splits(paragraph, errors):
    """
    分析段落和错误，规划run的拆分方案（支持多个错误）
    
    参数:
        paragraph: 段落对象
        errors: 错误列表
        
    返回:
        拆分计划列表，格式:
        [
            {
                'original_run': run对象,
                'run_start_pos': run在段落中的起始位置,
                'splits': [
                    {'text': '前部分', 'is_error': False, 'error_index': None},
                    {'text': '错误部分1', 'is_error': True, 'error_index': 0},
                    {'text': '中间部分', 'is_error': False, 'error_index': None},
                    {'text': '错误部分2', 'is_error': True, 'error_index': 1},
                    {'text': '后部分', 'is_error': False, 'error_index': None}
                ]
            },
            ...
        ]
        
    详细说明:
        支持同一段落中的多个错误。每个错误部分都标记了对应的错误索引。
    """
    logger.info(f"开始规划段落的run拆分方案，错误数量: {len(errors)}")
    
    # 创建所有需要拆分的位置列表（错误的开始和结束位置）
    split_positions = set()
    for error in errors:
        split_positions.add(error.start_pos)
        split_positions.add(error.end_pos)
    split_positions = sorted(split_positions)
    
    logger.debug(f"拆分位置: {split_positions}")
    
    # 构建拆分计划
    current_pos = 0
    split_plan = []
    
    for run in paragraph.runs:
        run_length = len(run.text)
        run_end = current_pos + run_length
        original_text = run.text
        
        # 找到这个run内的所有拆分点
        run_splits = [pos for pos in split_positions if current_pos < pos < run_end]
        
        if not run_splits:
            # 这个run不需要拆分
            # 检查整个run是否在某个错误范围内
            error_index = None
            for idx, error in enumerate(errors):
                if error.start_pos <= current_pos and run_end <= error.end_pos:
                    error_index = idx
                    break
            
            split_plan.append({
                'original_run': run,
                'run_start_pos': current_pos,
                'splits': [{
                    'text': run.text,
                    'is_error': error_index is not None,
                    'error_index': error_index
                }]
            })
        else:
            # 需要拆分这个run
            splits = []
            last_pos = 0  # run内的相对位置
            
            for split_pos in run_splits:
                relative_pos = split_pos - current_pos
                
                # 添加从last_pos到split_pos的部分
                if relative_pos > last_pos:
                    text_part = original_text[last_pos:relative_pos]
                    # 检查这部分是否在错误范围内
                    part_start = current_pos + last_pos
                    part_end = current_pos + relative_pos
                    error_index = None
                    for idx, error in enumerate(errors):
                        if error.start_pos <= part_start and part_end <= error.end_pos:
                            error_index = idx
                            break
                    
                    splits.append({
                        'text': text_part,
                        'is_error': error_index is not None,
                        'error_index': error_index
                    })
                
                last_pos = relative_pos
            
            # 添加最后一部分
            if last_pos < run_length:
                text_part = original_text[last_pos:]
                # 检查这部分是否在错误范围内
                part_start = current_pos + last_pos
                part_end = run_end
                error_index = None
                for idx, error in enumerate(errors):
                    if error.start_pos <= part_start and part_end <= error.end_pos:
                        error_index = idx
                        break
                
                splits.append({
                    'text': text_part,
                    'is_error': error_index is not None,
                    'error_index': error_index
                })
            
            split_plan.append({
                'original_run': run,
                'run_start_pos': current_pos,
                'splits': splits
            })
            logger.debug(f"拆分run '{original_text}' 为 {len(splits)} 部分")
        
        current_pos = run_end
    
    logger.info(f"拆分规划完成，共 {len(split_plan)} 个run")
    return split_plan


def rebuild_paragraph_with_splits(old_paragraph, split_plan, doc):
    """
    根据拆分计划重建段落
    
    参数:
        old_paragraph: 原段落对象
        split_plan: 拆分计划（由plan_run_splits生成）
        doc: Document对象
        
    返回:
        tuple: (新段落对象, 错误run字典)
            错误run字典格式: {error_index: run对象}
        
    详细说明:
        1. 在文档中创建新段落（在原段落位置）
        2. 复制段落格式
        3. 按计划添加所有run
        4. 记录所有错误run（按错误索引）
        5. 删除原段落
    """
    logger.info("开始重建段落")
    
    try:
        # 1. 获取原段落在文档中的位置
        parent = old_paragraph._element.getparent()
        old_para_index = parent.index(old_paragraph._element)
        logger.debug(f"原段落在文档中的索引: {old_para_index}")
        
        # 2. 创建新段落（在原段落后面）
        new_paragraph = doc.add_paragraph()
        
        # 3. 复制段落格式
        copy_paragraph_format(old_paragraph, new_paragraph)
        
        # 4. 按拆分计划添加run，记录所有错误run
        error_runs = {}  # {error_index: run对象}
        for plan_item in split_plan:
            original_run = plan_item['original_run']
            splits = plan_item['splits']
            
            for split in splits:
                # 创建新run
                new_run = new_paragraph.add_run(split['text'])
                # 复制原run的格式
                copy_run_formatting(original_run, new_run)
                
                # 记录错误run
                if split['is_error'] and split['error_index'] is not None:
                    error_runs[split['error_index']] = new_run
                    logger.debug(f"标记错误run[{split['error_index']}]: '{split['text']}'")
        
        # 5. 将新段落移动到原段落位置
        parent.insert(old_para_index, new_paragraph._element)
        logger.debug("新段落已插入到原位置")
        
        # 6. 删除原段落
        parent.remove(old_paragraph._element)
        logger.debug("原段落已删除")
        
        logger.info(f"段落重建完成，共标记 {len(error_runs)} 个错误run")
        return new_paragraph, error_runs
        
    except Exception as e:
        logger.error(f"重建段落失败: {str(e)}", exc_info=True)
        raise
