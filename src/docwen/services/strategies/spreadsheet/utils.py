"""
表格策略辅助函数模块

提供表格转换策略所需的辅助函数：
- 表格文件预处理（格式转换）

依赖：
- formats.spreadsheet: 表格格式转换
- workspace_manager: 工作空间管理
"""

import os
import logging
from typing import Optional

from docwen.utils.path_utils import generate_output_path

logger = logging.getLogger(__name__)


def _preprocess_table_file(file_path: str, temp_dir: str, cancel_event=None, actual_format: str = None) -> str:
    """
    预处理表格文件：创建输入副本，并将非标准格式（XLS/ET/ODS）转换为XLSX
    
    参数:
        file_path: 原始文件路径
        temp_dir: 临时目录路径，转换后的中间文件将输出到此目录
        cancel_event: 用于取消操作的事件对象
        actual_format: 实际文件格式（可选，如果不提供则自动检测）
        
    返回:
        str: 处理后的文件路径（临时目录中）
            - 所有格式都会先创建 input.{ext} 副本
            - 如果需要转换（XLS/ET/ODS），返回转换后的XLSX路径
            - 如果已是标准格式（XLSX/CSV），返回副本路径
            
    说明:
        - 使用actual_format参数避免重复检测文件格式
        - 所有中间文件都输出到temp_dir，由调用者的上下文管理器统一清理
    """
    # 如果没有提供actual_format，则检测
    if actual_format is None:
        from docwen.utils.file_type_utils import detect_actual_file_format
        actual_format = detect_actual_file_format(file_path)
        logger.debug(f"自动检测表格文件格式: {actual_format}")
    else:
        logger.debug(f"使用传入的文件格式: {actual_format}")
    
    # 步骤1：无论什么格式，都先创建输入副本 input.{ext}
    from docwen.utils.workspace_manager import prepare_input_file
    temp_input = prepare_input_file(file_path, temp_dir, actual_format)
    logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
    
    # 步骤2：如果是标准格式（XLSX/CSV），直接返回副本路径
    if actual_format in ['xlsx', 'csv']:
        logger.debug(f"文件已是标准表格格式({actual_format})，返回副本路径")
        return temp_input
    
    # 步骤3：需要转换的格式，从副本转换为XLSX
    if actual_format in ['xls', 'et', 'ods']:
        logger.info(f"检测到{actual_format.upper()}格式，从副本转换为XLSX: {os.path.basename(temp_input)}")
        
        try:
            from docwen.converter.formats.spreadsheet import (
                office_to_xlsx, ods_to_xlsx
            )
            
            # 生成目标文件路径
            output_filename = generate_output_path(
                file_path,
                section="",
                add_timestamp=True,
                description=f'from{actual_format.capitalize()}',
                file_type='xlsx'
            )
            output_path = os.path.join(temp_dir, os.path.basename(output_filename))
            
            # 根据格式选择转换函数，使用副本作为输入
            if actual_format == 'ods':
                converted_path = ods_to_xlsx(
                    temp_input,  # 使用副本
                    output_path,
                    cancel_event=cancel_event
                )
            else:  # xls 或 et
                converted_path = office_to_xlsx(
                    temp_input,  # 使用副本
                    output_path,
                    actual_format=actual_format,
                    cancel_event=cancel_event
                )
            
            if converted_path:
                logger.info(f"{actual_format.upper()}转XLSX成功: {os.path.basename(converted_path)}")
                return converted_path
            else:
                logger.error("格式转换失败，返回None")
                raise RuntimeError(f"{actual_format.upper()}转XLSX失败")
                
        except Exception as e:
            logger.error(f"{actual_format.upper()}转XLSX失败: {e}")
            raise RuntimeError(f"{actual_format.upper()}转XLSX失败: {e}")
    
    # 步骤4：其他不支持的格式，返回副本路径（尝试直接处理）
    logger.warning(f"不支持的表格格式: {actual_format}，返回副本路径尝试处理")
    return temp_input
