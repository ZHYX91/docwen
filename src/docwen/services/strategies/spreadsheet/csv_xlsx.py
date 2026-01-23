"""
CSV与XLSX互转策略模块

提供CSV和XLSX格式之间的转换策略。

依赖：
- pandas: 数据处理
- table_converters: 表格转换工具
"""

import os
import re
import logging
import tempfile
import shutil
from typing import Dict, Any, Callable, Optional

from ..base_strategy import BaseStrategy
from docwen.services.result import ConversionResult
from .. import register_conversion
from docwen.converter.formats.spreadsheet import xlsx_to_csv
from docwen.utils.path_utils import generate_output_path
from docwen.i18n import t

logger = logging.getLogger(__name__)


@register_conversion('csv', 'xlsx')
class CsvToXlsxStrategy(BaseStrategy):
    """
    将CSV文件转换为XLSX文件的策略
    
    转换说明：
    - CSV文件将被读取并保存为XLSX格式
    - 输出文件使用默认工作表名 "Sheet1"
    - 文件名标记为 "fromCsv"
    - 支持扩展名不匹配的文件
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行CSV到XLSX的转换
        
        Args:
            file_path: 输入的CSV文件路径
            options: 转换选项字典，包含：
                - actual_format: (可选) 文件的真实格式
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        try:
            if progress_callback:
                progress_callback(t('conversion.progress.preparing'))
            
            options = options or {}
            actual_format = options.get('actual_format', 'csv')
            cancel_event = options.get('cancel_event')
            output_dir = os.path.dirname(file_path)
            
            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：创建输入副本 input.csv
                if progress_callback:
                    progress_callback(t('conversion.progress.preparing_files'))
                
                from docwen.utils.workspace_manager import prepare_input_file
                temp_input = prepare_input_file(file_path, temp_dir, actual_format)
                logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                # 步骤2：读取CSV并转换
                if progress_callback:
                    progress_callback(t('conversion.progress.converting_to_format', format='XLSX'))
                
                import pandas as pd
                df = pd.read_csv(temp_input, header=None, keep_default_na=False)
                logger.debug(f"CSV 文件读取成功，数据形状: {df.shape}")
                
                # 步骤3：生成输出文件名
                output_filename = os.path.basename(
                    generate_output_path(
                        file_path,
                        section="",
                        add_timestamp=True,
                        description="fromCsv",
                        file_type="xlsx"
                    )
                )
                
                # 步骤4：保存到临时目录
                temp_output = os.path.join(temp_dir, output_filename)
                with pd.ExcelWriter(temp_output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
                logger.debug(f"XLSX文件已生成: {temp_output}")
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                # 步骤5：移动到目标位置
                final_output = os.path.join(output_dir, output_filename)
                shutil.move(temp_output, final_output)
                logger.info(f"CSV转XLSX成功: {final_output}")
            
            return ConversionResult(
                success=True,
                output_path=final_output,
                message=t('conversion.messages.conversion_to_format_success', format='XLSX')
            )
            
        except Exception as e:
            logger.error(f"CSV转XLSX失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"转换失败: {e}",
                error=e
            )


@register_conversion('xlsx', 'csv')
class XlsxToCsvStrategy(BaseStrategy):
    """
    将XLSX文件转换为CSV文件的策略
    
    转换说明：
    - 每个工作表将生成一个独立的CSV文件
    - 工作表名中的空格将被替换为下划线
    - 文件名包含工作表名作为section，标记为 "fromXlsx"
    - 所有CSV文件使用相同的时间戳
    - 支持扩展名不匹配的文件
    - 所有CSV文件输出到同一个子文件夹
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行XLSX到CSV的转换
        
        Args:
            file_path: 输入的XLSX文件路径
            options: 转换选项字典，包含：
                - actual_format: (可选) 文件的真实格式
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
            
        Note:
            output_path 返回子文件夹中第一个CSV文件的路径
        """
        try:
            if progress_callback:
                progress_callback(t('conversion.progress.preparing'))
            
            options = options or {}
            actual_format = options.get('actual_format', 'xlsx')
            cancel_event = options.get('cancel_event')
            
            from docwen.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用临时目录管理输出
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：创建输入副本 input.xlsx
                if progress_callback:
                    progress_callback(t('conversion.progress.preparing_files'))
                
                from docwen.utils.workspace_manager import prepare_input_file
                temp_input = prepare_input_file(file_path, temp_dir, actual_format)
                logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                # 步骤2：生成统一basename和时间戳描述部分
                base_path = generate_output_path(
                    file_path,
                    section="",
                    add_timestamp=True,
                    description="fromXlsx",
                    file_type="csv"
                )
                basename = os.path.splitext(os.path.basename(base_path))[0]
                logger.debug(f"统一basename: {basename}")
                
                # 提取原始文件名（不含扩展名、时间戳、描述）
                original_file_basename = os.path.splitext(os.path.basename(file_path))[0]
                # 移除可能存在的旧时间戳和描述
                timestamp_pattern = r'(_\d{8}_\d{6})(?:.*)?$'
                match = re.search(timestamp_pattern, original_file_basename)
                if match:
                    original_file_basename = original_file_basename[:match.start()]
                logger.debug(f"原始文件basename: {original_file_basename}")
                
                # 从basename中提取时间戳和描述部分
                parts = basename.split('_')
                unified_timestamp_desc = ""
                if len(parts) >= 3:
                    timestamp_idx = None
                    for i, part in enumerate(parts):
                        if len(part) == 8 and part.isdigit():
                            timestamp_idx = i
                            break
                    
                    if timestamp_idx is not None:
                        unified_timestamp_desc = '_'.join(parts[timestamp_idx:])
                
                if not unified_timestamp_desc:
                    unified_timestamp_desc = '_'.join(parts[1:]) if len(parts) > 1 else parts[0]
                
                logger.debug(f"统一时间戳描述: {unified_timestamp_desc}")
                
                # 步骤3：创建临时子文件夹
                temp_output_folder = os.path.join(temp_dir, basename)
                os.makedirs(temp_output_folder, exist_ok=True)
                logger.debug(f"创建临时子文件夹: {temp_output_folder}")
                
                # 步骤4：转换副本，输出到临时子文件夹
                if progress_callback:
                    progress_callback(t('conversion.progress.converting_to_format', format='CSV'))
                
                csv_files = xlsx_to_csv(
                    temp_input,
                    actual_format, 
                    output_dir=temp_output_folder,
                    original_basename=original_file_basename,
                    unified_timestamp_desc=unified_timestamp_desc
                )
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                if not csv_files:
                    return ConversionResult(
                        success=False,
                        message=t('conversion.messages.no_csv_generated')
                    )
                
                logger.info(f"已生成 {len(csv_files)} 个CSV文件到临时子文件夹")
                
                # 步骤5：移动整个文件夹到输出目录
                final_folder = os.path.join(output_dir, basename)
                
                if os.path.exists(final_folder):
                    shutil.rmtree(final_folder)
                    logger.debug(f"已删除现有文件夹: {final_folder}")
                
                shutil.move(temp_output_folder, final_folder)
                logger.info(f"已移动文件夹到: {final_folder}")
                
                # 准备返回路径（第一个CSV文件的完整路径）
                first_csv_name = os.path.basename(csv_files[0])
                output_path = os.path.join(final_folder, first_csv_name)
            
            if progress_callback:
                progress_callback(t('conversion.progress.csv_completed', count=len(csv_files)))
            
            return ConversionResult(
                success=True,
                output_path=output_path,
                message=t('conversion.messages.conversion_to_format_success', format='CSV')
            )
            
        except Exception as e:
            logger.error(f"XLSX转CSV失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"转换失败: {e}",
                error=e
            )
