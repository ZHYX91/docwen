"""
表格格式互转策略模块

提供表格格式之间的智能转换策略工厂。

依赖：
- smart_converter: 智能转换链
- formats.spreadsheet: 表格格式转换
"""

import os
import logging
import tempfile
import shutil
from typing import Dict, Any, Callable, Optional

from ..base_strategy import BaseStrategy
from docwen.services.result import ConversionResult
from .. import register_conversion
from docwen.i18n import t

logger = logging.getLogger(__name__)


def _create_spreadsheet_conversion_strategy(source_fmt: str, target_fmt: str):
    """
    策略工厂：动态创建表格格式转换策略
    
    使用智能转换链自动处理单步或多步转换
    
    参数:
        source_fmt: 源格式（如 'xls', 'ods'）
        target_fmt: 目标格式（如 'xlsx', 'ods'）
    
    返回:
        动态生成的策略类
    """
    
    @register_conversion(source_fmt, target_fmt)
    class DynamicSpreadsheetConversionStrategy(BaseStrategy):
        """动态生成的表格转换策略"""
        
        def execute(
            self,
            file_path: str,
            options: Optional[Dict[str, Any]] = None,
            progress_callback: Optional[Callable[[str], None]] = None
        ) -> ConversionResult:
            """
            执行表格格式转换
            
            说明:
                从options中提取actual_format参数并传递给SmartConverter，
                确保即使文件扩展名被修改也能正确转换
            """
            try:
                options = options or {}
                cancel_event = options.get('cancel_event')
                preferred_software = options.get('preferred_software')
                actual_format = options.get('actual_format')
                
                # 如果没有提供actual_format，从文件扩展名推断（降级方案）
                if not actual_format:
                    _, ext = os.path.splitext(file_path)
                    actual_format = ext.lower().lstrip('.')
                    logger.warning(f"未提供actual_format，从文件名推断: {actual_format}")
                
                logger.debug(f"表格转换策略: {source_fmt}→{target_fmt}, 真实格式: {actual_format}")
                
                # 使用智能转换链
                from docwen.converter.smart_converter import SmartConverter, OfficeSoftwareNotFoundError
                from docwen.utils.workspace_manager import get_output_directory
                
                converter = SmartConverter()
                output_dir = get_output_directory(file_path)
                
                # 使用临时目录管理中间文件
                with tempfile.TemporaryDirectory() as temp_dir:
                    # 调用SmartConverter，输出到临时目录
                    result_path = converter.convert(
                        input_path=file_path,
                        target_format=target_fmt,
                        category='spreadsheet',
                        actual_format=actual_format,
                        output_dir=temp_dir,
                        cancel_event=cancel_event,
                        progress_callback=progress_callback,
                        preferred_software=preferred_software
                    )
                    
                    if not result_path or not os.path.exists(result_path):
                        return ConversionResult(
                            success=False,
                            message=t('conversion.messages.conversion_to_format_failed', format=target_fmt.upper())
                        )
                    
                    # 检查是否是CSV转换（result_path在子文件夹中）
                    if target_fmt == 'csv':
                        # CSV转换：result_path是子文件夹内的文件
                        csv_subfolder = os.path.dirname(result_path)
                        subfolder_name = os.path.basename(csv_subfolder)
                        
                        # 移动整个CSV子文件夹到输出目录
                        final_output_folder = os.path.join(output_dir, subfolder_name)
                        
                        if os.path.exists(final_output_folder):
                            shutil.rmtree(final_output_folder)
                            logger.debug(f"已删除现有文件夹: {final_output_folder}")
                        
                        shutil.move(csv_subfolder, final_output_folder)
                        logger.info(f"CSV文件夹已移动: {subfolder_name}")
                        
                        final_output_path = os.path.join(final_output_folder, os.path.basename(result_path))
                        
                        # 处理中间文件（如果需要保留）
                        should_keep = self._should_keep_intermediates()
                        if should_keep:
                            logger.info("检查是否有中间文件需要保留")
                            for filename in os.listdir(temp_dir):
                                if filename.startswith('input.') or filename == subfolder_name:
                                    continue
                                src = os.path.join(temp_dir, filename)
                                if os.path.isfile(src):
                                    dst = os.path.join(output_dir, filename)
                                    shutil.move(src, dst)
                                    logger.info(f"保留中间文件: {filename}")
                    else:
                        # 非CSV转换：移动单个文件
                        final_output_path = os.path.join(output_dir, os.path.basename(result_path))
                        
                        should_keep = self._should_keep_intermediates()
                        if should_keep:
                            # 保留中间文件（排除输入副本）
                            logger.info("保留中间文件，移动规范命名的文件到输出目录")
                            for filename in os.listdir(temp_dir):
                                if filename.startswith('input.'):
                                    logger.debug(f"跳过输入副本: {filename}")
                                    continue
                                src = os.path.join(temp_dir, filename)
                                if os.path.isfile(src):
                                    dst = os.path.join(output_dir, filename)
                                    shutil.move(src, dst)
                                    logger.debug(f"保留中间文件: {filename}")
                        else:
                            # 只移动最终文件
                            logger.debug("清理中间文件，只移动最终文件")
                            shutil.move(result_path, final_output_path)
                            logger.debug(f"已移动最终文件: {os.path.basename(result_path)}")
                    
                    return ConversionResult(
                        success=True,
                        output_path=final_output_path,
                        message=t('conversion.messages.conversion_to_format_success', format=target_fmt.upper())
                    )
            
            except OfficeSoftwareNotFoundError as e:
                logger.error(f"未找到Office软件: {e}")
                return ConversionResult(
                    success=False,
                    message=t('conversion.messages.missing_office_for_conversion'),
                    error=str(e)
                )
            except Exception as e:
                logger.error(f"{source_fmt.upper()}转{target_fmt.upper()}失败: {e}", exc_info=True)
                return ConversionResult(
                    success=False,
                    message=t('conversion.messages.conversion_failed_with_error', error=str(e)),
                    error=e
                )
        
        @staticmethod
        def _should_keep_intermediates() -> bool:
            """判断是否应该保留中间文件"""
            try:
                from docwen.config.config_manager import config_manager
                return config_manager.get_save_intermediate_files()
            except Exception as e:
                logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
                return False
    
    return DynamicSpreadsheetConversionStrategy


# ==================== 批量注册表格格式转换策略 ====================

# 批量注册表格格式转换策略（排除 CSV ↔ XLSX，它们已有独立实现）
SPREADSHEET_FORMATS = ['xlsx', 'xls', 'ods', 'et']
for source in SPREADSHEET_FORMATS:
    for target in SPREADSHEET_FORMATS:
        if source != target:
            _create_spreadsheet_conversion_strategy(source, target)

logger.info("表格格式转换策略已通过智能转换链批量注册")

# 添加 CSV 到其他格式的转换支持（CSV → XLSX 已有独立实现）
# CSV → XLS: CSV → XLSX → XLS
# CSV → ODS: CSV → XLSX → ODS
for target in ['xls', 'ods']:
    _create_spreadsheet_conversion_strategy('csv', target)

logger.info("CSV 转换策略已注册（CSV → XLS, CSV → ODS）")

# 添加其他格式到 CSV 的转换支持（XLSX → CSV 已有独立实现）
# XLS → CSV: XLS → XLSX → CSV
# ODS → CSV: ODS → XLSX → CSV
# ET → CSV: ET → XLSX → CSV
for source in ['xls', 'ods', 'et']:
    _create_spreadsheet_conversion_strategy(source, 'csv')

logger.info("其他格式转CSV策略已注册（XLS/ODS/ET → CSV）")
