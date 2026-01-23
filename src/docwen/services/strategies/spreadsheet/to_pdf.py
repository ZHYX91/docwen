"""
表格转PDF策略模块

将表格文件转换为PDF格式。

依赖：
- formats.pdf_export: PDF导出
- formats.common: 异常类
"""

import os
import logging
import tempfile
import shutil
from typing import Dict, Any, Callable, Optional

from ..base_strategy import BaseStrategy
from docwen.services.result import ConversionResult
from .. import register_conversion, CATEGORY_SPREADSHEET
from docwen.utils.path_utils import generate_output_path
from docwen.i18n import t

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_SPREADSHEET, 'pdf')
class SpreadsheetToPdfStrategy(BaseStrategy):
    """
    将表格文件转换为PDF文件的策略。
    
    功能特性：
    - 使用本地Office软件（WPS或Microsoft Office）进行转换
    - 支持XLSX、XLS、ET、ODS、CSV等表格格式
    - 转换质量高，能保持表格格式和样式
    - 生成不可编辑的PDF文档，适合最终版本归档
    - 使用 office_to_pdf 配置的软件优先级
    
    支持的输入格式：
    - XLSX (Excel 2007+)
    - XLS (Excel 97-2003)
    - ET (WPS表格格式)
    - ODS (OpenDocument表格)
    - CSV (逗号分隔值)
    
    Note:
        需要本地安装WPS Office、Microsoft Office或LibreOffice才能使用此功能。
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行表格到PDF的转换。
        
        Args:
            file_path: 输入的表格文件路径（支持 XLSX/XLS/ET/ODS/CSV）
            options: 转换选项字典，包含：
                - cancel_event: (可选) 用于取消操作的事件对象
                - actual_format: (可选) 文件的真实格式
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        # 在try块外导入异常类，避免UnboundLocalError
        from docwen.converter.formats.common import OfficeSoftwareNotFoundError
        
        try:
            if progress_callback:
                progress_callback(t('conversion.progress.preparing'))
            
            options = options or {}
            cancel_event = options.get('cancel_event')
            actual_format = options.get('actual_format', 'xlsx')
            
            from docwen.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：创建输入副本 input.{ext}
                if progress_callback:
                    progress_callback(t('conversion.progress.preparing_files'))
                
                from docwen.utils.workspace_manager import prepare_input_file
                temp_input = prepare_input_file(file_path, temp_dir, actual_format)
                logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                # 步骤2：生成输出文件名
                output_filename = os.path.basename(
                    generate_output_path(
                        file_path,
                        section="",
                        add_timestamp=True,
                        description=f"from{actual_format.capitalize()}",
                        file_type="pdf"
                    )
                )
                
                # 步骤3：在临时目录进行转换
                temp_output = os.path.join(temp_dir, output_filename)
                
                if progress_callback:
                    progress_callback(t('conversion.progress.converting_to_format', format='PDF'))
                
                # 导入并调用转换函数
                from docwen.converter.formats.pdf_export import xlsx_to_pdf
                
                result_path = xlsx_to_pdf(
                    temp_input,  # 使用副本
                    output_path=temp_output,
                    cancel_event=cancel_event
                )
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                if not result_path or not os.path.exists(result_path):
                    return ConversionResult(success=False, message=t('conversion.messages.conversion_failed_check_log'))
                
                # 步骤4：移动到目标位置
                final_output = os.path.join(output_dir, output_filename)
                shutil.move(result_path, final_output)
                logger.info(f"PDF转换完成，文件已保存: {final_output}")
                
                # 步骤5：如果保留中间文件，移动temp_dir中的其他规范文件
                should_keep = self._should_keep_intermediates()
                if should_keep:
                    logger.info("检查并移动临时目录中的其他中间文件")
                    for filename in os.listdir(temp_dir):
                        if filename.startswith('input.') or filename == output_filename:
                            continue
                        src = os.path.join(temp_dir, filename)
                        if os.path.isfile(src):
                            dst = os.path.join(output_dir, filename)
                            shutil.move(src, dst)
                            logger.info(f"保留中间文件: {filename}")
                
                return ConversionResult(
                    success=True,
                    output_path=final_output,
                    message=t('conversion.messages.conversion_to_format_success', format='PDF')
                )
            
        except OfficeSoftwareNotFoundError as e:
            logger.error(f"表格转PDF失败 - 未找到Office软件: {e}")
            return ConversionResult(
                success=False,
                message=t('conversion.messages.missing_office_for_pdf'),
                error=str(e)
            )
        except Exception as e:
            logger.error(f"表格转PDF失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"转换失败: {e}",
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
