"""
文档转PDF/OFD策略模块

将文档文件转换为PDF或OFD格式。

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
from .. import register_conversion, CATEGORY_DOCUMENT
from docwen.utils.path_utils import generate_output_path
from docwen.i18n import t

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_DOCUMENT, 'pdf')
class DocumentToPdfStrategy(BaseStrategy):
    """
    将文档文件转换为PDF文件的策略。
    
    功能特性：
    - 使用本地Office软件（WPS或Microsoft Office）进行转换
    - 支持DOCX、DOC、WPS、RTF、ODT等文档格式
    - 转换质量高，能保持文档格式和样式
    - 生成不可编辑的PDF文档，适合最终版本归档
    - 使用 office_to_pdf 配置的软件优先级
    
    支持的输入格式：
    - DOCX (Word 2007+)
    - DOC (Word 97-2003)
    - WPS (WPS文字格式)
    - RTF (富文本格式)
    - ODT (OpenDocument文本)
    
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
        执行文档到PDF的转换。
        
        Args:
            file_path: 输入的文档文件路径（支持 DOCX/DOC/WPS/RTF/ODT）
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
            actual_format = options.get('actual_format', 'docx')
            
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
                from docwen.converter.formats.pdf_export import docx_to_pdf
                
                result_path = docx_to_pdf(
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
            logger.error(f"文档转PDF失败 - 未找到Office软件: {e}")
            return ConversionResult(
                success=False,
                message=t('conversion.messages.missing_office_for_pdf'),
                error=str(e)
            )
        except Exception as e:
            logger.error(f"文档转PDF失败: {e}", exc_info=True)
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


@register_conversion(CATEGORY_DOCUMENT, 'ofd')
class DocumentToOfdStrategy(BaseStrategy):
    """
    将DOCX文件转换为OFD文件的策略（占位实现）。
    
    当前状态：
    - 功能尚未实现
    - 调用时返回错误信息
    
    规划用途：
    - OFD是中国电子文件标准格式
    - 用于符合国产化文档标准的场景
    """

    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        OFD转换功能占位方法。
        
        Args:
            file_path: 输入的DOCX文件路径（暂未使用）
            options: 转换选项（暂未使用）
            progress_callback: 进度回调（暂未使用）
            
        Returns:
            ConversionResult: 返回失败结果，提示功能未实现
        """
        return ConversionResult(success=False, message=t('conversion.messages.ofd_conversion_not_implemented'))
