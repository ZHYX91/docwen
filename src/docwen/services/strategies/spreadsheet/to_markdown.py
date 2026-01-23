"""
表格转Markdown/TXT策略模块

将表格文件转换为Markdown或TXT格式。

依赖：
- xlsx2md: 表格转MD核心转换
- utils: 表格预处理
"""

import os
import logging
import tempfile
import shutil
from typing import Dict, Any, Callable, Optional

from ..base_strategy import BaseStrategy
from docwen.services.result import ConversionResult
from .. import register_conversion, CATEGORY_SPREADSHEET
from docwen.converter.xlsx2md import convert_spreadsheet_to_md
from docwen.utils.path_utils import generate_output_path
from .utils import _preprocess_table_file
from docwen.i18n import t

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_SPREADSHEET, 'md')
class SpreadsheetToMarkdownStrategy(BaseStrategy):
    """
    将表格文件转换为Markdown文件的策略。
    
    支持的输入格式：
    - XLSX (Excel 2007+)：直接处理
    - XLS (Excel 97-2003)：自动转换为XLSX后处理
    - ET (WPS表格格式)：自动转换为XLSX后处理
    - CSV (逗号分隔值)：直接处理
    
    输出格式：
    - Markdown表格语法
    - 根据输入文件类型自动命名（fromXlsx 或 fromCsv）
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行表格到Markdown的转换（支持XLS/ET自动转换）。
        
        Args:
            file_path: 输入的表格文件路径（XLSX/XLS/ET/CSV）
            options: 转换选项字典，包含：
                - cancel_event: (可选) 用于取消操作的事件对象
                - extract_image: (可选) 是否提取图片
                - extract_ocr: (可选) 是否进行OCR识别
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        try:
            if progress_callback:
                progress_callback(t('conversion.progress.converting_to_format', format='Markdown'))
            
            options = options or {}
            cancel_event = options.get("cancel_event")
            actual_format = options.get("actual_format")
            
            # 从GUI获取导出选项
            extract_image = options.get('extract_image', False)
            extract_ocr = options.get('extract_ocr', False)
            
            logger.info(f"表格转MD - 导出选项: 提取图片={extract_image}, OCR={extract_ocr}")
            
            from docwen.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：预处理 - 将XLS/ET转换为XLSX（如需要）
                if progress_callback:
                    progress_callback(t('conversion.progress.detecting_format'))
                
                processed_file = _preprocess_table_file(file_path, temp_dir, cancel_event, actual_format)
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                # 步骤2：生成统一basename和创建临时子文件夹
                description = f"from{actual_format.capitalize()}" if actual_format else "fromXlsx"
                
                base_path = generate_output_path(
                    file_path,
                    section="",
                    add_timestamp=True,
                    description=description,
                    file_type="md"
                )
                
                basename = os.path.splitext(os.path.basename(base_path))[0]
                logger.debug(f"统一basename: {basename}")
                
                # 创建临时子文件夹
                temp_output_folder = os.path.join(temp_dir, basename)
                os.makedirs(temp_output_folder, exist_ok=True)
                logger.debug(f"创建临时子文件夹: {temp_output_folder}")
                
                # 步骤3：调用核心转换函数
                markdown_content = convert_spreadsheet_to_md(
                    processed_file,
                    extract_image=extract_image,
                    extract_ocr=extract_ocr,
                    output_folder=temp_output_folder,
                    original_file_path=file_path,
                    progress_callback=progress_callback
                )

                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))

                # 步骤4：写入Markdown文件
                if progress_callback:
                    progress_callback(t('conversion.progress.writing_file'))
                
                md_filename = f"{basename}.md"
                temp_output = os.path.join(temp_output_folder, md_filename)
                
                logger.debug(f"准备将Markdown内容写入: {temp_output}")
                with open(temp_output, 'w', encoding='utf-8') as f:
                    f.write(markdown_content)
                logger.info(f"Markdown内容已写入: {temp_output}")
                
                # 步骤5：移动整个文件夹到输出目录
                final_folder = os.path.join(output_dir, basename)
                
                if os.path.exists(final_folder):
                    shutil.rmtree(final_folder)
                    logger.debug(f"已删除现有文件夹: {final_folder}")
                
                shutil.move(temp_output_folder, final_folder)
                logger.info(f"已移动文件夹到: {final_folder}")
                
                # 步骤6：如果保留中间文件，移动temp_dir中的其他规范文件
                should_keep = self._should_keep_intermediates()
                if should_keep:
                    logger.info("检查并移动临时目录中的其他中间文件")
                    for filename in os.listdir(temp_dir):
                        if filename.startswith('input.') or filename == basename:
                            continue
                        src = os.path.join(temp_dir, filename)
                        if os.path.isfile(src):
                            dst = os.path.join(output_dir, filename)
                            shutil.move(src, dst)
                            logger.info(f"保留中间文件: {filename}")
                
                output_path = os.path.join(final_folder, md_filename)
                
                return ConversionResult(
                    success=True,
                    output_path=output_path,
                    message=t('conversion.messages.conversion_to_format_success', format='Markdown')
                )

        except Exception as e:
            logger.error(f"表格转Markdown失败: {e}", exc_info=True)
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
            logger.warning(f"读取清理配置失败: {e}，使用默认设置（清理中间文件）")
            return False


@register_conversion(CATEGORY_SPREADSHEET, 'txt')
class SpreadsheetToTxtStrategy(BaseStrategy):
    """
    将表格文件转换为TXT文件的策略。
    
    实现说明：
    - 复用表格转Markdown的转换逻辑
    - 先将表格转换为Markdown格式
    - 然后以TXT扩展名保存（内容仍为Markdown格式）
    
    支持的输入格式：
    - XLSX (Excel 2007+)
    - XLS (Excel 97-2003)
    - ET (WPS表格格式)
    - CSV (逗号分隔值)
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行表格到TXT的转换。
        
        Args:
            file_path: 输入的表格文件路径
            options: 转换选项字典
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含转换结果的对象
        """
        try:
            if progress_callback:
                progress_callback(t('conversion.progress.converting_to_format', format='TXT'))

            options = options or {}
            actual_format = options.get("actual_format")
            
            description = f"from{actual_format.capitalize()}" if actual_format else "fromXlsx"

            # 先转换为Markdown格式（内部表示）
            markdown_content = convert_spreadsheet_to_md(file_path)
            
            # 生成TXT输出路径
            txt_output = generate_output_path(
                file_path, 
                section="", 
                add_timestamp=True, 
                description=description, 
                file_type="txt"
            )
            
            # 将Markdown格式内容写入TXT文件
            logger.debug(f"准备将内容写入TXT文件: {txt_output}")
            try:
                with open(txt_output, 'w', encoding='utf-8') as f:
                    f.write(markdown_content)
                logger.info(f"成功将内容写入TXT文件: {txt_output}")
            except Exception as e:
                logger.error(f"写入TXT文件失败: {e}", exc_info=True)
                raise IOError(f"无法写入输出文件: {e}")
            
            return ConversionResult(
                success=True,
                output_path=txt_output,
                message=t('conversion.messages.conversion_to_format_success', format='TXT')
            )

        except Exception as e:
            logger.error(f"表格转TXT失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                error=f"转换失败: {e}"
            )
