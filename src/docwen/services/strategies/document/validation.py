"""
文档校对策略模块

对文档文件执行错别字校对。

依赖：
- docx_spell: 文档校对核心功能
- utils: 文档预处理
"""

import os
import logging
import tempfile
import shutil
from typing import Dict, Any, Callable, Optional

from ..base_strategy import BaseStrategy
from docwen.services.result import ConversionResult
from .. import register_action
from docwen.docx_spell.core import process_docx
from docwen.utils.path_utils import generate_output_path
from .utils import _preprocess_document_file
from docwen.i18n import t

logger = logging.getLogger(__name__)


@register_action("validate")
class DocxValidationStrategy(BaseStrategy):
    """
    对文档文件执行错别字校对的策略。
    
    支持的输入格式：
    - DOCX：直接校对
    - DOC/WPS：自动转换为DOCX后校对
    - RTF：自动转换为DOCX后校对
    - ODT：自动转换为DOCX后校对
    
    校对功能：
    - 敏感词检查
    - 错别字检查
    - 标点符号规范检查
    - 在文档中用批注标记问题位置
    
    输出：
    - 生成校对后的新DOCX文件
    - 文件名标记为 "checked"
    - 保留原文件不变
    - 如果配置保留中间文件，也会保存转换步骤的DOCX（如有）
    """

    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行文档文件的错别字校对（支持DOC/WPS/RTF/ODT自动转换）。
        
        Args:
            file_path: 输入的文档文件路径（DOCX/DOC/WPS/RTF/ODT）
            options: 校对选项字典，包含：
                - proofread_options: (必需) 校对选项字典
                  * None或空字典: 不进行校对（直接返回）
                  * 包含键值对: 启用相应的校对规则
                - cancel_event: (可选) 用于取消操作的事件对象
                - actual_format: (可选) 文件的真实格式
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 包含校对结果的对象
        """
        if options is None:
            options = {}
            
        if progress_callback:
            progress_callback(t('conversion.progress.preparing_proofread'))

        try:
            # 检查是否启用了任何校对规则
            proofread_options = options.get("proofread_options")
            if not proofread_options or not any(proofread_options.values()):
                return ConversionResult(success=True, output_path=file_path, message=t('conversion.messages.proofread_skipped'))

            cancel_event = options.get("cancel_event")
            actual_format = options.get("actual_format", 'docx')
            
            from docwen.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：预处理 - 自动转换 DOC/WPS/RTF/ODT → DOCX（如需要）
                if progress_callback:
                    progress_callback(t('conversion.progress.detecting_format'))
                
                processed_docx = _preprocess_document_file(
                    file_path, temp_dir, cancel_event, actual_format
                )
                
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                # 步骤2：生成校对输出文件名
                output_filename = os.path.basename(
                    generate_output_path(
                        file_path,
                        section="",
                        add_timestamp=True,
                        description="checked",
                        file_type="docx"
                    )
                )
                
                # 步骤3：在临时目录进行校对
                temp_output = os.path.join(temp_dir, output_filename)
                
                if progress_callback:
                    progress_callback(t('conversion.progress.proofreading'))
                
                result_path = process_docx(
                    processed_docx,  # 使用预处理后的DOCX
                    output_path=temp_output,
                    proofread_options=proofread_options,
                    progress_callback=progress_callback,
                    cancel_event=cancel_event
                )

                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))

                if not result_path or not os.path.exists(result_path):
                    return ConversionResult(success=False, message=t('conversion.messages.proofread_failed'))
                
                # 准备最终输出路径
                final_output = os.path.join(output_dir, output_filename)
                
                # 步骤4：根据配置移动文件
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
                    shutil.move(result_path, final_output)
                    logger.debug(f"已移动最终文件: {os.path.basename(result_path)}")
                
                logger.info(f"校对完成，文件已保存: {final_output}")
                
                return ConversionResult(success=True, output_path=final_output, message=t('conversion.messages.proofread_completed'))

        except Exception as e:
            logger.error(f"执行 DocxValidationStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"发生未知错误: {e}", error=e)
    
    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件"""
        try:
            from docwen.config.config_manager import config_manager
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
            return False
