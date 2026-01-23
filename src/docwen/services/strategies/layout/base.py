"""
版式文件转文档策略基类

提供 PDF/OFD/XPS/CAJ 转换为文档格式的模板方法实现。
子类只需实现 _convert_docx_to_target() 方法即可支持新格式。

依赖：
- .utils: 预处理函数
- converter.formats.layout: PDF转DOCX
- converter.formats.common: Office软件检测
"""

import os
import logging
import tempfile
import shutil
from abc import abstractmethod
from typing import Dict, Any, Callable, Optional

from docwen.services.result import ConversionResult
from docwen.services.strategies.base_strategy import BaseStrategy
from docwen.i18n import t
from docwen.converter.formats.common import check_office_availability, OfficeSoftwareNotFoundError

from .utils import preprocess_layout_file, should_keep_intermediates

logger = logging.getLogger(__name__)


class LayoutToDocumentBaseStrategy(BaseStrategy):
    """
    PDF转文档格式的基类策略
    
    使用模板方法模式，封装公共流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. PDF → DOCX（使用外部工具）
    3. DOCX → 目标格式（子类实现）
    4. 文件移动和清理
    
    子类只需实现：
    - _convert_docx_to_target(): DOCX转目标格式的具体实现
    - _get_target_extension(): 返回目标文件扩展名
    - _get_format_name(): 返回格式名称（用于日志和消息）
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行版式文件到文档格式的转换（模板方法）
        
        参数:
            file_path: 输入的版式文件路径
            options: 转换选项字典
            progress_callback: 进度更新回调函数
            
        返回:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        actual_format = options.get("actual_format")
        target_ext = self._get_target_extension()
        format_name = self._get_format_name()
        
        try:
            # ========== 预检查: Office 软件可用性 ==========
            available, error_msg = check_office_availability(target_ext)
            if not available:
                logger.error(f"预检查失败: {target_ext} 格式需要 Office 软件")
                raise OfficeSoftwareNotFoundError(error_msg)
            
            if progress_callback:
                progress_callback(t('conversion.progress.preparing'))
            
            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # ========== 步骤1: 预处理（公共） ==========
                if actual_format and actual_format != 'pdf':
                    if progress_callback:
                        progress_callback(t('conversion.progress.converting_format_to_pdf', format=actual_format.upper()))
                    pdf_path, _ = preprocess_layout_file(
                        file_path, temp_dir, cancel_event, actual_format
                    )
                else:
                    pdf_path = file_path
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                # ========== 步骤2: PDF → DOCX（公共） ==========
                if progress_callback:
                    progress_callback(t('conversion.progress.step_pdf_to_docx'))
                
                from docwen.utils.path_utils import generate_output_path
                
                # 生成中间DOCX的标准化路径（在临时目录）
                docx_temp_path = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description="fromPdf",
                    file_type="docx"
                )
                
                from docwen.converter.formats.layout import pdf_to_docx
                
                docx_path = pdf_to_docx(
                    pdf_path,
                    docx_temp_path,
                    cancel_event
                )
                
                if not docx_path or not os.path.exists(docx_path):
                    return ConversionResult(
                        success=False,
                        message=t('conversion.messages.pdf_to_docx_failed')
                    )
                
                logger.info(f"中间DOCX已生成: {os.path.basename(docx_path)}")
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                # ========== 步骤3: DOCX → 目标格式（子类实现） ==========
                if progress_callback:
                    progress_callback(t('conversion.progress.step_docx_to_format', format=format_name))
                
                # 生成最终文件的标准化路径（在临时目录）
                target_temp_path = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description="fromPdf",
                    file_type=target_ext
                )
                
                target_path = self._convert_docx_to_target(
                    docx_path, target_temp_path, cancel_event
                )
                
                if not target_path or not os.path.exists(target_path):
                    return ConversionResult(
                        success=False,
                        message=t('conversion.messages.docx_to_format_failed', format=format_name)
                    )
                
                logger.info(f"最终{format_name}已生成: {os.path.basename(target_path)}")
                
                # ========== 步骤4: 移动文件到输出目录（公共） ==========
                from docwen.utils.workspace_manager import get_output_directory
                output_dir = get_output_directory(file_path)
                final_target_path = os.path.join(output_dir, os.path.basename(target_path))
                shutil.move(target_path, final_target_path)
                logger.info(f"{format_name}文件已移动到: {final_target_path}")
                
                # 根据配置决定是否保留中间DOCX文件
                if should_keep_intermediates():
                    final_docx_path = os.path.join(output_dir, os.path.basename(docx_path))
                    shutil.move(docx_path, final_docx_path)
                    logger.info(f"保留中间文件: {final_docx_path}")
                else:
                    logger.debug("清理中间DOCX文件")
                
                return ConversionResult(
                    success=True,
                    output_path=final_target_path,
                    message=t('conversion.messages.conversion_to_format_success', format=format_name)
                )
        
        except InterruptedError:
            return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
        except OfficeSoftwareNotFoundError as e:
            logger.error(f"Office软件未找到: {e}")
            return ConversionResult(
                success=False,
                message=t('conversion.messages.missing_office_for_conversion'),
                error=e
            )
        except Exception as e:
            logger.error(f"执行 {self.__class__.__name__} 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=t('conversion.messages.conversion_failed_with_error', error=str(e)), error=e)
    
    @abstractmethod
    def _convert_docx_to_target(
        self,
        docx_path: str,
        output_path: str,
        cancel_event: Optional[Any] = None
    ) -> Optional[str]:
        """
        将DOCX转换为目标格式（子类必须实现）
        
        参数:
            docx_path: 输入DOCX文件路径
            output_path: 输出文件路径
            cancel_event: 取消事件（可选）
            
        返回:
            成功时返回输出文件路径，失败时返回None
        """
        pass
    
    @abstractmethod
    def _get_target_extension(self) -> str:
        """返回目标文件扩展名（不含点），如 'doc', 'odt', 'rtf'"""
        pass
    
    @abstractmethod
    def _get_format_name(self) -> str:
        """返回格式名称（用于日志和用户消息），如 'DOC', 'ODT', 'RTF'"""
        pass
