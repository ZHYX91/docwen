"""
版式文件转文档格式策略

将PDF/OFD/XPS/CAJ等版式文件转换为可编辑文档格式。

支持的目标格式：
- DOCX：Word文档（直接转换）
- DOC：旧版Word文档（通过DOCX中转）
- ODT：OpenDocument文档（通过DOCX中转）
- RTF：富文本格式（通过DOCX中转）

依赖：
- .base: 模板方法基类
- .utils: 预处理函数
- converter.formats.layout: PDF转DOCX
- converter.formats.document: DOCX格式转换
"""

import logging
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from docwen.translation import t
from docwen.services.error_codes import (
    ERROR_CODE_CONVERSION_FAILED,
    ERROR_CODE_DEPENDENCY_MISSING,
    ERROR_CODE_OPERATION_CANCELLED,
)
from docwen.services.result import ConversionResult
from docwen.services.strategies import CATEGORY_LAYOUT, register_conversion
from docwen.services.strategies.base_strategy import BaseStrategy

from .base import LayoutToDocumentBaseStrategy
from .utils import preprocess_layout_file, should_keep_intermediates

logger = logging.getLogger(__name__)


# ==================== 独立实现的策略 ====================


@register_conversion(CATEGORY_LAYOUT, "docx")
class LayoutToDocxStrategy(BaseStrategy):
    """
    使用外部工具将PDF转换为DOCX的策略

    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. 核心转换：PDF → DOCX（使用外部工具）
    3. 移动到输出目录

    特点：
    - 直接转换，不需要DOCX中转
    - 使用 generate_output_path() 统一命名
    - description 统一为 "from{实际格式}"
    """

    def execute(
        self,
        file_path: str,
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行PDF到DOCX的转换

        参数:
            file_path: 输入的PDF文件路径
            options: 转换选项字典
            progress_callback: 进度更新回调函数

        返回:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        actual_format = options.get("actual_format")

        if not actual_format:
            from docwen.utils.file_type_utils import detect_actual_file_format

            actual_format = detect_actual_file_format(file_path)

        try:
            if progress_callback:
                progress_callback(t("conversion.progress.preparing"))

            with tempfile.TemporaryDirectory() as temp_dir:
                # 预处理：确保文件是PDF格式
                if actual_format and actual_format != "pdf" and progress_callback:
                    progress_callback(t("conversion.progress.converting_format_to_pdf", format=actual_format.upper()))

                preprocess_result = preprocess_layout_file(file_path, temp_dir, cancel_event, actual_format)
                pdf_path = preprocess_result.processed_file
                actual_format = preprocess_result.actual_format
                options["actual_format"] = actual_format

                if cancel_event and cancel_event.is_set():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.operation_cancelled"),
                        error_code=ERROR_CODE_OPERATION_CANCELLED,
                    )

                # 使用标准化路径生成
                if progress_callback:
                    progress_callback(t("conversion.progress.using_external_tool_pdf"))

                from docwen.utils.path_utils import generate_output_path

                docx_temp_path = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description=f"from{actual_format.capitalize()}",
                    file_type="docx",
                )

                from docwen.converter.formats.layout import pdf_to_docx

                docx_path = pdf_to_docx(
                    pdf_path, docx_temp_path, cancel_event=cancel_event, headless=options.get("headless", False)
                )

                if not docx_path or not Path(docx_path).exists():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.external_tool_failed"),
                        error_code=ERROR_CODE_CONVERSION_FAILED,
                    )

                logger.info(f"DOCX文件已生成: {Path(docx_path).name}")

                # 移动文件到输出目录
                from docwen.utils.workspace_manager import get_output_directory, save_output_with_fallback

                output_dir = get_output_directory(file_path)
                final_docx_path = str(Path(output_dir) / Path(docx_path).name)
                final_saved_path, _ = save_output_with_fallback(
                    docx_path, final_docx_path, original_input_file=file_path
                )
                if final_saved_path:
                    final_docx_path = final_saved_path
                logger.info(f"DOCX文件已移动到: {final_docx_path}")

                if should_keep_intermediates():
                    from docwen.utils.workspace_manager import save_intermediate_items

                    saved_items = save_intermediate_items(preprocess_result.intermediates, output_dir, move=True)
                    for _, saved_path in saved_items:
                        logger.info(f"保留中间文件: {Path(saved_path).name}")

                return ConversionResult(
                    success=True,
                    output_path=final_docx_path,
                    message=t("conversion.messages.conversion_to_format_success", format="DOCX"),
                )

        except InterruptedError:
            return ConversionResult(
                success=False,
                message=t("conversion.messages.operation_cancelled"),
                error_code=ERROR_CODE_OPERATION_CANCELLED,
            )
        except ImportError as e:
            error_msg = str(e)
            logger.error(f"缺少必要的库: {error_msg}")
            try:
                if should_keep_intermediates() and "preprocess_result" in locals():
                    from docwen.utils.workspace_manager import get_output_directory, save_intermediate_items

                    failure_output_dir = get_output_directory(file_path)

                    saved_items = save_intermediate_items(
                        preprocess_result.intermediates, failure_output_dir, move=True
                    )
                    for _, saved_path in saved_items:
                        logger.info(f"保留中间文件: {Path(saved_path).name}")
            except Exception:
                pass
            return ConversionResult(
                success=False,
                message=error_msg,
                error=e,
                error_code=ERROR_CODE_DEPENDENCY_MISSING,
                details=error_msg or None,
            )
        except Exception as e:
            logger.error(f"执行 LayoutToDocxStrategy 时出错: {e}", exc_info=True)
            try:
                if should_keep_intermediates() and "preprocess_result" in locals():
                    from docwen.utils.workspace_manager import get_output_directory, save_intermediate_items

                    failure_output_dir = get_output_directory(file_path)

                    saved_items = save_intermediate_items(
                        preprocess_result.intermediates, failure_output_dir, move=True
                    )
                    for _, saved_path in saved_items:
                        logger.info(f"保留中间文件: {Path(saved_path).name}")
            except Exception:
                pass
            return ConversionResult(
                success=False,
                message=f"转换失败: {e!s}",
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )


# ==================== 使用模板方法的策略 ====================


@register_conversion(CATEGORY_LAYOUT, "doc")
class LayoutToDocStrategy(LayoutToDocumentBaseStrategy):
    """
    将PDF转换为DOC格式的策略

    继承 LayoutToDocumentBaseStrategy，只需实现 DOCX→DOC 转换。
    转换流程由基类模板方法控制。
    """

    def _convert_docx_to_target(self, docx_path: str, output_path: str, cancel_event: Any | None = None) -> str | None:
        """使用 formats.document 模块转换 DOCX→DOC"""
        from docwen.converter.formats.document import docx_to_doc

        return docx_to_doc(docx_path, output_path, cancel_event=cancel_event)

    def _get_target_extension(self) -> str:
        return "doc"

    def _get_format_name(self) -> str:
        return "DOC"


@register_conversion(CATEGORY_LAYOUT, "odt")
class LayoutToOdtStrategy(LayoutToDocumentBaseStrategy):
    """
    将PDF转换为ODT格式的策略

    继承 LayoutToDocumentBaseStrategy，只需实现 DOCX→ODT 转换。
    转换流程由基类模板方法控制。
    """

    def _convert_docx_to_target(self, docx_path: str, output_path: str, cancel_event: Any | None = None) -> str | None:
        """使用 formats.document 模块转换 DOCX→ODT"""
        from docwen.converter.formats.document import docx_to_odt

        return docx_to_odt(docx_path, output_path, cancel_event=cancel_event)

    def _get_target_extension(self) -> str:
        return "odt"

    def _get_format_name(self) -> str:
        return "ODT"


@register_conversion(CATEGORY_LAYOUT, "rtf")
class LayoutToRtfStrategy(LayoutToDocumentBaseStrategy):
    """
    将PDF转换为RTF格式的策略

    继承 LayoutToDocumentBaseStrategy，只需实现 DOCX→RTF 转换。
    转换流程由基类模板方法控制。
    """

    def _convert_docx_to_target(self, docx_path: str, output_path: str, cancel_event: Any | None = None) -> str | None:
        """使用 formats.document 模块转换 DOCX→RTF"""
        from docwen.converter.formats.document import docx_to_rtf

        return docx_to_rtf(docx_path, output_path, cancel_event=cancel_event)

    def _get_target_extension(self) -> str:
        return "rtf"

    def _get_format_name(self) -> str:
        return "RTF"
