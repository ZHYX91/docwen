"""
文档格式互转策略模块

提供文档格式之间的智能转换策略工厂。

依赖：
- smart_converter: 智能转换链
- formats.document: 文档格式转换
"""

import logging
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from docwen.services.error_codes import (
    ERROR_CODE_CONVERSION_FAILED,
    ERROR_CODE_DEPENDENCY_MISSING,
    ERROR_CODE_OPERATION_CANCELLED,
)
from docwen.services.result import ConversionResult
from docwen.translation import t

from .. import register_conversion
from ..base_strategy import BaseStrategy

logger = logging.getLogger(__name__)


def _create_document_conversion_strategy(source_fmt: str, target_fmt: str):
    """
    策略工厂：动态创建文档格式转换策略

    使用智能转换链自动处理单步或多步转换

    参数:
        source_fmt: 源格式（如 'doc', 'odt'）
        target_fmt: 目标格式（如 'docx', 'odt'）

    返回:
        动态生成的策略类
    """

    @register_conversion(source_fmt, target_fmt)
    class DynamicDocumentConversionStrategy(BaseStrategy):
        """动态生成的文档转换策略"""

        def execute(
            self,
            file_path: str,
            options: dict[str, Any] | None = None,
            progress_callback: Callable[[str], None] | None = None,
        ) -> ConversionResult:
            """
            执行文档格式转换

            说明:
                从options中提取actual_format参数并传递给SmartConverter，
                确保即使文件扩展名被修改也能正确转换
            """
            try:
                options = options or {}
                cancel_event = options.get("cancel_event")
                preferred_software = options.get("preferred_software")
                actual_format = options.get("actual_format")

                # 如果没有提供actual_format，从文件扩展名推断（降级方案）
                if not actual_format:
                    actual_format = Path(file_path).suffix.lower().lstrip(".")
                    logger.warning(f"未提供actual_format，从文件名推断: {actual_format}")

                logger.debug(f"文档转换策略: {source_fmt}→{target_fmt}, 真实格式: {actual_format}")

                # 使用智能转换链
                from docwen.converter.smart_converter import OfficeSoftwareNotFoundError, SmartConverter
                from docwen.utils.workspace_manager import get_output_directory, save_output_with_fallback

                converter = SmartConverter()
                output_dir = get_output_directory(file_path)

                # 使用临时目录管理中间文件
                with tempfile.TemporaryDirectory() as temp_dir:
                    # 调用SmartConverter，输出到临时目录
                    result_path = converter.convert(
                        input_path=file_path,
                        target_format=target_fmt,
                        category="document",
                        actual_format=actual_format,
                        output_dir=temp_dir,
                        cancel_event=cancel_event,
                        progress_callback=progress_callback,
                        preferred_software=preferred_software,
                    )
                    if cancel_event and cancel_event.is_set():
                        return ConversionResult(
                            success=False,
                            message=t("conversion.messages.operation_cancelled"),
                            error_code=ERROR_CODE_OPERATION_CANCELLED,
                        )

                    if not result_path or not Path(result_path).exists():
                        return ConversionResult(
                            success=False,
                            message=t("conversion.messages.conversion_to_format_failed", format=target_fmt.upper()),
                            error_code=ERROR_CODE_CONVERSION_FAILED,
                        )

                    # 准备最终输出路径
                    final_output_path = str(Path(output_dir) / Path(result_path).name)

                    # 根据配置移动文件
                    should_keep = self._should_keep_intermediates()
                    if should_keep:
                        # 保留中间文件（排除输入副本）
                        logger.info("保留中间文件，移动规范命名的文件到输出目录")
                        result_filename = Path(result_path).name
                        from docwen.utils.workspace_manager import save_intermediates_from_temp_dir

                        saved_items = save_intermediates_from_temp_dir(
                            temp_dir,
                            output_dir,
                            move=True,
                            include_dirs=False,
                        )
                        for item, saved_path in saved_items:
                            filename = Path(item.path).name
                            logger.debug(f"保留中间文件: {filename}")
                            if filename == result_filename:
                                final_output_path = saved_path
                    else:
                        # 只移动最终文件
                        logger.debug("清理中间文件，只移动最终文件")
                        final_saved_path, _ = save_output_with_fallback(
                            result_path, final_output_path, original_input_file=file_path
                        )
                        if final_saved_path:
                            final_output_path = final_saved_path
                        logger.debug(f"已移动最终文件: {Path(final_output_path).name}")

                    return ConversionResult(
                        success=True,
                        output_path=final_output_path,
                        message=t("conversion.messages.conversion_to_format_success", format=target_fmt.upper()),
                    )

            except OfficeSoftwareNotFoundError as e:
                logger.error(f"未找到Office软件: {e}")
                return ConversionResult(
                    success=False,
                    message=t("conversion.messages.missing_office_for_conversion"),
                    error=e,
                    error_code=ERROR_CODE_DEPENDENCY_MISSING,
                    details=str(e),
                )
            except Exception as e:
                logger.error(f"{source_fmt.upper()}转{target_fmt.upper()}失败: {e}", exc_info=True)
                return ConversionResult(
                    success=False,
                    message=t("conversion.messages.conversion_failed_with_error", error=str(e)),
                    error=e,
                    error_code=ERROR_CODE_CONVERSION_FAILED,
                    details=str(e) or None,
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

    return DynamicDocumentConversionStrategy


# ==================== 批量注册文档格式转换策略 ====================

DOCUMENT_FORMATS = ["docx", "doc", "odt", "rtf", "wps"]
for source in DOCUMENT_FORMATS:
    for target in DOCUMENT_FORMATS:
        if source != target:
            _create_document_conversion_strategy(source, target)

logger.info("文档格式转换策略已通过智能转换链批量注册")
