"""
版式文件转PDF策略

将OFD/XPS/CAJ等版式文件转换为PDF格式。

支持的源格式：
- OFD：国产版式文件
- XPS：微软版式文件
- CAJ：中国知网文件

依赖：
- converter.formats.layout: 版式格式转换
"""

import logging
from collections.abc import Callable
from pathlib import Path
from typing import Any

from docwen.translation import t
from docwen.services.error_codes import (
    ERROR_CODE_CONVERSION_FAILED,
    ERROR_CODE_NOT_IMPLEMENTED,
    ERROR_CODE_OPERATION_CANCELLED,
    ERROR_CODE_UNSUPPORTED_FORMAT,
)
from docwen.services.result import ConversionResult
from docwen.services.strategies import CATEGORY_LAYOUT, register_conversion
from docwen.services.strategies.base_strategy import BaseStrategy

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_LAYOUT, "pdf")
class LayoutToPdfStrategy(BaseStrategy):
    """
    将版式文件（OFD/XPS/CAJ）统一转换为PDF的策略

    这是一个统一的入口策略，会根据实际文件格式自动分发到对应的转换函数。
    解决了GUI中convert_layout_to_pdf策略缺失的问题。

    转换流程：
    1. 检测或使用提供的actual_format
    2. 根据格式调用对应的转换函数：
       - PDF: 直接返回（无需转换）
       - OFD: 调用ofd_to_pdf
       - XPS: 调用xps_to_pdf
       - CAJ: 调用caj_to_pdf（待实现）
    3. 保存PDF到输出目录

    支持的输入格式：
    - PDF：直接返回，无需转换
    - OFD：转换为PDF
    - XPS：转换为PDF
    - CAJ：转换为PDF（待实现）
    """

    def execute(
        self,
        file_path: str,
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行版式文件到PDF的统一转换

        参数:
            file_path: 输入的版式文件路径
            options: 转换选项字典，包含：
                - cancel_event: 取消事件
                - actual_format: 实际文件格式（可选，如果不提供则自动检测）
            progress_callback: 进度更新回调函数

        返回:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        actual_format = options.get("actual_format")

        try:
            # 自动检测格式（如果未提供）
            if not actual_format:
                from docwen.utils.file_type_utils import detect_actual_file_format

                actual_format = detect_actual_file_format(file_path)
                logger.debug(f"自动检测版式文件格式: {actual_format}")
            else:
                logger.debug(f"使用提供的文件格式: {actual_format}")

            # 如果已经是PDF，直接返回成功
            if actual_format == "pdf":
                logger.info("文件已是PDF格式，无需转换")
                return ConversionResult(
                    success=True, output_path=file_path, message=t("conversion.messages.file_already_pdf")
                )

            if progress_callback:
                progress_callback(t("conversion.progress.preparing_format", format=actual_format.upper()))

            # 确定输出目录
            from docwen.utils.workspace_manager import get_output_directory

            output_dir = get_output_directory(file_path)

            # 根据格式调用对应的转换函数
            if actual_format == "ofd":
                if progress_callback:
                    progress_callback(t("conversion.progress.converting_ofd_to_pdf"))

                from docwen.converter.formats.layout import ofd_to_pdf

                result_path = ofd_to_pdf(file_path, cancel_event, output_dir=output_dir)

            elif actual_format == "xps":
                if progress_callback:
                    progress_callback(t("conversion.progress.converting_xps_to_pdf"))

                from docwen.converter.formats.layout import xps_to_pdf

                result_path = xps_to_pdf(file_path, cancel_event, output_dir=output_dir)

            elif actual_format == "caj":
                if progress_callback:
                    progress_callback(t("conversion.progress.converting_caj_to_pdf"))

                from docwen.converter.formats.layout import caj_to_pdf

                result_path = caj_to_pdf(file_path, cancel_event, output_dir=output_dir)

            else:
                return ConversionResult(
                    success=False,
                    message=t("conversion.messages.unsupported_layout_format", format=actual_format),
                    error_code=ERROR_CODE_UNSUPPORTED_FORMAT,
                )

            # 检查取消
            if cancel_event and cancel_event.is_set():
                return ConversionResult(
                    success=False,
                    message=t("conversion.messages.operation_cancelled"),
                    error_code=ERROR_CODE_OPERATION_CANCELLED,
                )

            # 检查转换结果
            if not result_path or not Path(result_path).exists():
                return ConversionResult(
                    success=False,
                    message=t("conversion.messages.conversion_failed_check_log"),
                    error_code=ERROR_CODE_CONVERSION_FAILED,
                )

            logger.info(f"{actual_format.upper()}转PDF成功: {result_path}")

            return ConversionResult(
                success=True,
                output_path=result_path,
                message=t("conversion.messages.conversion_to_format_success", format="PDF"),
            )

        except InterruptedError:
            return ConversionResult(
                success=False,
                message=t("conversion.messages.operation_cancelled"),
                error_code=ERROR_CODE_OPERATION_CANCELLED,
            )
        except NotImplementedError as e:
            format_name = actual_format.upper() if actual_format else "Unknown"
            logger.error(f"{format_name}转PDF功能尚未实现: {e}")
            return ConversionResult(
                success=False,
                message=t(
                    "conversion.messages.format_to_pdf_not_implemented",
                    default="{format} 转 PDF 未实现",
                    format=format_name,
                ),
                error_code=ERROR_CODE_NOT_IMPLEMENTED,
            )
        except Exception as e:
            logger.error(f"执行 LayoutToPdfStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=t(
                    "conversion.messages.conversion_failed_with_error",
                    default="转换失败: {error}",
                    error=str(e),
                ),
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )


@register_conversion("ofd", "pdf")
class OfdToPdfStrategy(BaseStrategy):
    """
    将OFD文件转换为PDF的策略

    转换流程：
    1. 调用底层ofd_to_pdf转换函数
    2. 保存PDF到输出目录（与原文件同目录）

    特点：
    - 直接保存为PDF，不作为中间文件
    - 支持取消操作
    - 提供进度反馈
    """

    def execute(
        self,
        file_path: str,
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行OFD到PDF的转换

        参数:
            file_path: 输入的OFD文件路径
            options: 转换选项字典，包含：
                - cancel_event: 取消事件
            progress_callback: 进度更新回调函数

        返回:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")

        try:
            if progress_callback:
                progress_callback(t("conversion.progress.preparing_format", format="OFD"))

            # 确定输出目录和文件名
            from docwen.utils.workspace_manager import get_output_directory

            output_dir = get_output_directory(file_path)
            basename = Path(file_path).stem
            output_path = str(Path(output_dir) / f"{basename}.pdf")

            # 调用底层转换函数
            if progress_callback:
                progress_callback(t("conversion.progress.converting_ofd_to_pdf"))

            from docwen.converter.formats.layout import ofd_to_pdf

            result_path = ofd_to_pdf(file_path, cancel_event, output_dir=output_dir)

            # 检查取消
            if cancel_event and cancel_event.is_set():
                return ConversionResult(
                    success=False,
                    message=t("conversion.messages.operation_cancelled"),
                    error_code=ERROR_CODE_OPERATION_CANCELLED,
                )

            if not result_path or not Path(result_path).exists():
                return ConversionResult(
                    success=False,
                    message=t("conversion.messages.ofd_to_pdf_failed"),
                    error_code=ERROR_CODE_CONVERSION_FAILED,
                )

            # 如果输出路径不同，移动文件
            if result_path != output_path:
                from docwen.utils.workspace_manager import save_output_with_fallback

                saved_path, _ = save_output_with_fallback(result_path, output_path, original_input_file=file_path)
                if saved_path:
                    output_path = saved_path
                logger.info(f"PDF文件已移动到: {output_path}")

            logger.info(f"OFD转PDF成功: {output_path}")

            return ConversionResult(
                success=True,
                output_path=output_path,
                message=t("conversion.messages.conversion_to_format_success", format="PDF"),
            )

        except InterruptedError:
            return ConversionResult(
                success=False,
                message=t("conversion.messages.operation_cancelled"),
                error_code=ERROR_CODE_OPERATION_CANCELLED,
            )
        except NotImplementedError as e:
            logger.error(f"OFD转PDF功能尚未实现: {e}")
            return ConversionResult(
                success=False,
                message=t(
                    "conversion.messages.format_to_pdf_not_implemented",
                    default="{format} 转 PDF 未实现",
                    format="OFD",
                ),
                error_code=ERROR_CODE_NOT_IMPLEMENTED,
            )
        except Exception as e:
            logger.error(f"执行 OfdToPdfStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=t(
                    "conversion.messages.conversion_failed_with_error",
                    default="转换失败: {error}",
                    error=str(e),
                ),
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )


@register_conversion("xps", "pdf")
class XpsToPdfStrategy(BaseStrategy):
    """
    将XPS文件转换为PDF的策略

    转换流程：
    1. 调用底层xps_to_pdf转换函数
    2. 保存PDF到输出目录（与原文件同目录）

    特点：
    - 直接保存为PDF，不作为中间文件
    - 支持取消操作
    - 提供进度反馈
    """

    def execute(
        self,
        file_path: str,
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行XPS到PDF的转换

        参数:
            file_path: 输入的XPS文件路径
            options: 转换选项字典，包含：
                - cancel_event: 取消事件
            progress_callback: 进度更新回调函数

        返回:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")

        try:
            if progress_callback:
                progress_callback(t("conversion.progress.preparing_format", format="XPS"))

            # 确定输出目录和文件名
            from docwen.utils.workspace_manager import get_output_directory

            output_dir = get_output_directory(file_path)
            basename = Path(file_path).stem
            output_path = str(Path(output_dir) / f"{basename}.pdf")

            # 调用底层转换函数
            if progress_callback:
                progress_callback(t("conversion.progress.converting_xps_to_pdf"))

            from docwen.converter.formats.layout import xps_to_pdf

            result_path = xps_to_pdf(file_path, cancel_event, output_dir=output_dir)

            # 检查取消
            if cancel_event and cancel_event.is_set():
                return ConversionResult(
                    success=False,
                    message=t("conversion.messages.operation_cancelled"),
                    error_code=ERROR_CODE_OPERATION_CANCELLED,
                )

            if not result_path or not Path(result_path).exists():
                return ConversionResult(
                    success=False,
                    message=t("conversion.messages.xps_to_pdf_failed"),
                    error_code=ERROR_CODE_CONVERSION_FAILED,
                )

            # 如果输出路径不同，移动文件
            if result_path != output_path:
                from docwen.utils.workspace_manager import save_output_with_fallback

                saved_path, _ = save_output_with_fallback(result_path, output_path, original_input_file=file_path)
                if saved_path:
                    output_path = saved_path
                logger.info(f"PDF文件已移动到: {output_path}")

            logger.info(f"XPS转PDF成功: {output_path}")

            return ConversionResult(
                success=True,
                output_path=output_path,
                message=t("conversion.messages.conversion_to_format_success", format="PDF"),
            )

        except InterruptedError:
            return ConversionResult(
                success=False,
                message=t("conversion.messages.operation_cancelled"),
                error_code=ERROR_CODE_OPERATION_CANCELLED,
            )
        except Exception as e:
            logger.error(f"执行 XpsToPdfStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=t(
                    "conversion.messages.conversion_failed_with_error",
                    default="转换失败: {error}",
                    error=str(e),
                ),
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )
