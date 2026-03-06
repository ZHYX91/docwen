"""
表格格式互转策略模块

提供表格格式之间的智能转换策略工厂。

依赖：
- smart_converter: 智能转换链
- formats.spreadsheet: 表格格式转换
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

from .. import register_conversion
from ..base_strategy import BaseStrategy

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

        @staticmethod
        def _finalize_csv_output(result_path: str, temp_dir: str, output_dir: str, original_input_file: str) -> str:
            from docwen.utils.workspace_manager import save_output_with_fallback

            result_path_path = Path(result_path)
            if result_path_path.is_dir():
                csv_folder_path = result_path_path
                csv_filename = None
            else:
                csv_folder_path = result_path_path.parent
                csv_filename = result_path_path.name

            if csv_folder_path.resolve() == Path(temp_dir).resolve():
                final_output_path = str(Path(output_dir) / result_path_path.name)
                saved_path, _ = save_output_with_fallback(
                    result_path, final_output_path, original_input_file=original_input_file
                )
                return saved_path or final_output_path

            folder_name = csv_folder_path.name
            final_output_folder = str(Path(output_dir) / folder_name)

            saved_folder, _ = save_output_with_fallback(
                str(csv_folder_path), final_output_folder, original_input_file=original_input_file
            )
            final_output_folder = saved_folder or final_output_folder
            logger.info(f"CSV文件夹已移动: {Path(final_output_folder).name}")

            if csv_filename:
                return str(Path(final_output_folder) / csv_filename)
            return final_output_folder

        def execute(
            self,
            file_path: str,
            options: dict[str, Any] | None = None,
            progress_callback: Callable[[str], None] | None = None,
        ) -> ConversionResult:
            """
            执行表格格式转换

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
                    try:
                        from docwen.utils.file_type_utils import detect_actual_file_format

                        actual_format = detect_actual_file_format(file_path)
                        logger.warning(f"未提供actual_format，自动检测: {actual_format}")
                    except Exception:
                        actual_format = Path(file_path).suffix.lower().lstrip(".")
                        logger.warning(f"未提供actual_format，从文件名推断: {actual_format}")

                logger.debug(f"表格转换策略: {source_fmt}→{target_fmt}, 真实格式: {actual_format}")

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
                        category="spreadsheet",
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

                    if target_fmt == "csv":
                        final_output_path = self._finalize_csv_output(
                            result_path, temp_dir, output_dir, original_input_file=file_path
                        )

                        # 处理中间文件（如果需要保留）
                        should_keep = self._should_keep_intermediates()
                        if should_keep:
                            logger.info("检查是否有中间文件需要保留")
                            from docwen.utils.workspace_manager import save_intermediates_from_temp_dir

                            saved_items = save_intermediates_from_temp_dir(
                                temp_dir,
                                output_dir,
                                move=True,
                                include_dirs=False,
                            )
                            for item, _saved_path in saved_items:
                                logger.info(f"保留中间文件: {Path(item.path).name}")
                    else:
                        # 非CSV转换：移动单个文件
                        final_output_path = str(Path(output_dir) / Path(result_path).name)

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

    return DynamicSpreadsheetConversionStrategy


# ==================== 批量注册表格格式转换策略 ====================

# 批量注册表格格式转换策略（排除 CSV ↔ XLSX，它们已有独立实现）
SPREADSHEET_FORMATS = ["xlsx", "xls", "ods", "et"]
for source in SPREADSHEET_FORMATS:
    for target in SPREADSHEET_FORMATS:
        if source != target:
            _create_spreadsheet_conversion_strategy(source, target)

logger.info("表格格式转换策略已通过智能转换链批量注册")

# 添加 CSV 到其他格式的转换支持（CSV → XLSX 已有独立实现）
# CSV → XLS: CSV → XLSX → XLS
# CSV → ODS: CSV → XLSX → ODS
for target in ["xls", "ods"]:
    _create_spreadsheet_conversion_strategy("csv", target)

logger.info("CSV 转换策略已注册（CSV → XLS, CSV → ODS）")

# 添加其他格式到 CSV 的转换支持（XLSX → CSV 已有独立实现）
# XLS → CSV: XLS → XLSX → CSV
# ODS → CSV: ODS → XLSX → CSV
# ET → CSV: ET → XLSX → CSV
for source in ["xls", "ods", "et"]:
    _create_spreadsheet_conversion_strategy(source, "csv")

logger.info("其他格式转CSV策略已注册（XLS/ODS/ET → CSV）")
