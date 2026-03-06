"""
CSV与XLSX互转策略模块

提供CSV和XLSX格式之间的转换策略。

依赖：
- pandas: 数据处理
- table_converters: 表格转换工具
"""

import logging
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from docwen.converter.formats.spreadsheet import xlsx_to_csv
from docwen.translation import t
from docwen.services.error_codes import ERROR_CODE_CONVERSION_FAILED, ERROR_CODE_OPERATION_CANCELLED
from docwen.services.result import ConversionResult
from docwen.utils.path_utils import generate_output_path

from .. import register_conversion
from ..base_strategy import BaseStrategy

logger = logging.getLogger(__name__)


@register_conversion("csv", "xlsx")
class CsvToXlsxStrategy(BaseStrategy):
    """
    将CSV文件转换为XLSX文件的策略

    转换说明：
    - CSV文件将被读取并保存为XLSX格式
    - 输出文件使用默认工作表名 "Sheet1"
    - 文件名标记为 "fromCsv"
    - 支持扩展名不匹配的文件
    """

    def execute(
        self,
        file_path: str,
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行CSV到XLSX的转换

        Args:
            file_path: 输入的CSV文件路径
            options: 转换选项字典，包含：
                - actual_format: (可选) 文件的真实格式
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数

        Returns:
            ConversionResult: 包含转换结果的对象
        """
        try:
            if progress_callback:
                progress_callback(t("conversion.progress.preparing"))

            options = options or {}
            actual_format = options.get("actual_format")
            if not actual_format:
                try:
                    from docwen.utils.file_type_utils import detect_actual_file_format

                    actual_format = detect_actual_file_format(file_path)
                except Exception:
                    actual_format = "csv"
            cancel_event = options.get("cancel_event")
            output_dir = str(Path(file_path).parent)

            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：创建输入副本 input.csv
                if progress_callback:
                    progress_callback(t("conversion.progress.preparing_files"))

                from docwen.utils.workspace_manager import prepare_input_file

                temp_input = prepare_input_file(file_path, temp_dir, actual_format)
                logger.debug(f"已创建输入副本: {Path(temp_input).name}")

                if cancel_event and cancel_event.is_set():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.operation_cancelled"),
                        error_code=ERROR_CODE_OPERATION_CANCELLED,
                    )

                # 步骤2：读取CSV并转换
                if progress_callback:
                    progress_callback(t("conversion.progress.converting_to_format", format="XLSX"))

                import pandas as pd

                df = pd.read_csv(temp_input, header=None, keep_default_na=False)
                logger.debug(f"CSV 文件读取成功，数据形状: {df.shape}")

                # 步骤3：生成输出文件名
                output_filename = Path(
                    generate_output_path(
                        file_path, section="", add_timestamp=True, description="fromCsv", file_type="xlsx"
                    )
                ).name

                # 步骤4：保存到临时目录
                temp_output = str(Path(temp_dir) / output_filename)
                with pd.ExcelWriter(temp_output, engine="openpyxl") as writer:
                    df.to_excel(writer, sheet_name="Sheet1", index=False, header=False)
                logger.debug(f"XLSX文件已生成: {temp_output}")

                if cancel_event and cancel_event.is_set():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.operation_cancelled"),
                        error_code=ERROR_CODE_OPERATION_CANCELLED,
                    )

                # 步骤5：移动到目标位置
                final_output = str(Path(output_dir) / output_filename)
                from docwen.utils.workspace_manager import save_output_with_fallback

                saved_path, _ = save_output_with_fallback(temp_output, final_output, original_input_file=file_path)
                if saved_path:
                    final_output = saved_path
                logger.info(f"CSV转XLSX成功: {final_output}")

            return ConversionResult(
                success=True,
                output_path=final_output,
                message=t("conversion.messages.conversion_to_format_success", format="XLSX"),
            )

        except Exception as e:
            logger.error(f"CSV转XLSX失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"转换失败: {e}",
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )


@register_conversion("xlsx", "csv")
class XlsxToCsvStrategy(BaseStrategy):
    """
    将XLSX文件转换为CSV文件的策略

    转换说明：
    - 每个工作表将生成一个独立的CSV文件
    - 工作表名中的空格将被替换为下划线
    - 文件名包含工作表名作为section，标记为 "fromXlsx"
    - 所有CSV文件使用相同的时间戳
    - 支持扩展名不匹配的文件
    - 所有CSV文件输出到同一个子文件夹
    """

    def execute(
        self,
        file_path: str,
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行XLSX到CSV的转换

        Args:
            file_path: 输入的XLSX文件路径
            options: 转换选项字典，包含：
                - actual_format: (可选) 文件的真实格式
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数

        Returns:
            ConversionResult: 包含转换结果的对象

        Note:
            output_path 返回子文件夹中第一个CSV文件的路径
        """
        try:
            if progress_callback:
                progress_callback(t("conversion.progress.preparing"))

            options = options or {}
            actual_format = options.get("actual_format")
            if not actual_format:
                try:
                    from docwen.utils.file_type_utils import detect_actual_file_format

                    actual_format = detect_actual_file_format(file_path)
                except Exception:
                    actual_format = "xlsx"
            cancel_event = options.get("cancel_event")

            from docwen.utils.workspace_manager import get_output_directory

            output_dir = get_output_directory(file_path)

            # 使用临时目录管理输出
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：创建输入副本 input.xlsx
                if progress_callback:
                    progress_callback(t("conversion.progress.preparing_files"))

                from docwen.utils.workspace_manager import prepare_input_file

                temp_input = prepare_input_file(file_path, temp_dir, actual_format)
                logger.debug(f"已创建输入副本: {Path(temp_input).name}")

                if cancel_event and cancel_event.is_set():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.operation_cancelled"),
                        error_code=ERROR_CODE_OPERATION_CANCELLED,
                    )

                # 步骤2：生成统一basename和时间戳描述部分
                from docwen.utils.path_utils import generate_timestamp, make_output_stem, strip_timestamp_suffix

                timestamp = generate_timestamp()
                basename = make_output_stem(
                    file_path,
                    add_timestamp=True,
                    description="fromXlsx",
                    timestamp_override=timestamp,
                )
                logger.debug(f"统一basename: {basename}")

                # 提取原始文件名（不含扩展名、时间戳、描述）
                original_file_basename = Path(file_path).stem
                original_file_basename = strip_timestamp_suffix(original_file_basename)
                logger.debug(f"原始文件basename: {original_file_basename}")

                # 从basename中提取时间戳和描述部分
                unified_timestamp_desc = f"{timestamp}_fromXlsx"

                logger.debug(f"统一时间戳描述: {unified_timestamp_desc}")

                # 步骤3：创建临时子文件夹
                temp_output_folder = str(Path(temp_dir) / basename)
                Path(temp_output_folder).mkdir(parents=True, exist_ok=True)
                logger.debug(f"创建临时子文件夹: {temp_output_folder}")

                # 步骤4：转换副本，输出到临时子文件夹
                if progress_callback:
                    progress_callback(t("conversion.progress.converting_to_format", format="CSV"))

                csv_files = xlsx_to_csv(
                    temp_input,
                    output_dir=temp_output_folder,
                    original_basename=original_file_basename,
                    unified_timestamp_desc=unified_timestamp_desc,
                )

                if cancel_event and cancel_event.is_set():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.operation_cancelled"),
                        error_code=ERROR_CODE_OPERATION_CANCELLED,
                    )

                if not csv_files:
                    msg = t("conversion.messages.no_csv_generated")
                    return ConversionResult(
                        success=False, message=msg, error_code=ERROR_CODE_CONVERSION_FAILED, details=msg
                    )

                logger.info(f"已生成 {len(csv_files)} 个CSV文件到临时子文件夹")

                # 步骤5：移动整个文件夹到输出目录
                final_folder = str(Path(output_dir) / basename)
                from docwen.utils.workspace_manager import save_output_with_fallback

                saved_folder, _ = save_output_with_fallback(
                    temp_output_folder, final_folder, original_input_file=file_path
                )
                if not saved_folder:
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.conversion_failed"),
                        error_code=ERROR_CODE_CONVERSION_FAILED,
                    )
                final_folder = saved_folder
                logger.info(f"已移动文件夹到: {final_folder}")

                # 准备返回路径（第一个CSV文件的完整路径）
                first_csv_name = Path(csv_files[0]).name
                output_path = str(Path(final_folder) / first_csv_name)

            if progress_callback:
                progress_callback(t("conversion.progress.csv_completed", count=len(csv_files)))

            return ConversionResult(
                success=True,
                output_path=output_path,
                message=t("conversion.messages.conversion_to_format_success", format="CSV"),
            )

        except Exception as e:
            logger.error(f"XLSX转CSV失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"转换失败: {e}",
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )
