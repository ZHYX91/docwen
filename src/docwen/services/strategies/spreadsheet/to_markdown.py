"""
表格转Markdown/TXT策略模块

将表格文件转换为Markdown或TXT格式。

依赖：
- xlsx2md: 表格转MD核心转换
- utils: 表格预处理
"""

import logging
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from docwen.config import config_manager
from docwen.converter.xlsx2md import convert_spreadsheet_to_md
from docwen.services.error_codes import (
    ERROR_CODE_CONVERSION_FAILED,
    ERROR_CODE_INVALID_INPUT,
    ERROR_CODE_OPERATION_CANCELLED,
)
from docwen.services.result import ConversionResult
from docwen.translation import t
from docwen.utils.path_utils import generate_output_path
from docwen.utils.validation_utils import validate_ocr_requires_images

from .. import CATEGORY_SPREADSHEET, register_conversion
from ..base_strategy import BaseStrategy
from .utils import preprocess_table_file

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_SPREADSHEET, "md")
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
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
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
                progress_callback(t("conversion.progress.converting_to_format", format="Markdown"))

            options = options or {}
            cancel_event = options.get("cancel_event")
            optimize_for_type = options.get("optimize_for_type")
            if optimize_for_type:
                logger.info(f"表格转MD使用优化类型: {optimize_for_type}")
            actual_format = options.get("actual_format")
            if not actual_format:
                try:
                    from docwen.utils.file_type_utils import detect_actual_file_format

                    actual_format = detect_actual_file_format(file_path)
                except Exception:
                    actual_format = None

            # 从GUI获取导出选项
            extract_image = bool(options.get("extract_image", False))
            extract_ocr = options.get("extract_ocr", False)

            options.setdefault(
                "to_md_image_extraction_mode",
                config_manager.get_xlsx_to_md_image_extraction_mode(),
            )
            options.setdefault(
                "to_md_ocr_placement_mode",
                config_manager.get_xlsx_to_md_ocr_placement_mode(),
            )

            ok, reason = validate_ocr_requires_images(extract_image, extract_ocr)
            if not ok:
                return ConversionResult(
                    success=False, message=reason, error_code=ERROR_CODE_INVALID_INPUT, details=reason
                )

            logger.info(f"表格转MD - 导出选项: 提取图片={extract_image}, OCR={extract_ocr}")

            from docwen.utils.workspace_manager import get_output_directory

            output_dir = get_output_directory(file_path)

            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：预处理 - 将XLS/ET转换为XLSX（如需要）
                if progress_callback:
                    progress_callback(t("conversion.progress.detecting_format"))

                preprocess_result = preprocess_table_file(file_path, temp_dir, cancel_event, actual_format)
                processed_file = preprocess_result.processed_file
                actual_format = preprocess_result.actual_format
                options["actual_format"] = actual_format

                if cancel_event and cancel_event.is_set():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.operation_cancelled"),
                        error_code=ERROR_CODE_OPERATION_CANCELLED,
                    )

                # 步骤2：生成统一basename和创建临时子文件夹
                description = f"from{actual_format.capitalize()}"

                base_path = generate_output_path(
                    file_path, section="", add_timestamp=True, description=description, file_type="md"
                )

                basename = Path(base_path).stem
                logger.debug(f"统一basename: {basename}")

                # 创建临时子文件夹
                temp_output_folder = str(Path(temp_dir) / basename)
                Path(temp_output_folder).mkdir(parents=True, exist_ok=True)
                logger.debug(f"创建临时子文件夹: {temp_output_folder}")

                # 步骤3：调用核心转换函数
                markdown_content = convert_spreadsheet_to_md(
                    processed_file,
                    extract_image=extract_image,
                    extract_ocr=extract_ocr,
                    extraction_mode=options.get("to_md_image_extraction_mode"),
                    ocr_placement_mode=options.get("to_md_ocr_placement_mode"),
                    output_folder=temp_output_folder,
                    original_file_path=file_path,
                    progress_callback=progress_callback,
                    cancel_event=cancel_event,
                )

                if cancel_event and cancel_event.is_set():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.operation_cancelled"),
                        error_code=ERROR_CODE_OPERATION_CANCELLED,
                    )

                # 步骤4：写入Markdown文件
                if progress_callback:
                    progress_callback(t("conversion.progress.writing_file"))

                md_filename = f"{basename}.md"
                temp_output = str(Path(temp_output_folder) / md_filename)

                logger.debug(f"准备将Markdown内容写入: {temp_output}")
                with Path(temp_output).open("w", encoding="utf-8") as f:
                    f.write(markdown_content)
                logger.info(f"Markdown内容已写入: {temp_output}")

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

                # 步骤6：如果保留中间文件，移动temp_dir中的其他规范文件
                should_keep = self._should_keep_intermediates()
                saved_intermediate_items = []
                if should_keep:
                    from docwen.utils.workspace_manager import save_intermediate_items

                    saved_intermediate_items = save_intermediate_items(
                        preprocess_result.intermediates, output_dir, move=True
                    )
                    for _, saved_path in saved_intermediate_items:
                        logger.info(f"保留中间文件: {Path(saved_path).name}")

                output_path = str(Path(final_folder) / md_filename)

                try:
                    if config_manager.get_save_manifest():
                        from docwen.utils.workspace_manager import build_manifest, write_manifest_json

                        manifest = build_manifest(
                            file_path=file_path,
                            actual_format=actual_format,
                            preprocess_chain=preprocess_result.preprocess_chain,
                            saved_intermediate_items=saved_intermediate_items,
                            options=options,
                            success=True,
                            message=t("conversion.messages.conversion_to_format_success", format="Markdown"),
                            output_path=output_path,
                            mask_input=config_manager.get_mask_manifest_input_path(),
                        )
                        write_manifest_json(final_folder, manifest)
                except Exception:
                    pass

                return ConversionResult(
                    success=True,
                    output_path=output_path,
                    message=t("conversion.messages.conversion_to_format_success", format="Markdown"),
                )

        except Exception as e:
            logger.error(f"表格转Markdown失败: {e}", exc_info=True)
            try:
                from docwen.utils.workspace_manager import build_manifest, write_manifest_json

                should_keep = False
                try:
                    should_keep = config_manager.get_save_intermediate_files()
                except Exception:
                    should_keep = False

                failure_output_dir = None
                try:
                    from docwen.utils.workspace_manager import get_output_directory

                    failure_output_dir = get_output_directory(file_path)
                except Exception:
                    failure_output_dir = str(Path(file_path).parent)

                saved_failure_items = []
                if should_keep and "preprocess_result" in locals():
                    try:
                        from docwen.utils.workspace_manager import save_intermediate_items

                        saved_failure_items = save_intermediate_items(
                            preprocess_result.intermediates, failure_output_dir, move=True
                        )
                    except Exception:
                        saved_failure_items = []

                if config_manager.get_save_manifest():
                    from datetime import datetime

                    manifest = build_manifest(
                        file_path=file_path,
                        actual_format=locals().get("actual_format"),
                        preprocess_chain=getattr(locals().get("preprocess_result"), "preprocess_chain", None),
                        saved_intermediate_items=saved_failure_items,
                        options=(options or {}),
                        success=False,
                        message=str(e),
                        output_path=None,
                        mask_input=config_manager.get_mask_manifest_input_path(),
                    )
                    write_manifest_json(
                        failure_output_dir,
                        manifest,
                        filename=f"manifest_failed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    )
            except Exception:
                pass
            return ConversionResult(
                success=False,
                message=f"转换失败: {e}",
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )

    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件"""
        try:
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取清理配置失败: {e}，使用默认设置（清理中间文件）")
            return False
