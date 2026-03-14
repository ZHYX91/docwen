"""
文档转Markdown策略模块

将文档文件转换为Markdown格式。

依赖：
- docx2md: 文档转MD核心转换
- utils: 文档预处理
"""

import logging
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from docwen.config import config_manager
from docwen.converter.docx2md.core import convert_docx_to_md
from docwen.services.error_codes import (
    ERROR_CODE_CONVERSION_FAILED,
    ERROR_CODE_INVALID_INPUT,
    ERROR_CODE_OPERATION_CANCELLED,
    ERROR_CODE_UNKNOWN_ERROR,
)
from docwen.services.result import ConversionResult
from docwen.translation import t
from docwen.utils.path_utils import generate_output_path
from docwen.utils.validation_utils import validate_ocr_requires_images

from .. import CATEGORY_DOCUMENT, register_conversion
from ..base_strategy import BaseStrategy
from .utils import preprocess_document_file

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_DOCUMENT, "md")
class DocxToMdStrategy(BaseStrategy):
    """
    将DOCX文件转换为Markdown文件的策略。

    功能特性：
    - 自动识别主要部分和附件部分
    - 主要部分和附件部分分别生成独立的Markdown文件
    - 支持取消操作
    - 带YAML元数据头部
    - **支持DOC/WPS/RTF格式自动转换**

    支持的输入格式：
    - DOCX：直接处理
    - DOC/WPS/RTF：自动转换为DOCX后处理

    输出文件：
    - 主要部分：文件名标记为 "主要部分"
    - 附件部分（如果存在）：文件名标记为 "附件部分"
    """

    def execute(
        self,
        file_path: str,
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行DOCX到Markdown的转换（支持DOC/WPS/RTF自动转换）。

        Args:
            file_path: 输入的文档文件路径（DOCX/DOC/WPS/RTF）
            options: 转换选项字典，包含：
                - cancel_event: (可选) 用于取消操作的事件对象
            progress_callback: 进度更新回调函数

        Returns:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}

        # 调试日志：查看收到的options
        logger.info(f"DocxToMdStrategy收到options: {options}")
        logger.info(f"  extract_image: {options.get('extract_image')}")
        logger.info(f"  extract_ocr: {options.get('extract_ocr')}")

        cancel_event = options.get("cancel_event")
        actual_format = options.get("actual_format")

        if progress_callback:
            progress_callback(t("conversion.progress.preparing"))

        try:
            from docwen.utils.workspace_manager import get_output_directory

            output_dir = get_output_directory(file_path)

            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：预处理 - 将DOC/WPS/RTF转换为DOCX（如需要）
                if progress_callback:
                    progress_callback(t("conversion.progress.detecting_format"))

                preprocess_result = preprocess_document_file(file_path, temp_dir, cancel_event, actual_format)
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
                    file_path,
                    output_dir=output_dir,
                    section="",
                    add_timestamp=True,
                    description=description,
                    file_type="md",
                )

                basename = Path(base_path).stem

                # 创建临时子文件夹
                temp_output_folder = Path(temp_dir) / basename
                temp_output_folder.mkdir(parents=True, exist_ok=True)
                logger.debug(f"创建临时子文件夹: {temp_output_folder}")

                # 步骤3：从GUI获取导出选项
                extract_image = bool(options.get("extract_image", True))
                extract_ocr = options.get("extract_ocr", False)
                optimize_for_type = options.get("optimize_for_type") or ""

                options.setdefault(
                    "to_md_image_extraction_mode",
                    config_manager.get_docx_to_md_image_extraction_mode(),
                )
                options.setdefault(
                    "to_md_ocr_placement_mode",
                    config_manager.get_docx_to_md_ocr_placement_mode(),
                )

                ok, reason = validate_ocr_requires_images(extract_image, extract_ocr)
                if not ok:
                    return ConversionResult(
                        success=False,
                        message=reason,
                        error_code=ERROR_CODE_INVALID_INPUT,
                        details=reason,
                    )

                logger.info(
                    f"从options提取参数 - extract_image: {extract_image}, extract_ocr: {extract_ocr}, optimize_for_type: {optimize_for_type}"
                )

                # 步骤4：调用核心转换函数
                result = convert_docx_to_md(
                    processed_file,
                    extract_image=extract_image,
                    extract_ocr=extract_ocr,
                    optimize_for_type=optimize_for_type,
                    progress_callback=progress_callback,
                    cancel_event=cancel_event,
                    output_folder=str(temp_output_folder),
                    original_file_path=file_path,
                    options=options,
                )

                if cancel_event and cancel_event.is_set():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.operation_cancelled"),
                        error_code=ERROR_CODE_OPERATION_CANCELLED,
                    )

                if not result["success"]:
                    details = str(result.get("error") or "")
                    return ConversionResult(
                        success=False,
                        message=f"DOCX到MD转换失败: {details}",
                        error_code=ERROR_CODE_CONVERSION_FAILED,
                        details=details or None,
                    )

                # 步骤5：写入文件
                if progress_callback:
                    progress_callback(t("conversion.progress.writing_file"))

                main_filename = f"{basename}.md"
                temp_main_output = temp_output_folder / main_filename
                with temp_main_output.open("w", encoding="utf-8") as f:
                    f.write(result["main_content"])

                logger.info(f"主要部分文件已写入: {temp_main_output}")

                # 步骤6：如果有附件内容，写入附件MD
                if result["attachment_content"]:
                    original_basename = Path(file_path).stem

                    parts = basename.split("_")
                    if len(parts) >= 3:
                        timestamp_idx = None
                        for i, part in enumerate(parts):
                            if len(part) == 8 and part.isdigit():
                                timestamp_idx = i
                                break

                        if timestamp_idx is not None:
                            attachment_filename = f"{original_basename}_附件部分_{'_'.join(parts[timestamp_idx:])}.md"
                        else:
                            attachment_filename = f"{basename}_附件部分.md"
                    else:
                        attachment_filename = f"{basename}_附件部分.md"

                    temp_attachment_output = temp_output_folder / attachment_filename
                    with temp_attachment_output.open("w", encoding="utf-8") as f:
                        f.write(result["attachment_content"])

                    logger.info(f"附件文件已写入: {temp_attachment_output}")

                # 步骤7：移动整个文件夹到输出目录
                final_folder = str(Path(output_dir) / basename)
                from docwen.utils.workspace_manager import save_output_with_fallback

                saved_folder, _ = save_output_with_fallback(
                    str(temp_output_folder), final_folder, original_input_file=file_path
                )
                if not saved_folder:
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.conversion_failed"),
                        error_code=ERROR_CODE_CONVERSION_FAILED,
                    )
                final_folder = saved_folder
                logger.info(f"已移动文件夹到: {final_folder}")

                # 步骤8：如果保留中间文件，移动temp_dir中的其他规范文件
                should_keep = self._should_keep_intermediates()
                saved_intermediate_items = []
                if should_keep:
                    from docwen.utils.workspace_manager import save_intermediate_items

                    saved_intermediate_items = save_intermediate_items(
                        preprocess_result.intermediates, output_dir, move=True
                    )
                    for _, saved_path in saved_intermediate_items:
                        logger.info(f"保留中间文件: {Path(saved_path).name}")

                main_output_path = str(Path(final_folder) / main_filename)

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
                            output_path=main_output_path,
                            mask_input=config_manager.get_mask_manifest_input_path(),
                        )
                        write_manifest_json(final_folder, manifest)
                except Exception:
                    pass

            return ConversionResult(
                success=True,
                output_path=main_output_path,
                message=t("conversion.messages.conversion_to_format_success", format="Markdown"),
            )

        except Exception as e:
            logger.error(f"执行 DocxToMdStrategy 时出错: {e}", exc_info=True)
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
                message=f"发生未知错误: {e}",
                error=e,
                error_code=ERROR_CODE_UNKNOWN_ERROR,
                details=str(e) or None,
            )

    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件"""
        try:
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
            return False
