"""
图片转Markdown策略模块

将图片文件转换为Markdown格式，支持OCR文字识别。

依赖：
- PIL: 图片处理
- ocr_utils: OCR识别
- yaml_utils: YAML头部生成
"""

import logging
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from docwen.translation import t
from docwen.services.error_codes import (
    ERROR_CODE_CONVERSION_FAILED,
    ERROR_CODE_INVALID_INPUT,
    ERROR_CODE_OPERATION_CANCELLED,
)
from docwen.services.result import ConversionResult
from docwen.utils.path_utils import generate_output_path
from docwen.utils.validation_utils import validate_ocr_requires_images
from docwen.utils.yaml_utils import generate_basic_yaml_frontmatter

from .. import CATEGORY_IMAGE, register_conversion
from ..base_strategy import BaseStrategy
from .utils import extract_tiff_pages, get_image_format_description, is_multipage_tiff

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_IMAGE, "md")
class ImageToMarkdownStrategy(BaseStrategy):
    """
    将图片文件转换为Markdown格式的策略（含OCR）

    输出结构：
    文件夹名_时间戳_fromFormat/
    ├── 图片文件（原始或转换后）
    └── 文件夹名_时间戳_fromFormat.md

    MD格式：
    ![图片文件名](./图片文件名)

    OCR识别内容

    多页TIFF会拆分为多个PNG，每个PNG对应一个图片链接和OCR代码块
    """

    def execute(
        self,
        file_path: str,
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行图片到Markdown的转换

        Args:
            file_path: 输入图片文件路径
            options: 转换选项，包含actual_format
            progress_callback: 进度更新回调函数

        Returns:
            ConversionResult: 转换结果对象
        """
        try:
            actual_format = options.get("actual_format") if options else None
            if not actual_format:
                try:
                    from docwen.utils.file_type_utils import detect_actual_file_format

                    actual_format = detect_actual_file_format(file_path)
                except Exception:
                    actual_format = None
            if not actual_format:
                actual_format = Path(file_path).suffix.lstrip(".").lower() or "jpg"
            actual_format = actual_format.lower()

            from docwen.utils.workspace_manager import get_output_directory

            output_dir = get_output_directory(file_path)

            original_basename = Path(file_path).stem

            with tempfile.TemporaryDirectory() as temp_dir:
                # === 步骤1：创建输入副本 input.{ext} ===
                if progress_callback:
                    progress_callback(t("conversion.progress.preparing"))

                from docwen.utils.workspace_manager import prepare_input_file

                temp_image = prepare_input_file(file_path, temp_dir, actual_format or "jpg")
                logger.debug(f"已创建输入副本: {Path(temp_image).name}")

                # === 步骤2：生成统一的basename ===
                description = get_image_format_description(actual_format) if actual_format else "fromImage"
                output_path_with_ext = generate_output_path(file_path, None, "", True, description, "md")
                basename = Path(output_path_with_ext).stem
                logger.info(f"统一basename: {basename}")

                # === 步骤3：创建子文件夹 ===
                from docwen.utils.path_utils import ensure_unique_directory_path

                final_folder = ensure_unique_directory_path(str(Path(output_dir) / basename))
                final_basename = Path(final_folder).name

                folder_path = Path(temp_dir) / final_basename
                folder_path.mkdir(parents=True, exist_ok=True)
                logger.debug(f"子文件夹: {folder_path}")

                # === 步骤4：处理图片到子文件夹 ===
                images_in_folder = []  # [(图片路径, 图片文件名), ...]

                if is_multipage_tiff(temp_image):
                    # 情况1：多页TIFF - 使用标准化命名直接提取到子文件夹
                    logger.info("检测到多页TIFF，使用标准化命名拆分")
                    pages = extract_tiff_pages(
                        file_path=temp_image,
                        output_dir=str(folder_path),
                        actual_format=actual_format or "tiff",
                        progress_callback=progress_callback,
                    )
                    images_in_folder = [(path, Path(path).name) for _, path in pages]
                    logger.debug(f"多页TIFF拆分完成，共 {len(pages)} 页")

                elif actual_format in ["heic", "heif", "bmp"]:
                    # 情况2：需要转换格式 - 直接转换到子文件夹
                    if progress_callback:
                        progress_callback(
                            t("conversion.progress.converting_format_to_png", format=actual_format.upper())
                        )

                    logger.info(f"{actual_format.upper()}转换为PNG")
                    image_filename = f"{original_basename}.png"
                    final_image_path = str(folder_path / image_filename)

                    from docwen.converter.formats.image import bmp_to_png, heic_to_png

                    if actual_format in ["heic", "heif"]:
                        heic_to_png(temp_image, output_path=final_image_path)
                    else:  # bmp
                        bmp_to_png(temp_image, output_path=final_image_path)

                    images_in_folder = [(final_image_path, image_filename)]

                else:
                    # 情况3：标准格式 - 移动到子文件夹
                    if progress_callback:
                        progress_callback(t("conversion.progress.processing_images"))

                    logger.info(f"标准{actual_format.upper()}格式")
                    image_filename = f"{original_basename}.{actual_format}"
                    final_image_path = str(folder_path / image_filename)
                    from docwen.utils.workspace_manager import move_file_with_retry

                    moved_path = move_file_with_retry(temp_image, final_image_path)
                    if not moved_path:
                        return ConversionResult(
                            success=False,
                            message=t("conversion.messages.conversion_failed"),
                            error_code=ERROR_CODE_CONVERSION_FAILED,
                        )
                    final_image_path = moved_path

                    images_in_folder = [(final_image_path, image_filename)]

                # === 步骤5：统一OCR子文件夹的所有图片并生成MD内容 ===
                from docwen.config.config_manager import config_manager

                options = options or {}
                optimize_for_type = options.get("optimize_for_type")
                options.setdefault(
                    "to_md_image_extraction_mode",
                    config_manager.get_image_to_md_image_extraction_mode(),
                )
                options.setdefault(
                    "to_md_ocr_placement_mode",
                    config_manager.get_image_to_md_ocr_placement_mode(),
                )

                extraction_mode = options.get("to_md_image_extraction_mode", "file")
                ocr_placement_mode = options.get("to_md_ocr_placement_mode", "image_md")

                extract_image = bool(
                    options.get(
                        "extract_image",
                        config_manager.get_image_to_md_keep_images(),
                    )
                )
                extract_ocr = bool(options.get("extract_ocr", True))

                # 计算总图片数（用于进度显示）
                total_images = len(images_in_folder)

                # 获取cancel_event（如果有）
                cancel_event = options.get("cancel_event")

                if optimize_for_type == "invoice_cn":
                    from docwen.converter.layout2md.invoice_cn_ocr import (
                        build_invoice_md_text,
                        parse_invoice_from_image,
                    )

                    def _insert_section_before_timestamp(stem: str, section: str) -> str:
                        import re

                        m = re.search(r"(_\d{8}_\d{6})(.*)$", stem)
                        if not m:
                            return f"{stem}_{section}"
                        prefix = stem[: m.start()]
                        ts = m.group(1)
                        suffix = m.group(2)
                        return f"{prefix}_{section}{ts}{suffix}"

                    md_filenames: list[str] = []
                    for idx, (img_path, _img_filename) in enumerate(images_in_folder, 1):
                        if cancel_event and cancel_event.is_set():
                            logger.info("用户取消操作，停止OCR识别")
                            return ConversionResult(
                                success=False,
                                message=t("conversion.messages.operation_cancelled"),
                                error_code=ERROR_CODE_OPERATION_CANCELLED,
                            )

                        section = t("conversion.filenames.page_n", n=idx) if total_images > 1 else ""
                        file_stem = f"{original_basename}_{section}" if section else original_basename
                        output_stem = (
                            _insert_section_before_timestamp(final_basename, section) if section else final_basename
                        )

                        metadata, rows = parse_invoice_from_image(img_path, cancel_event=cancel_event)
                        md_text = build_invoice_md_text(
                            file_stem=file_stem, metadata=metadata, rows=rows, include_empty=True
                        )

                        md_filename = f"{output_stem}.md"
                        (folder_path / md_filename).write_text(md_text, encoding="utf-8")
                        md_filenames.append(md_filename)

                    if extraction_mode == "base64" or not extract_image:
                        from contextlib import suppress

                        for img_path, _ in images_in_folder:
                            with suppress(Exception):
                                Path(img_path).unlink(missing_ok=True)

                    if progress_callback:
                        progress_callback(t("conversion.progress.writing_file"))

                    from docwen.utils.workspace_manager import save_output_with_fallback

                    saved_folder, _ = save_output_with_fallback(
                        str(folder_path),
                        final_folder,
                        original_input_file=file_path,
                    )
                    if not saved_folder:
                        return ConversionResult(
                            success=False,
                            message=t("conversion.messages.conversion_failed"),
                            error_code=ERROR_CODE_CONVERSION_FAILED,
                        )
                    final_folder = saved_folder

                    message = (
                        t("conversion.messages.invoice_md_multi_page_success", count=total_images)
                        if total_images > 1
                        else t("conversion.messages.conversion_to_format_success", format="Markdown")
                    )

                    return ConversionResult(
                        success=True,
                        output_path=str(Path(final_folder) / md_filenames[0]),
                        message=message,
                    )

                ok, reason = validate_ocr_requires_images(extract_image, extract_ocr)
                if not ok:
                    return ConversionResult(
                        success=False,
                        message=reason,
                        error_code=ERROR_CODE_INVALID_INPUT,
                        details=reason,
                    )

                logger.info(f"导出选项 - 提取图片: {extract_image}, OCR: {extract_ocr}")

                md_content = generate_basic_yaml_frontmatter(original_basename)

                from docwen.converter.shared.image_md import process_image_with_ocr

                link_settings = config_manager.get_markdown_link_style_settings()
                image_link_style = link_settings.get("image_link_style", "wiki_embed")
                md_file_link_style = link_settings.get("md_file_link_style", "wiki_embed")

                ocr_title = ""
                if config_manager.get_ocr_blockquote_title_enabled():
                    override = config_manager.get_ocr_blockquote_title_override_text()
                    if override and override.strip():
                        ocr_title = override.strip()
                    else:
                        ocr_title = str(t("conversion.ocr_output.blockquote_prefix", default="") or "")

                def _ocr_progress(current: int, total: int) -> str:
                    return t("conversion.progress.ocr_recognizing", current=current, total=total)

                for idx, (img_path, img_filename) in enumerate(images_in_folder, 1):
                    # 检查取消事件（在处理每张图片前）
                    if cancel_event and cancel_event.is_set():
                        logger.info("用户取消操作，停止OCR识别")
                        return ConversionResult(
                            success=False,
                            message=t("conversion.messages.operation_cancelled"),
                            error_code=ERROR_CODE_OPERATION_CANCELLED,
                        )

                    img_info = {"filename": img_filename, "image_path": img_path}
                    link_text = process_image_with_ocr(
                        img=img_info,
                        keep_images=extract_image,
                        enable_ocr=extract_ocr,
                        output_folder=str(folder_path),
                        progress_callback=progress_callback,
                        current_index=idx,
                        total_images=total_images,
                        cancel_event=cancel_event,
                        extraction_mode=extraction_mode,
                        ocr_placement_mode=ocr_placement_mode,
                        image_link_style=image_link_style,
                        md_file_link_style=md_file_link_style,
                        ocr_blockquote_title=ocr_title,
                        ocr_progress_message_factory=_ocr_progress,
                    )
                    if link_text:
                        md_content += f"{link_text}\n\n"

                if extraction_mode == "base64" or not extract_image:
                    from contextlib import suppress

                    for img_path, _ in images_in_folder:
                        with suppress(Exception):
                            Path(img_path).unlink(missing_ok=True)

                # === 步骤6：保存MD文件 ===
                if progress_callback:
                    progress_callback(t("conversion.progress.writing_file"))

                md_filename = f"{final_basename}.md"
                md_path = folder_path / md_filename
                with md_path.open("w", encoding="utf-8") as f:
                    f.write(md_content)
                logger.info(f"MD文件已生成: {md_path}")

                # === 步骤7：移动子文件夹到目标目录 ===
                from docwen.utils.workspace_manager import save_output_with_fallback

                saved_folder, _ = save_output_with_fallback(
                    str(folder_path), final_folder, original_input_file=file_path
                )
                if not saved_folder:
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.conversion_failed"),
                        error_code=ERROR_CODE_CONVERSION_FAILED,
                    )
                final_folder = saved_folder
                logger.info(f"子文件夹已移动到: {final_folder}")

                return ConversionResult(
                    success=True,
                    output_path=str(Path(final_folder) / md_filename),
                    message=t("conversion.messages.conversion_to_format_success", format="Markdown"),
                )

        except Exception as e:
            logger.error(f"图片转Markdown失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=t("conversion.messages.conversion_failed_check_log"),
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )
