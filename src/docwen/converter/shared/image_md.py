from __future__ import annotations

import logging
from collections.abc import Callable
from pathlib import Path

from docwen.utils.markdown_utils import format_image_link

logger = logging.getLogger(__name__)


def create_image_md_file(
    image_path: str,
    image_filename: str,
    output_folder: str,
    include_image: bool,
    include_ocr: bool,
    progress_callback: Callable[[str], None] | None = None,
    current_index: int = 1,
    total_images: int = 1,
    cancel_event=None,
    extraction_mode: str = "file",
    *,
    image_link_style: str = "wiki_embed",
    ocr_progress_message_factory: Callable[[int, int], str] | None = None,
) -> str:
    """
    创建图片的 markdown 文件（与图片同名，扩展名改为 .md）

    文件内容可选包含：
    - 图片链接
    - OCR 文本
    """
    base_name = Path(image_filename).stem
    from docwen.utils.path_utils import generate_named_output_path

    md_path = generate_named_output_path(
        output_dir=output_folder,
        base_name=base_name,
        file_type="md",
        add_timestamp=False,
    )
    md_filename = Path(md_path).name

    try:
        lines: list[str] = []

        if include_image:
            if extraction_mode == "base64":
                from docwen.converter.shared.data_uri_image import build_base64_image_link

                lines.append(build_base64_image_link(image_path, image_link_style))
            else:
                lines.append(format_image_link(image_filename, image_link_style))
            lines.append("")

        if include_ocr:
            if cancel_event and cancel_event.is_set():
                logger.info("OCR识别被取消")
                return image_filename

            if progress_callback:
                if ocr_progress_message_factory is not None:
                    progress_callback(ocr_progress_message_factory(current_index, total_images))
                else:
                    progress_callback(f"OCR {current_index}/{total_images}")

            logger.info(f"开始OCR识别: {image_filename}")
            from docwen.utils.ocr_utils import extract_text_simple

            ocr_text = extract_text_simple(image_path, cancel_event)
            if ocr_text:
                lines.append(ocr_text)
                logger.info(f"OCR识别完成: {image_filename}, 识别出 {len(ocr_text)} 个字符")

        with Path(md_path).open("w", encoding="utf-8") as f:
            f.write("\n".join(lines))

        logger.info(f"创建图片md文件: {md_filename}")
        return md_filename

    except Exception as e:
        logger.error(f"创建图片md文件失败: {md_filename}, 错误: {e}", exc_info=True)
        return image_filename


def process_image_with_ocr(
    img: dict,
    keep_images: bool,
    enable_ocr: bool,
    output_folder: str,
    progress_callback: Callable[[str], None] | None = None,
    current_index: int = 1,
    total_images: int = 1,
    cancel_event=None,
    extraction_mode: str = "file",
    ocr_placement_mode: str = "image_md",
    *,
    image_link_style: str = "wiki_embed",
    md_file_link_style: str = "wiki_embed",
    ocr_blockquote_title: str = "",
    ocr_progress_message_factory: Callable[[int, int], str] | None = None,
) -> str:
    """
    根据选项处理图片，返回应写入主 Markdown 的链接文本。
    """
    from docwen.utils.markdown_utils import (
        format_image_link,
        format_md_file_link,
    )

    filename = img["filename"]
    image_path = img["image_path"]

    logger.debug(
        f"处理图片: {filename}, 保留图片: {keep_images}, OCR: {enable_ocr}, 提取方式: {extraction_mode}, OCR位置: {ocr_placement_mode}"
    )

    # 组合 1: 关 关
    if not keep_images and not enable_ocr:
        logger.info(f"场景：忽略图片 - {filename}")
        return ""

    def _run_ocr() -> str:
        if cancel_event and cancel_event.is_set():
            logger.info("OCR识别被取消")
            return ""

        if progress_callback:
            if ocr_progress_message_factory is not None:
                progress_callback(ocr_progress_message_factory(current_index, total_images))
            else:
                progress_callback(f"OCR {current_index}/{total_images}")

        logger.info(f"开始OCR识别: {filename}")
        from docwen.utils.ocr_utils import extract_text_simple

        ocr_text = extract_text_simple(image_path, cancel_event)
        if ocr_text:
            logger.info(f"OCR识别完成: {filename}, 识别出 {len(ocr_text)} 个字符")
        return ocr_text or ""

    def _format_ocr_blockquote(text: str) -> str:
        if not text:
            return ""
        lines = []
        if ocr_blockquote_title:
            lines.append(f"> {ocr_blockquote_title}")
        for line in text.splitlines():
            lines.append(f"> {line}")
        return "\n".join(lines)

    def _get_image_link() -> str:
        if extraction_mode == "base64":
            from docwen.converter.shared.data_uri_image import build_base64_image_link

            return build_base64_image_link(image_path, image_link_style)
        else:
            return format_image_link(filename, image_link_style)

    # 组合 2 & 3: 开 关 (file / base64)
    if keep_images and not enable_ocr:
        logger.info(f"场景：只保留图片({extraction_mode}) - {filename}")
        return _get_image_link()

    # 组合 8 & 9: 关 开 (main_md / image_md)
    if not keep_images and enable_ocr:
        if ocr_placement_mode == "main_md":
            logger.info(f"场景：只OCR识别(主文档) - {filename}")
            ocr_text = _run_ocr()
            return _format_ocr_blockquote(ocr_text) if ocr_text else ""
        else:
            logger.info(f"场景：只OCR识别(图片对应MD) - {filename}")
            md_filename = create_image_md_file(
                image_path,
                filename,
                output_folder,
                include_image=False,
                include_ocr=True,
                progress_callback=progress_callback,
                current_index=current_index,
                total_images=total_images,
                cancel_event=cancel_event,
                extraction_mode=extraction_mode,
                image_link_style=image_link_style,
                ocr_progress_message_factory=ocr_progress_message_factory,
            )
            return format_md_file_link(md_filename, md_file_link_style)

    # 组合 4-7: 开 开
    if keep_images and enable_ocr:
        if ocr_placement_mode == "main_md":
            logger.info(f"场景：图片 + OCR(主文档, {extraction_mode}) - {filename}")
            img_link = _get_image_link()
            ocr_text = _run_ocr()
            if ocr_text:
                return f"{img_link}\n\n{_format_ocr_blockquote(ocr_text)}"
            return img_link
        else:
            logger.info(f"场景：图片 + OCR(图片对应MD, {extraction_mode}) - {filename}")
            md_filename = create_image_md_file(
                image_path,
                filename,
                output_folder,
                include_image=True,
                include_ocr=True,
                progress_callback=progress_callback,
                current_index=current_index,
                total_images=total_images,
                cancel_event=cancel_event,
                extraction_mode=extraction_mode,
                image_link_style=image_link_style,
                ocr_progress_message_factory=ocr_progress_message_factory,
            )
            return format_md_file_link(md_filename, md_file_link_style)

    return ""
