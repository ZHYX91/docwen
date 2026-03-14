"""
XLSX图片提取模块

从XLSX表格文件中提取内嵌图片并保存到指定文件夹
与DOCX的image_processor模块保持一致的设计模式
"""

import logging
from collections.abc import Callable
from pathlib import Path

from docwen.translation import t
from docwen.utils.path_utils import generate_output_path

logger = logging.getLogger(__name__)


def extract_images_from_xlsx(
    workbook, output_folder: str, original_file_path: str, progress_callback: Callable[[str], None] | None = None
) -> list[dict]:
    """
    从XLSX工作簿对象中提取所有图片

    设计说明：
        此函数采用与文档图片提取一致的设计模式，接收已加载的工作簿对象，
        实现了职责分离（文件加载 vs 图片提取）和性能优化（避免重复I/O）。

    参数:
        workbook: openpyxl的Workbook对象（已加载）
        output_folder: 图片保存的文件夹路径
        original_file_path: 原始文件路径（用于生成规范的图片文件名）
        progress_callback: 进度回调函数（可选）

    返回:
        List[dict]: 图片信息列表，每个元素包含：
            - filename: 图片文件名（如 '销售数据_image1_20250125_091530_fromXlsx.png'）
            - image_path: 图片完整路径
            - sheet_name: 所在工作表名称
            - row: 图片所在行号（可选）
            - col: 图片所在列号（可选）
    """
    images_info = []
    image_counter = 1

    logger.info(f"开始从XLSX工作簿提取图片到: {output_folder}")

    # 确保输出文件夹存在
    Path(output_folder).mkdir(parents=True, exist_ok=True)

    try:
        # 使用传入的workbook对象，避免重复加载
        wb = workbook

        # 遍历所有工作表
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            # 检查工作表中的图片
            if hasattr(ws, "_images") and ws._images:
                logger.debug(f"工作表 '{sheet_name}' 包含 {len(ws._images)} 张图片")

                for img in ws._images:
                    try:
                        # 获取图片数据
                        image_data = img._data()

                        # 确定图片格式
                        # openpyxl的图片对象可能有format属性
                        ext = "png"  # 默认格式
                        if hasattr(img, "format"):
                            ext = img.format.lower()
                        elif hasattr(img, "_format"):
                            ext = img._format.lower()

                        # 使用统一的路径生成工具生成规范文件名
                        # 使用original_file_path进行命名，确保文件名基于原始文件而非临时副本
                        image_path = generate_output_path(
                            input_path=original_file_path,
                            output_dir=output_folder,
                            section=f"image{image_counter}",
                            add_timestamp=True,
                            description="fromXlsx",
                            file_type=ext,
                        )
                        filename = Path(image_path).name

                        # 保存图片
                        with Path(image_path).open("wb") as f:
                            f.write(image_data)

                        # 获取图片位置
                        row, col = None, None
                        if hasattr(img, "anchor") and img.anchor and hasattr(img.anchor, "_from"):
                            row = img.anchor._from.row + 1  # openpyxl从0开始，转为从1开始
                            col = img.anchor._from.col + 1

                        logger.debug(f"提取图片: {filename} (工作表 '{sheet_name}', 行{row}, 列{col})")

                        # 记录图片信息（包括位置）
                        img_info = {"filename": filename, "image_path": image_path, "sheet_name": sheet_name}
                        if row is not None and col is not None:
                            img_info["row"] = row
                            img_info["col"] = col

                        images_info.append(img_info)

                        # 进度回调
                        if progress_callback:
                            progress_callback(t("conversion.progress.extracting_images", count=image_counter))

                        image_counter += 1

                    except Exception as e:
                        logger.warning(f"提取图片失败 (工作表 '{sheet_name}'): {e}")
                        continue

        logger.info(f"XLSX图片提取完成，共提取 {len(images_info)} 张图片")

    except Exception as e:
        logger.error(f"提取XLSX图片时出错: {e}", exc_info=True)

    return images_info


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
) -> str:
    """
    根据选项处理图片，生成对应的Markdown链接（与文档转MD逻辑一致）

    参数:
        img: 图片信息字典，包含 'filename', 'image_path' 等
        keep_images: 是否保留图片
        enable_ocr: 是否启用OCR识别
        output_folder: 输出文件夹路径
        progress_callback: 进度回调函数（可选）

    返回:
        str: Markdown链接（图片链接或图片md文件链接）
    """
    from docwen.config.config_manager import config_manager
    from docwen.converter.shared.image_md import process_image_with_ocr as shared_process_image_with_ocr
    from docwen.translation import t

    link_settings = config_manager.get_markdown_link_style_settings()
    image_link_style = link_settings.get("image_link_style", "wiki_embed")
    md_file_link_style = link_settings.get("md_file_link_style", "wiki_embed")

    title = ""
    if config_manager.get_ocr_blockquote_title_enabled():
        override = config_manager.get_ocr_blockquote_title_override_text()
        if override and override.strip():
            title = override.strip()
        else:
            title = str(t("conversion.ocr_output.blockquote_prefix", default="") or "")

    def _ocr_progress(current: int, total: int) -> str:
        return t("conversion.progress.ocr_recognizing", current=current, total=total)

    return shared_process_image_with_ocr(
        img=img,
        keep_images=keep_images,
        enable_ocr=enable_ocr,
        output_folder=output_folder,
        progress_callback=progress_callback,
        current_index=current_index,
        total_images=total_images,
        cancel_event=cancel_event,
        extraction_mode=extraction_mode,
        ocr_placement_mode=ocr_placement_mode,
        image_link_style=image_link_style,
        md_file_link_style=md_file_link_style,
        ocr_blockquote_title=title,
        ocr_progress_message_factory=_ocr_progress,
    )


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
) -> str:
    """
    创建图片的markdown文件（与文档转MD逻辑一致）

    参数:
        image_path: 图片文件的完整路径
        image_filename: 图片文件名（如 'image_1.png'）
        output_folder: 输出文件夹路径
        include_image: 是否在md文件中包含图片链接
        include_ocr: 是否在md文件中包含OCR识别文本
        progress_callback: 进度回调函数（可选）

    返回:
        str: 创建的md文件名（如 'image_1.md'）
    """
    from docwen.config.config_manager import config_manager
    from docwen.converter.shared.image_md import create_image_md_file as shared_create_image_md_file
    from docwen.translation import t

    link_settings = config_manager.get_markdown_link_style_settings()
    image_link_style = link_settings.get("image_link_style", "wiki_embed")

    def _ocr_progress(current: int, total: int) -> str:
        return t("conversion.progress.ocr_recognizing", current=current, total=total)

    return shared_create_image_md_file(
        image_path=image_path,
        image_filename=image_filename,
        output_folder=output_folder,
        include_image=include_image,
        include_ocr=include_ocr,
        progress_callback=progress_callback,
        current_index=current_index,
        total_images=total_images,
        cancel_event=cancel_event,
        image_link_style=image_link_style,
        ocr_progress_message_factory=_ocr_progress,
    )


def replace_image_markers(
    markdown_text: str,
    sheet_images: list[dict],
    keep_images: bool,
    enable_ocr: bool,
    extraction_mode: str,
    ocr_placement_mode: str,
    output_folder: str,
    progress_callback: Callable[[str], None] | None = None,
    cancel_event=None,
) -> str:
    """
    替换Markdown文本中的图片标记为实际的图片链接

    参数:
        markdown_text: 包含图片标记的Markdown文本
        sheet_images: 工作表的图片信息列表
        keep_images: 是否保留图片
        enable_ocr: 是否启用OCR
        output_folder: 输出文件夹路径
        progress_callback: 进度回调函数（可选）

    返回:
        str: 替换后的Markdown文本
    """

    # 计算需要OCR的图片总数
    total_images = len(sheet_images)
    current_index = 0

    # 查找所有图片标记: {{IMAGE:filename}}
    for img_info in sheet_images:
        # 检查取消（在处理每张图片前）
        if cancel_event and cancel_event.is_set():
            logger.info(f"图片标记替换被取消，已处理 {current_index}/{total_images} 张")
            break

        filename = img_info.get("filename")
        if not filename:
            continue

        # 创建标记模式
        marker = f"{{{{IMAGE:{filename}}}}}"

        # 如果Markdown中包含这个标记，替换为实际链接
        if marker in markdown_text:
            # 增加当前索引
            current_index += 1

            # 处理图片时传递progress_callback、索引信息和cancel_event
            image_link = process_image_with_ocr(
                img_info,
                keep_images,
                enable_ocr,
                output_folder,
                progress_callback,
                current_index,
                total_images,
                cancel_event,
                extraction_mode=extraction_mode,
                ocr_placement_mode="image_md" if ocr_placement_mode == "main_md" else ocr_placement_mode,
            )

            markdown_text = markdown_text.replace(marker, image_link)
            logger.debug(f"替换图片标记: {marker} -> {image_link}")

    return markdown_text
