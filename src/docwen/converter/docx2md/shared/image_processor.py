"""
DOCX图片提取模块

从DOCX文档中提取内嵌图片并保存到指定文件夹
"""

import logging
from collections.abc import Callable
from pathlib import Path

from docwen.translation import t
from docwen.utils.path_utils import generate_output_path

logger = logging.getLogger(__name__)

# 支持的图片格式（真实的外部图片）
SUPPORTED_IMAGE_FORMATS = {"jpeg", "jpg", "png", "gif", "tiff", "tif", "bmp"}

# 忽略的格式（Office绘图形状）
IGNORED_FORMATS = {"emf", "wmf"}


def get_paragraph_images(para, images_info):
    """
    获取段落关联的图片列表（通过 paragraph 对象精确匹配）

    这是一个共享函数，可被所有文档转换器使用。

    参数:
        para: Word段落对象
        images_info: 图片信息列表

    返回:
        list: 与该段落关联的图片列表
    """
    para_images = []
    for img in images_info:
        img_para = img.get("paragraph")
        # 使用底层 XML 元素比较确保匹配准确性
        if img_para is not None and para._element is not None and img_para._element == para._element:
            para_images.append(img)
    return para_images


def extract_images_from_docx(
    doc, output_folder: str, original_file_path: str, progress_callback: Callable[[str], None] | None = None
) -> list[dict]:
    """
    从DOCX文档中递归提取所有图片（包括嵌套表格中的图片）

    参数:
        doc: Document对象（python-docx）
        output_folder: 图片保存的文件夹路径
        original_file_path: 原始DOCX文件路径（用于生成规范的图片文件名）
        progress_callback: 进度回调函数（可选）

    返回:
        List[Dict]: 图片信息列表，每个元素包含：
            - paragraph: 图片所在的Paragraph对象（关键用于匹配）
            - filename: 图片文件名
            - image_path: 图片完整路径
    """
    logger.info(f"开始递归提取DOCX中的图片到: {output_folder}")

    # 确保输出文件夹存在
    Path(output_folder).mkdir(parents=True, exist_ok=True)

    images_info = []

    # 使用可变对象存储计数器，以便在递归函数中修改
    counter = {"val": 1}

    para_index_by_element_id = {
        id(p._element): idx
        for idx, p in enumerate(getattr(doc, "paragraphs", []))
        if getattr(p, "_element", None) is not None
    }

    # 内部递归处理函数
    def _process_container(container):
        # 1. 处理容器中的段落
        if hasattr(container, "paragraphs"):
            for para in container.paragraphs:
                _process_paragraph_images(
                    para,
                    doc,
                    output_folder,
                    original_file_path,
                    images_info,
                    counter,
                    para_index_by_element_id,
                    progress_callback,
                )

        # 2. 处理容器中的表格
        if hasattr(container, "tables"):
            for table in container.tables:
                for row in table.rows:
                    for cell in row.cells:
                        # 递归处理单元格（单元格也是一个容器，包含段落和可能的嵌套表格）
                        _process_container(cell)

    try:
        # 从文档根节点开始处理
        _process_container(doc)
    except Exception as e:
        logger.error(f"递归提取图片时出错: {e}", exc_info=True)

    logger.info(f"图片提取完成，共提取 {len(images_info)} 张图片")
    return images_info


def _process_paragraph_images(
    para, doc, output_folder, original_file_path, images_info, counter, para_index_by_element_id, progress_callback
):
    """
    处理单个段落中的图片

    支持两种图片格式：
    - <w:drawing>: Office 2007+ 新格式，通过 DrawingML 嵌入图片
    - <w:pict>: 兼容模式（.doc 转换或旧版保存），通过 VML 嵌入图片
    """
    # XML 命名空间定义
    NS_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    NS_V = "{urn:schemas-microsoft-com:vml}"

    # 收集所有图片关系ID
    image_rIds = []

    try:
        # 1. 处理 <w:drawing> 元素（Office 2007+ 新格式）
        drawings = para._element.findall(f".//{NS_W}drawing")
        for drawing in drawings:
            blips = drawing.findall(f".//{NS_A}blip")
            for blip in blips:
                rId = blip.get(f"{NS_R}embed")
                if rId:
                    image_rIds.append(rId)

        # 2. 处理 <w:pict> 元素（兼容模式/VML 格式）
        picts = para._element.findall(f".//{NS_W}pict")
        for pict in picts:
            # VML 图片数据在 <v:imagedata> 元素中
            imagedatas = pict.findall(f".//{NS_V}imagedata")
            for imagedata in imagedatas:
                rId = imagedata.get(f"{NS_R}id")
                if rId:
                    image_rIds.append(rId)

        # 3. 处理收集到的所有图片
        for rId in image_rIds:
            try:
                # 注意：如果是在嵌套部分，doc.part 可能访问不到图片，需要从 para.part 访问
                # python-docx 的 ElementProxy (如 Paragraph) 通常有 part 属性指向它所属的 Part (DocumentPart, HeaderPart 等)
                part = para.part
                image_part = part.related_parts[rId]

                content_type = image_part.content_type
                ext = content_type.split("/")[-1] if "/" in content_type else "png"
                ext = ext.lower()

                if ext in SUPPORTED_IMAGE_FORMATS:
                    image_path = generate_output_path(
                        input_path=original_file_path,
                        output_dir=output_folder,
                        section=f"image{counter['val']}",
                        add_timestamp=True,
                        description="fromDocx",
                        file_type=ext,
                    )
                    filename = Path(image_path).name

                    with Path(image_path).open("wb") as f:
                        f.write(image_part.blob)

                    logger.debug(f"提取图片: {filename}")

                    # 记录图片信息，关键是存储 paragraph 对象用于后续匹配
                    para_index = (
                        para_index_by_element_id.get(id(para._element), -1)
                        if getattr(para, "_element", None) is not None
                        else -1
                    )
                    images_info.append(
                        {
                            "paragraph": para,  # 关键：存储对象引用
                            "filename": filename,
                            "image_path": image_path,
                            "para_index": para_index,
                        }
                    )

                    if progress_callback:
                        progress_callback(t("conversion.progress.extracting_images", count=counter["val"]))

                    counter["val"] += 1

                elif ext in IGNORED_FORMATS:
                    pass
                else:
                    logger.warning(f"未知图片格式: {ext}")

            except KeyError:
                logger.warning(f"无法找到关系ID为 {rId} 的图片部分")
            except Exception as e:
                logger.error(f"提取图片时出错: {e}")

    except Exception as e:
        logger.error(f"处理段落图片时出错: {e}")


def process_image_with_ocr(
    img: dict,
    keep_images: bool,
    enable_ocr: bool,
    output_folder: str,
    progress_callback=None,
    current_index: int = 1,
    total_images: int = 1,
    cancel_event=None,
    extraction_mode: str = "file",
    ocr_placement_mode: str = "image_md",
) -> str:
    """
    根据选项处理图片，生成对应的Markdown链接

    这是一个共享函数，可被所有文种转换器使用。

    参数:
        img: 图片信息字典，包含 'filename', 'image_path', 'para_index' 等
        keep_images: 是否保留图片
        enable_ocr: 是否启用OCR识别
        output_folder: 输出文件夹路径
        progress_callback: 进度回调函数（可选）
        extraction_mode: 提取模式
        ocr_placement_mode: OCR位置

    返回:
        str: Markdown链接（图片链接或图片md文件链接）
    """
    from docwen.converter.shared.image_md import process_image_with_ocr as shared_process_image_with_ocr
    from docwen.config.config_manager import config_manager
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
    extraction_mode: str = "file",
) -> str:
    """
    创建图片的markdown文件

    这是一个共享函数，可被所有文种转换器使用。

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
    from docwen.converter.shared.image_md import create_image_md_file as shared_create_image_md_file
    from docwen.config.config_manager import config_manager
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
        extraction_mode=extraction_mode,
        image_link_style=image_link_style,
        ocr_progress_message_factory=_ocr_progress,
    )
