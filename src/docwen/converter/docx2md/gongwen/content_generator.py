"""
公文内容生成模块

负责生成公文模式的Markdown内容。

主要功能：
- generate_main_content(): 生成主文档Markdown内容（含YAML头部）
- generate_attachment_content(): 生成附件Markdown内容

设计说明：
    公文模式**不保留**粗体、斜体等字符格式，原因：
    1. 公文有严格格式规范（GB/T 9704-2012），不允许装饰性格式
    2. 侧重于结构化元素提取，而非保留视觉格式
    3. 确保输出规范性

依赖：
- shared/image_processor: 图片处理
- shared/formula_processor: 公式处理
- utils/heading_utils: 标题处理
- utils/date_utils: 日期处理
"""

import datetime
import logging
import re
import threading
from collections.abc import Callable
from pathlib import Path

from docwen.utils.heading_utils import add_markdown_heading, detect_heading_level, split_content_by_delimiters
from docwen.utils.text_utils import format_yaml_value

from ..shared.break_processor import detect_page_or_section_break
from ..shared.formula_processor import has_formulas_in_paragraph, is_formula_supported, process_paragraph_with_formulas
from ..shared.image_processor import get_paragraph_images, process_image_with_ocr

logger = logging.getLogger(__name__)


def generate_main_content(
    yaml_info: dict,
    doc,
    skip_para_indices: list,
    images_info: list | None = None,
    keep_images: bool = True,
    enable_ocr: bool = False,
    output_folder: str | None = None,
    progress_callback: Callable[[str], None] | None = None,
    cancel_event: threading.Event | None = None,
    remove_numbering: bool = False,
    add_numbering: bool = False,
    scheme_name: str = "gongwen_standard",
    extraction_mode: str = "file",
    ocr_placement_mode: str = "image_md",
) -> str:
    """
    生成主要Markdown内容

    处理流程：
    1. 生成YAML头部（包含14个公文元数据字段）
    2. 遍历段落，跳过已提取到YAML的段落
    3. 处理公式、分页符、标题、普通段落
    4. 插入图片链接

    参数:
        yaml_info: dict - YAML元数据字典，包含以下字段：
            - aliases: list[str] - 标题别名列表
            - 标题: str - 公文标题
            - 份号: str - 份号
            - 密级和保密期限: str - 密级信息
            - 紧急程度: str - 紧急程度
            - 发文字号: str - 发文字号
            - 发文机关标志: str - 发文机关标志
            - 签发人: list[str] - 签发人列表
            - 发文机关署名: str - 发文机关署名
            - 成文日期: str - 成文日期
            - 印发日期: str - 印发日期
            - 主送机关: str - 主送机关
            - 附注: str - 附注
            - 印发机关: str - 印发机关
            - 抄送机关: list[str] - 抄送机关列表
            - 附件说明: list[str] - 附件说明列表
            - 公开方式: str - 公开方式
        doc: Document - Word文档对象
        skip_para_indices: list[int] - 需要跳过的段落索引列表
        images_info: list - 图片信息列表（可选）
        keep_images: bool - 是否保留图片（默认True）
        enable_ocr: bool - 是否启用OCR识别（默认False）
        output_folder: str - 输出文件夹路径（可选）
        progress_callback: Callable - 进度回调函数（可选）
        cancel_event: threading.Event - 取消事件（可选）
        remove_numbering: bool - 是否清除原有序号（默认False）
        add_numbering: bool - 是否添加新序号（默认False）
        scheme_name: str - 序号方案名称（默认"gongwen_standard"）

    返回:
        str: 完整的Markdown内容，包含YAML头部和正文
    """
    if images_info is None:
        images_info = []

    logger.info(f"生成Markdown内容 - 图片数量: {len(images_info)}, 保留图片: {keep_images}, OCR: {enable_ocr}")
    logger.info(f"序号配置 - 清除: {remove_numbering}, 添加: {add_numbering}, 方案: {scheme_name}")

    # 初始化序号格式化器
    formatter = _init_formatter(add_numbering, scheme_name)
    if formatter is None and add_numbering:
        add_numbering = False

    # 计算需要OCR的图片总数
    total_images = sum(1 for img in images_info if img["para_index"] not in skip_para_indices)
    current_image_index = 0

    lines = []

    # 写入YAML头部
    lines.extend(_build_yaml_header(yaml_info))
    lines.append("")

    # 创建内容列表
    content_lines = []

    # 导入配置管理器
    from docwen.config.config_manager import config_manager

    # 处理所有段落
    for idx, para in enumerate(doc.paragraphs):
        # 跳过已提取到YAML的段落
        if idx in skip_para_indices:
            logger.debug(f"跳过已提取到YAML的段落: {idx + 1}")
            continue

        # 处理公式
        if is_formula_supported() and has_formulas_in_paragraph(para):
            formula_md = process_paragraph_with_formulas(para)
            if formula_md:
                content_lines.append(formula_md)
                logger.debug(f"段落 {idx + 1} 包含公式，已转换")

                # 处理图片
                para_images = get_paragraph_images(para, images_info)
                for img in para_images:
                    current_image_index += 1
                    image_link = process_image_with_ocr(
                        img,
                        keep_images,
                        enable_ocr,
                        output_folder or "",
                        progress_callback,
                        current_image_index,
                        total_images,
                        cancel_event,
                        extraction_mode=extraction_mode,
                        ocr_placement_mode=ocr_placement_mode,
                    )
                    content_lines.append(image_link)
                continue

        # 处理分页符/分节符
        if config_manager.is_horizontal_rule_enabled():
            break_type, _break_value = detect_page_or_section_break(para)
            if break_type:
                md_separator = config_manager.get_md_separator_for_break_type(break_type)
                if md_separator:
                    content_lines.append(md_separator)
                    logger.debug(f"段落 {idx + 1} 包含 {break_type}")
                continue

        # 获取段落文本
        para_text = para.text.strip()

        # 处理空段落
        if not para_text:
            para_images = get_paragraph_images(para, images_info)
            if para_images:
                for img in para_images:
                    current_image_index += 1
                    image_link = process_image_with_ocr(
                        img,
                        keep_images,
                        enable_ocr,
                        output_folder or "",
                        progress_callback,
                        current_image_index,
                        total_images,
                        cancel_event,
                        extraction_mode=extraction_mode,
                        ocr_placement_mode=ocr_placement_mode,
                    )
                    content_lines.append(image_link)
                    logger.debug(f"在空段落 {idx + 1} 插入图片")
            continue

        # 处理段落内容
        paragraph_content = _process_paragraph_content(para, para_text, idx, remove_numbering, add_numbering, formatter)
        content_lines.append(paragraph_content)

        # 处理段落后的图片
        para_images = get_paragraph_images(para, images_info)
        for img in para_images:
            current_image_index += 1
            image_link = process_image_with_ocr(
                img,
                keep_images,
                enable_ocr,
                output_folder or "",
                progress_callback,
                current_image_index,
                total_images,
                cancel_event,
                extraction_mode=extraction_mode,
                ocr_placement_mode=ocr_placement_mode,
            )
            content_lines.append(image_link)
            logger.debug(f"在段落 {idx + 1} 后插入图片")

    # 添加正文内容
    if content_lines:
        lines.append("\n\n".join(content_lines))

    return "\n".join(lines)


def generate_attachment_content(
    doc,
    content_indices: list,
    original_docx_path: str,
    images_info: list | None = None,
    keep_images: bool = True,
    enable_ocr: bool = False,
    output_folder: str | None = None,
    progress_callback: Callable[[str], None] | None = None,
    cancel_event: threading.Event | None = None,
    remove_numbering: bool = False,
    add_numbering: bool = False,
    scheme_name: str = "gongwen_standard",
    extraction_mode: str = "file",
    ocr_placement_mode: str = "image_md",
) -> str:
    """
    生成附件Markdown内容

    处理流程：
    1. 生成简化的YAML头部（来源文件、提取时间、类型）
    2. 遍历附件段落
    3. 处理公式、标题、普通段落
    4. 插入图片链接

    参数:
        doc: Document - Word文档对象
        content_indices: list[int] - 附件内容段落的索引列表
        original_docx_path: str - 原始DOCX文件路径
        images_info: list - 图片信息列表（可选）
        keep_images: bool - 是否保留图片（默认True）
        enable_ocr: bool - 是否启用OCR识别（默认False）
        output_folder: str - 输出文件夹路径（可选）
        progress_callback: Callable - 进度回调函数（可选）
        cancel_event: threading.Event - 取消事件（可选）
        remove_numbering: bool - 是否清除原有序号（默认False）
        add_numbering: bool - 是否添加新序号（默认False）
        scheme_name: str - 序号方案名称（默认"gongwen_standard"）

    返回:
        str: 附件Markdown内容，包含YAML头部和正文
    """
    if images_info is None:
        images_info = []

    # 初始化序号格式化器
    formatter = _init_formatter(add_numbering, scheme_name)
    if formatter is None and add_numbering:
        add_numbering = False

    # 计算附件中的图片总数
    total_images = sum(1 for img in images_info if img["para_index"] in content_indices)
    current_image_index = 0

    lines = []

    # 添加YAML头部
    lines.append("---")
    lines.append(f"来源文件: {Path(original_docx_path).name}")
    lines.append(f"提取时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("类型: 附件内容")
    lines.append("---")
    lines.append("")

    # 提取附件内容
    attachment_lines = []
    for idx in content_indices:
        if idx >= len(doc.paragraphs):
            continue

        para = doc.paragraphs[idx]

        # 处理公式
        if is_formula_supported() and has_formulas_in_paragraph(para):
            formula_md = process_paragraph_with_formulas(para)
            if formula_md:
                attachment_lines.append(formula_md)
                logger.debug(f"附件段落 {idx + 1} 包含公式")

                para_images = get_paragraph_images(para, images_info)
                for img in para_images:
                    current_image_index += 1
                    image_link = process_image_with_ocr(
                        img,
                        keep_images,
                        enable_ocr,
                        output_folder or "",
                        progress_callback,
                        current_image_index,
                        total_images,
                        cancel_event,
                        extraction_mode=extraction_mode,
                        ocr_placement_mode=ocr_placement_mode,
                    )
                    attachment_lines.append(image_link)
                continue

        para_text = para.text.strip()

        # 处理空段落
        if not para_text:
            para_images = get_paragraph_images(para, images_info)
            if para_images:
                for img in para_images:
                    current_image_index += 1
                    image_link = process_image_with_ocr(
                        img,
                        keep_images,
                        enable_ocr,
                        output_folder or "",
                        progress_callback,
                        current_image_index,
                        total_images,
                        cancel_event,
                        extraction_mode=extraction_mode,
                        ocr_placement_mode=ocr_placement_mode,
                    )
                    attachment_lines.append(image_link)
                    logger.debug(f"在附件空段落 {idx + 1} 插入图片")
            continue

        # 处理段落内容
        paragraph_content = _process_paragraph_content(para, para_text, idx, remove_numbering, add_numbering, formatter)
        attachment_lines.append(paragraph_content)

        # 处理段落后的图片
        para_images = get_paragraph_images(para, images_info)
        for img in para_images:
            current_image_index += 1
            image_link = process_image_with_ocr(
                img,
                keep_images,
                enable_ocr,
                output_folder or "",
                progress_callback,
                current_image_index,
                total_images,
                cancel_event,
                extraction_mode=extraction_mode,
                ocr_placement_mode=ocr_placement_mode,
            )
            attachment_lines.append(image_link)
            logger.debug(f"在附件段落 {idx + 1} 后插入图片")

    # 添加正文内容
    if attachment_lines:
        lines.append("\n\n".join(attachment_lines))

    return "\n".join(lines)


# ========== 私有辅助函数 ==========


def _init_formatter(add_numbering: bool, scheme_name: str):
    """
    初始化序号格式化器

    参数:
        add_numbering: bool - 是否需要添加序号
        scheme_name: str - 序号方案名称

    返回:
        HeadingFormatter or None: 格式化器实例，如果不需要或创建失败则返回None
    """
    if not add_numbering:
        return None

    from docwen.config.config_manager import config_manager
    from docwen.utils.heading_numbering import get_formatter_from_config

    formatter = get_formatter_from_config(config_manager, scheme_name)
    if formatter:
        formatter.reset_counters()
        logger.info(f"已创建序号格式化器，方案：{scheme_name}")
        return formatter
    else:
        logger.warning(f"无法创建序号格式化器（方案：{scheme_name}）")
        return None


def _build_yaml_header(yaml_info: dict) -> list:
    """
    构建YAML头部行列表

    使用 format_yaml_value() 安全处理特殊字符（如 [] {} : # ' " 等）。

    参数:
        yaml_info: dict - YAML元数据字典

    返回:
        list[str]: YAML头部行列表
    """
    lines = ["---"]

    for key, value in yaml_info.items():
        # 特殊处理列表类型
        if isinstance(value, list):
            if value:
                lines.append(f"{key}:")
                for item in value:
                    # 使用 format_yaml_value 安全处理（处理特殊字符、数字、引号等）
                    safe_item = format_yaml_value(item)
                    lines.append(f"  - {safe_item}")
            else:
                lines.append(f"{key}: []")
        else:
            # 使用 format_yaml_value 安全处理
            safe_value = format_yaml_value(value)
            lines.append(f"{key}: {safe_value}")

    lines.append("---")
    return lines


def _process_paragraph_content(
    para, para_text: str, idx: int, remove_numbering: bool, add_numbering: bool, formatter
) -> str:
    """
    处理段落内容，生成Markdown文本

    参数:
        para: Paragraph - 段落对象
        para_text: str - 段落文本
        idx: int - 段落索引
        remove_numbering: bool - 是否清除序号
        add_numbering: bool - 是否添加序号
        formatter: HeadingFormatter - 序号格式化器

    返回:
        str: 处理后的Markdown文本
    """
    # 检测标题级别（双重检测策略）
    # 1. 先根据序号格式检测
    _, heading_level = detect_heading_level(para_text)

    # 2. 如果序号格式没有匹配，尝试Word内置样式
    if heading_level == 0:
        style_name = para.style.name if para.style else None
        if style_name and style_name.startswith("Heading"):
            match = re.match(r"Heading (\d)", style_name)
            if match:
                heading_level = int(match.group(1))
                logger.debug(f"段落 {idx + 1} 通过Word样式识别为 {heading_level} 级标题")

    # 应用序号配置
    if heading_level > 0:
        # 是标准小标题
        base_text = para_text

        # 清除序号
        if remove_numbering:
            from docwen.utils.heading_utils import remove_heading_numbering

            base_text = remove_heading_numbering(para_text)
            logger.debug(f"清除序号: '{para_text}' -> '{base_text}'")

        # 添加新序号
        if add_numbering and formatter:
            formatter.increment_level(heading_level)
            numbering = formatter.format_heading(heading_level)
            final_title_text = numbering + base_text
            logger.debug(f"添加序号到标题: '{base_text}' -> '{final_title_text}'")
        else:
            final_title_text = base_text

        # 分割内容（标题可能与正文在同一段）
        content1, content2 = split_content_by_delimiters(final_title_text)

        # 添加Markdown标题符号
        if content1:
            content1 = add_markdown_heading(content1, heading_level)

        # 组合段落内容
        if content1 and content2:
            return f"{content1}\n{content2}"
        elif content1:
            return content1
        else:
            return final_title_text
    else:
        # 非小标题段落直接使用原始文本
        return para_text
