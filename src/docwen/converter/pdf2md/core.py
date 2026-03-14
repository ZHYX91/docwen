"""
PDF转Markdown核心模块 - 基于pymupdf4llm

使用pymupdf4llm提取PDF内容并生成Markdown文件
    支持4种提取组合：纯文本、文本+图片、文本+图片+OCR、文本+OCR（不保留图片）

优化方案：
- 在临时目录中创建子文件夹
- pymupdf4llm直接输出到子文件夹
- MD文件也保存在子文件夹
- 整体移动子文件夹到目标位置
"""

import logging
import tempfile
import threading
import time
from collections.abc import Callable
from pathlib import Path
from typing import Any

from docwen.translation import t
from docwen.utils.workspace_manager import finalize_output
from docwen.utils.yaml_utils import generate_basic_yaml_frontmatter

logger = logging.getLogger(__name__)

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif"}


def _ensure_markdown_text(value: Any) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, list):
        return "\n".join(str(item) for item in value)
    return str(value)


def _monitor_image_extraction(
    output_folder: str,
    stop_event: threading.Event,
    progress_callback: Callable[[str], Any] | None,
) -> None:
    """
    监控图片提取进度的后台线程

    用于监控pymupdf4llm在提取图片时的实时进度，
    定期检查输出文件夹中的图片文件数量并通过回调函数反馈

    参数:
        output_folder: 输出文件夹路径
        stop_event: 停止监控的事件标志
        progress_callback: 进度回调函数
    """
    if not progress_callback:
        return

    last_count = 0

    try:
        while not stop_event.is_set():
            try:
                # 统计当前图片数量
                folder_path = Path(output_folder)
                if folder_path.exists():
                    current_count = sum(
                        1 for f in folder_path.iterdir() if f.is_file() and f.suffix.lower() in IMAGE_EXTENSIONS
                    )

                    # 如果有新图片
                    if current_count > last_count:
                        progress_callback(t("conversion.progress.extracting_images", count=current_count))
                        last_count = current_count

                # 每2秒检查一次（减少性能开销，小文件可能直接显示总数）
                time.sleep(2)

            except Exception as e:
                logger.error(f"监控图片提取时出错: {e}")
                break
    except Exception as e:
        logger.error(f"监控线程异常: {e}")


def extract_pdf_with_pymupdf4llm(
    pdf_path: str,
    extract_images: bool,
    extract_ocr: bool,
    output_dir: str,
    basename_for_output: str | None = None,
    cancel_event: threading.Event | None = None,
    progress_callback: Callable[[str], Any] | None = None,
    ocr_placement_mode: str = "image_md",
    extraction_mode: str = "file",
) -> dict[str, Any]:
    """
    使用pymupdf4llm提取PDF内容（支持3种组合）

    优化方案：临时子文件夹 → 整体移动

    四种组合模式：
    1. ❌图片 ❌OCR：纯文本MD
    2. ✅图片 ❌OCR：MD + 图片
    3. ✅图片 ✅OCR：MD + 图片 + OCR
    4. ❌图片 ✅OCR：MD（图片位置以 OCR 文本替换，不保留图片文件）

    参数:
        pdf_path: PDF文件路径
        extract_images: 是否提取图片
        extract_ocr: 是否进行OCR识别
        output_dir: 输出目录（父目录）
        basename_for_output: 输出文件夹和文件的基础名（不含.md）
        cancel_event: 取消事件(可选)
        progress_callback: 进度回调函数(可选)

    返回:
        Dict: 提取结果，包含：
            - md_path: str - Markdown文件完整路径
            - folder_path: str - 输出文件夹路径
            - image_count: int - 图片数量
            - ocr_count: int - OCR识别数量
    """
    try:
        import pymupdf4llm
    except ImportError as e:
        raise ImportError("pymupdf4llm库未安装。\n请安装: pip install pymupdf4llm") from e

    if not Path(pdf_path).exists():
        raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")

    logger.info(f"开始提取PDF内容: {Path(pdf_path).name}")
    logger.info(f"提取选项: 图片={extract_images}, OCR={extract_ocr}")

    # 检查取消
    if cancel_event and cancel_event.is_set():
        raise InterruptedError("操作已被取消")

    # 确定输出文件名（不含.md后缀）
    if not basename_for_output:
        basename_for_output = Path(pdf_path).stem

    # 使用临时目录
    with tempfile.TemporaryDirectory() as temp_dir:
        logger.debug(f"临时目录: {temp_dir}")

        # 创建临时输出文件夹
        temp_output_folder = Path(temp_dir) / basename_for_output
        temp_output_folder.mkdir(parents=True, exist_ok=True)
        logger.debug(f"临时输出文件夹: {temp_output_folder}")

        # 情况1：纯文本（❌图片 ❌OCR）
        if not extract_images and not extract_ocr:
            logger.info("模式: 纯文本提取")

            md_text = _ensure_markdown_text(pymupdf4llm.to_markdown(pdf_path, write_images=False))

            # 生成YAML头部（使用原始文件名，不含扩展名）
            original_file_stem = Path(pdf_path).stem
            yaml_frontmatter = generate_basic_yaml_frontmatter(original_file_stem)

            # 保存MD文件（YAML头部 + 正文内容）
            md_filename = f"{basename_for_output}.md"
            md_path_temp = temp_output_folder / md_filename
            with md_path_temp.open("w", encoding="utf-8") as f:
                f.write(yaml_frontmatter + md_text)

            logger.debug(f"已生成临时MD: {md_path_temp}")

            # 移动整个文件夹到目标位置
            final_folder, _ = finalize_output(
                str(temp_output_folder),
                str(Path(output_dir) / basename_for_output),
                original_input_file=pdf_path,
            )
            if not final_folder:
                raise RuntimeError("保存输出文件夹失败")
            logger.info(f"已保存到: {final_folder}")

            final_md_path = str(Path(final_folder) / md_filename)

            return {"md_path": final_md_path, "folder_path": final_folder, "image_count": 0, "ocr_count": 0}

        # 情况2：文本+图片（✅图片 ❌OCR）
        elif extract_images and not extract_ocr:
            logger.info("模式: 文本+图片")

            # 启动图片提取监控线程
            stop_monitor = threading.Event()
            monitor_thread = None
            if progress_callback:
                monitor_thread = threading.Thread(
                    target=_monitor_image_extraction,
                    args=(str(temp_output_folder), stop_monitor, progress_callback),
                    daemon=True,
                )
                monitor_thread.start()
                logger.debug("已启动图片提取监控线程")

            try:
                # pymupdf4llm直接输出到临时文件夹
                md_text = _ensure_markdown_text(
                    pymupdf4llm.to_markdown(
                        pdf_path, write_images=True, image_path=str(temp_output_folder), image_format="png"
                    )
                )
            finally:
                # 停止监控线程
                if monitor_thread:
                    stop_monitor.set()
                    monitor_thread.join(timeout=2)
                    logger.debug("图片提取监控线程已停止")

            # 检查取消
            if cancel_event and cancel_event.is_set():
                raise InterruptedError("操作已被取消")

            # 统计图片数量
            image_count = _count_images_in_folder(str(temp_output_folder))
            logger.info(f"提取了{image_count}张图片")

            # 提示图片提取完成的总数
            if progress_callback:
                progress_callback(t("conversion.progress.images_extracted", count=image_count))

            md_text, _ = _apply_image_rules(
                md_text,
                str(temp_output_folder),
                keep_images=True,
                enable_ocr=False,
                extraction_mode=extraction_mode,
                cancel_event=cancel_event,
            )
            if extraction_mode == "base64":
                _delete_images_in_folder(str(temp_output_folder))

            # 生成YAML头部（使用原始文件名，不含扩展名）
            original_file_stem = Path(pdf_path).stem
            yaml_frontmatter = generate_basic_yaml_frontmatter(original_file_stem)

            # 保存MD文件（YAML头部 + 正文内容）
            md_filename = f"{basename_for_output}.md"
            md_path_temp = temp_output_folder / md_filename
            with md_path_temp.open("w", encoding="utf-8") as f:
                f.write(yaml_frontmatter + md_text)

            logger.debug(f"已生成临时MD: {md_path_temp}")

            # 移动整个文件夹到目标位置
            final_folder, _ = finalize_output(
                str(temp_output_folder),
                str(Path(output_dir) / basename_for_output),
                original_input_file=pdf_path,
            )
            if not final_folder:
                raise RuntimeError("保存输出文件夹失败")
            logger.info(f"已保存到: {final_folder}")

            final_md_path = str(Path(final_folder) / md_filename)

            return {"md_path": final_md_path, "folder_path": final_folder, "image_count": image_count, "ocr_count": 0}
        # 情况3：文本+图片+OCR（✅图片 ✅OCR）
        elif extract_images and extract_ocr:
            logger.info("模式: 文本+图片+OCR")

            # 启动图片提取监控线程
            stop_monitor = threading.Event()
            monitor_thread = None
            if progress_callback:
                monitor_thread = threading.Thread(
                    target=_monitor_image_extraction,
                    args=(str(temp_output_folder), stop_monitor, progress_callback),
                    daemon=True,
                )
                monitor_thread.start()
                logger.debug("已启动图片提取监控线程")

            try:
                # pymupdf4llm直接输出到临时文件夹
                md_text = _ensure_markdown_text(
                    pymupdf4llm.to_markdown(
                        pdf_path, write_images=True, image_path=str(temp_output_folder), image_format="png"
                    )
                )
            finally:
                # 停止监控线程
                if monitor_thread:
                    stop_monitor.set()
                    monitor_thread.join(timeout=2)
                    logger.debug("图片提取监控线程已停止")

            # 检查取消
            if cancel_event and cancel_event.is_set():
                raise InterruptedError("操作已被取消")

            # 统计图片数量
            image_count = _count_images_in_folder(str(temp_output_folder))
            logger.info(f"提取了{image_count}张图片")

            # 提示图片提取完成的总数
            if progress_callback:
                progress_callback(t("conversion.progress.images_extracted_preparing_ocr", count=image_count))

            md_text, ocr_count = _apply_image_rules(
                md_text,
                str(temp_output_folder),
                keep_images=True,
                enable_ocr=True,
                extraction_mode=extraction_mode,
                ocr_placement_mode=ocr_placement_mode,
                cancel_event=cancel_event,
                progress_callback=progress_callback,
            )
            logger.info(f"OCR识别完成: {ocr_count}个结果")

            if extraction_mode == "base64":
                _delete_images_in_folder(str(temp_output_folder))

            # 生成YAML头部（使用原始文件名，不含扩展名）
            original_file_stem = Path(pdf_path).stem
            yaml_frontmatter = generate_basic_yaml_frontmatter(original_file_stem)

            # 保存MD文件（YAML头部 + 正文内容）
            md_filename = f"{basename_for_output}.md"
            md_path_temp = temp_output_folder / md_filename
            with md_path_temp.open("w", encoding="utf-8") as f:
                f.write(yaml_frontmatter + md_text)

            logger.debug(f"已生成临时MD: {md_path_temp}")

            # 移动整个文件夹到目标位置
            final_folder, _ = finalize_output(
                str(temp_output_folder),
                str(Path(output_dir) / basename_for_output),
                original_input_file=pdf_path,
            )
            if not final_folder:
                raise RuntimeError("保存输出文件夹失败")
            logger.info(f"已保存到: {final_folder}")

            final_md_path = str(Path(final_folder) / md_filename)

            return {
                "md_path": final_md_path,
                "folder_path": final_folder,
                "image_count": image_count,
                "ocr_count": ocr_count,
            }

        # 情况4：文本+OCR（❌图片 ✅OCR）
        elif not extract_images and extract_ocr:
            logger.info("模式: 文本+OCR（不保留图片）")

            # 为 OCR 需要先导出图片到临时文件夹
            md_text = _ensure_markdown_text(
                pymupdf4llm.to_markdown(
                    pdf_path, write_images=True, image_path=str(temp_output_folder), image_format="png"
                )
            )

            # 检查取消
            if cancel_event and cancel_event.is_set():
                raise InterruptedError("操作已被取消")

            # 统计图片数量
            image_count = _count_images_in_folder(str(temp_output_folder))
            logger.info(f"提取了{image_count}张图片用于OCR")

            if progress_callback:
                progress_callback(t("conversion.progress.images_extracted_preparing_ocr", count=image_count))

            md_text, ocr_count = _apply_image_rules(
                md_text,
                str(temp_output_folder),
                keep_images=False,
                enable_ocr=True,
                ocr_placement_mode=ocr_placement_mode,
                cancel_event=cancel_event,
                progress_callback=progress_callback,
            )

            _delete_images_in_folder(str(temp_output_folder))

            original_file_stem = Path(pdf_path).stem
            yaml_frontmatter = generate_basic_yaml_frontmatter(original_file_stem)

            md_filename = f"{basename_for_output}.md"
            md_path_temp = temp_output_folder / md_filename
            with md_path_temp.open("w", encoding="utf-8") as f:
                f.write(yaml_frontmatter + md_text)

            final_folder, _ = finalize_output(
                str(temp_output_folder),
                str(Path(output_dir) / basename_for_output),
                original_input_file=pdf_path,
            )
            if not final_folder:
                raise RuntimeError("保存输出文件夹失败")
            logger.info(f"已保存到: {final_folder}")

            final_md_path = str(Path(final_folder) / md_filename)

            return {
                "md_path": final_md_path,
                "folder_path": final_folder,
                "image_count": 0,
                "ocr_count": ocr_count,
            }

        raise ValueError("无效的提取选项组合")


def _count_images_in_folder(folder_path: str) -> int:
    """
    统计文件夹中的图片数量

    参数:
        folder_path: 文件夹路径

    返回:
        int: 图片数量
    """
    return sum(1 for f in Path(folder_path).iterdir() if f.is_file() and f.suffix.lower() in IMAGE_EXTENSIONS)


def _delete_images_in_folder(folder_path: str):
    """
    删除文件夹中的所有图片文件

    参数:
        folder_path: 文件夹路径
    """
    for f in Path(folder_path).iterdir():
        if f.is_file() and f.suffix.lower() in IMAGE_EXTENSIONS:
            f.unlink(missing_ok=True)
            logger.debug(f"删除图片: {f.name}")


def _convert_to_simple_paths(md_text: str) -> str:
    """
    将Markdown中的图片路径转换为纯文件名，并应用配置的链接格式

    例如: C:/temp/xyz/image_1.png -> image_1.png
    同时根据配置转换为 Markdown 或 Wiki 格式

    参数:
        md_text: Markdown文本

    返回:
        str: 更新后的Markdown文本
    """
    import re

    from docwen.config.config_manager import config_manager
    from docwen.utils.markdown_utils import format_image_link

    # 获取链接格式配置
    link_settings = config_manager.get_markdown_link_style_settings()
    image_link_style = link_settings.get("image_link_style", "wiki_embed")

    # 查找所有图片链接 ![alt](path) 或 ![alt](<path> "title")
    pattern = r'!\[([^\]]*)\]\(\s*(<[^>]+>|[^)\s]+)(?:\s+["\'][^"\']*["\'])?\s*\)'

    def replace_path(match):
        img_path = match.group(2).strip()
        if img_path.startswith("<") and img_path.endswith(">"):
            img_path = img_path[1:-1].strip()

        # 提取纯文件名
        img_filename = Path(img_path).name

        # 使用配置的格式
        return format_image_link(img_filename, image_link_style)

    result = re.sub(pattern, replace_path, md_text)

    # 统一使用正斜杠
    result = result.replace("\\", "/")

    return result


def _apply_image_rules(
    md_text: str,
    images_folder: str,
    keep_images: bool,
    enable_ocr: bool,
    *,
    extraction_mode: str = "file",
    ocr_placement_mode: str = "image_md",
    cancel_event: threading.Event | None = None,
    progress_callback: Callable[[str], Any] | None = None,
) -> tuple[str, int]:
    import re

    from docwen.config.config_manager import config_manager
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

    pattern = r'(!\[([^\]]*)\]\(\s*(<[^>]+>|[^)\s]+)(?:\s+["\'][^"\']*["\'])?\s*\))'
    image_matches = list(re.finditer(pattern, md_text))
    total_images = len(image_matches)

    ocr_count = 0
    current_index = 0

    def replace_match(match: re.Match[str]) -> str:
        nonlocal ocr_count, current_index

        current_index += 1
        img_path = match.group(3).strip()
        if img_path.startswith("<") and img_path.endswith(">"):
            img_path = img_path[1:-1].strip()

        if cancel_event and cancel_event.is_set():
            raise InterruptedError("操作已被取消")

        img_path_obj = Path(img_path)
        full_img_path = img_path_obj if img_path_obj.is_absolute() else Path(images_folder) / img_path_obj.name
        if not full_img_path.exists():
            logger.warning(f"图片不存在: {full_img_path}")
            return ""

        img_info = {"filename": full_img_path.name, "image_path": str(full_img_path)}
        link_text = process_image_with_ocr(
            img=img_info,
            keep_images=keep_images,
            enable_ocr=enable_ocr,
            output_folder=images_folder,
            progress_callback=progress_callback if enable_ocr else None,
            current_index=current_index,
            total_images=total_images,
            cancel_event=cancel_event,
            extraction_mode=extraction_mode,
            ocr_placement_mode=ocr_placement_mode,
            image_link_style=image_link_style,
            md_file_link_style=md_file_link_style,
            ocr_blockquote_title=ocr_title,
            ocr_progress_message_factory=_ocr_progress,
        )
        if enable_ocr and ocr_placement_mode == "main_md" and ("\n\n>" in link_text or link_text.startswith("> ")):
            ocr_count += 1
        return link_text

    result = re.sub(pattern, replace_match, md_text)

    if enable_ocr and ocr_placement_mode == "image_md":
        ocr_count = total_images

    return result, ocr_count


def _create_image_md_file(
    image_path: str,
    image_filename: str,
    output_folder: str,
    include_image: bool,
    include_ocr: bool,
    cancel_event: threading.Event | None = None,
    progress_callback: Callable[[str], Any] | None = None,
    current_index: int = 1,
    total_images: int = 1,
    extraction_mode: str = "file",
) -> str:
    """
    创建图片的markdown文件（与docx转MD保持一致）

    参数:
        image_path: 图片文件的完整路径
        image_filename: 图片文件名（如 'image_1.png'）
        output_folder: 输出文件夹路径
        include_image: 是否在md文件中包含图片链接
        include_ocr: 是否在md文件中包含OCR识别文本
        cancel_event: 取消事件(可选)

    返回:
        str: 创建的md文件名（如 'image_1.md'）
    """
    from docwen.config.config_manager import config_manager
    from docwen.converter.shared.image_md import create_image_md_file as shared_create_image_md_file

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


def _replace_images_with_ocr(
    md_text: str,
    images_folder: str,
    cancel_event: threading.Event | None = None,
    progress_callback: Callable[[str], Any] | None = None,
) -> tuple[str, int]:
    """
    替换Markdown中的图片链接为图片md文件链接（只含OCR文本）

    用于情况3：只要OCR文本，不保存图片文件
    现在改为与docx转MD一致：生成图片.md文件并插入链接

    参数:
        md_text: Markdown文本
        images_folder: 图片文件夹路径
        cancel_event: 取消事件(可选)
        progress_callback: 进度回调函数(可选)

    返回:
        tuple: (更新后的Markdown文本, OCR识别数量)
    """
    logger.info("开始OCR识别并生成图片md文件...")
    result, ocr_count = _apply_image_rules(
        md_text,
        images_folder,
        keep_images=False,
        enable_ocr=True,
        ocr_placement_mode="image_md",
        cancel_event=cancel_event,
        progress_callback=progress_callback,
    )
    logger.info(f"图片md文件生成完成，共{ocr_count}个")
    return result, ocr_count


def _replace_images_with_ocr_blockquote(
    md_text: str,
    images_folder: str,
    cancel_event: threading.Event | None = None,
    progress_callback: Callable[[str], Any] | None = None,
) -> tuple[str, int]:
    return _apply_image_rules(
        md_text,
        images_folder,
        keep_images=False,
        enable_ocr=True,
        ocr_placement_mode="main_md",
        cancel_event=cancel_event,
        progress_callback=progress_callback,
    )


def _add_ocr_blockquote_after_images(
    md_text: str,
    images_folder: str,
    cancel_event: threading.Event | None = None,
    progress_callback: Callable[[str], Any] | None = None,
) -> tuple[str, int]:
    return _apply_image_rules(
        md_text,
        images_folder,
        keep_images=True,
        enable_ocr=True,
        ocr_placement_mode="main_md",
        cancel_event=cancel_event,
        progress_callback=progress_callback,
    )


def _add_ocr_after_images(
    md_text: str,
    images_folder: str,
    cancel_event: threading.Event | None = None,
    progress_callback: Callable[[str], Any] | None = None,
    extraction_mode: str = "file",
) -> tuple[str, int]:
    """
    将Markdown中的图片链接替换为图片md文件链接（图片md含图片链接+OCR文本）

    用于情况4：提取图片 + OCR
    主MD中使用图片.md链接代替图片链接；图片.md中包含图片链接与OCR结果

    参数:
        md_text: Markdown文本
        images_folder: 图片文件夹路径
        cancel_event: 取消事件(可选)
        progress_callback: 进度回调函数(可选)

    返回:
        tuple: (更新后的Markdown文本, OCR识别数量)
    """
    logger.info("开始OCR识别并生成图片md文件...")
    result, ocr_count = _apply_image_rules(
        md_text,
        images_folder,
        keep_images=True,
        enable_ocr=True,
        extraction_mode=extraction_mode,
        ocr_placement_mode="image_md",
        cancel_event=cancel_event,
        progress_callback=progress_callback,
    )
    logger.info(f"图片md文件生成完成，共{ocr_count}个")
    return result, ocr_count


# ==================== 公开API ====================

__all__ = [
    "extract_pdf_with_pymupdf4llm",
]
