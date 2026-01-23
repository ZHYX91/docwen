"""
PDF转Markdown核心模块 - 基于pymupdf4llm

使用pymupdf4llm提取PDF内容并生成Markdown文件
支持4种提取组合：纯文本、文本+图片、文本+OCR、文本+图片+OCR

优化方案：
- 在临时目录中创建子文件夹
- pymupdf4llm直接输出到子文件夹
- MD文件也保存在子文件夹
- 整体移动子文件夹到目标位置
"""

import os
import logging
import tempfile
import shutil
import time
from typing import Dict, Optional
import threading
from pathlib import Path

from docwen.utils.yaml_utils import generate_basic_yaml_frontmatter
from docwen.i18n import t

logger = logging.getLogger(__name__)


def _monitor_image_extraction(
    output_folder: str,
    stop_event: threading.Event,
    progress_callback: Optional[callable]
):
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
    image_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif'}
    
    try:
        while not stop_event.is_set():
            try:
                # 统计当前图片数量
                folder_path = Path(output_folder)
                if folder_path.exists():
                    current_count = sum(
                        1 for f in folder_path.iterdir()
                        if f.is_file() and f.suffix.lower() in image_extensions
                    )
                    
                    # 如果有新图片
                    if current_count > last_count:
                        progress_callback(t('conversion.progress.extracting_images', count=current_count))
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
    basename_for_output: str = None,
    cancel_event: Optional[threading.Event] = None,
    progress_callback: Optional[callable] = None
) -> Dict:
    """
    使用pymupdf4llm提取PDF内容（支持4种组合）
    
    优化方案：临时子文件夹 → 整体移动
    
    四种组合模式：
    1. ❌图片 ❌OCR：纯文本MD
    2. ✅图片 ❌OCR：MD + 图片
    3. ❌图片 ✅OCR：MD + OCR文本
    4. ✅图片 ✅OCR：MD + 图片 + OCR
    
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
    except ImportError:
        raise ImportError(
            "pymupdf4llm库未安装。\n"
            "请安装: pip install pymupdf4llm"
        )
    
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
    
    logger.info(f"开始提取PDF内容: {os.path.basename(pdf_path)}")
    logger.info(f"提取选项: 图片={extract_images}, OCR={extract_ocr}")
    
    # 检查取消
    if cancel_event and cancel_event.is_set():
        raise InterruptedError("操作已被取消")
    
    # 确定输出文件名（不含.md后缀）
    if not basename_for_output:
        basename_for_output = os.path.splitext(os.path.basename(pdf_path))[0]
    
    # 使用临时目录
    with tempfile.TemporaryDirectory() as temp_dir:
        logger.debug(f"临时目录: {temp_dir}")
        
        # 创建临时输出文件夹
        temp_output_folder = os.path.join(temp_dir, basename_for_output)
        os.makedirs(temp_output_folder, exist_ok=True)
        logger.debug(f"临时输出文件夹: {temp_output_folder}")
        
        # 情况1：纯文本（❌图片 ❌OCR）
        if not extract_images and not extract_ocr:
            logger.info("模式: 纯文本提取")
            
            md_text = pymupdf4llm.to_markdown(pdf_path, write_images=False)
            
            # 生成YAML头部（使用原始文件名，不含扩展名）
            original_file_stem = os.path.splitext(os.path.basename(pdf_path))[0]
            yaml_frontmatter = generate_basic_yaml_frontmatter(original_file_stem)
            
            # 保存MD文件（YAML头部 + 正文内容）
            md_filename = f"{basename_for_output}.md"
            md_path_temp = os.path.join(temp_output_folder, md_filename)
            with open(md_path_temp, 'w', encoding='utf-8') as f:
                f.write(yaml_frontmatter + md_text)
            
            logger.debug(f"已生成临时MD: {md_path_temp}")
            
            # 移动整个文件夹到目标位置
            final_folder = os.path.join(output_dir, basename_for_output)
            if os.path.exists(final_folder):
                shutil.rmtree(final_folder)
            
            shutil.move(temp_output_folder, final_folder)
            logger.info(f"已移动到: {final_folder}")
            
            final_md_path = os.path.join(final_folder, md_filename)
            
            return {
                'md_path': final_md_path,
                'folder_path': final_folder,
                'image_count': 0,
                'ocr_count': 0
            }
        
        # 情况2：文本+图片（✅图片 ❌OCR）
        elif extract_images and not extract_ocr:
            logger.info("模式: 文本+图片")
            
            # 启动图片提取监控线程
            stop_monitor = threading.Event()
            monitor_thread = None
            if progress_callback:
                monitor_thread = threading.Thread(
                    target=_monitor_image_extraction,
                    args=(temp_output_folder, stop_monitor, progress_callback),
                    daemon=True
                )
                monitor_thread.start()
                logger.debug("已启动图片提取监控线程")
            
            try:
                # pymupdf4llm直接输出到临时文件夹
                md_text = pymupdf4llm.to_markdown(
                    pdf_path,
                    write_images=True,
                    image_path=temp_output_folder,
                    image_format="png"
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
            image_count = _count_images_in_folder(temp_output_folder)
            logger.info(f"提取了{image_count}张图片")
            
            # 提示图片提取完成的总数
            if progress_callback:
                progress_callback(t('conversion.progress.images_extracted', count=image_count))
            
            # 替换MD中的绝对路径为相对路径，并应用配置的链接格式
            md_text = _convert_to_simple_paths(md_text)
            
            # 生成YAML头部（使用原始文件名，不含扩展名）
            original_file_stem = os.path.splitext(os.path.basename(pdf_path))[0]
            yaml_frontmatter = generate_basic_yaml_frontmatter(original_file_stem)
            
            # 保存MD文件（YAML头部 + 正文内容）
            md_filename = f"{basename_for_output}.md"
            md_path_temp = os.path.join(temp_output_folder, md_filename)
            with open(md_path_temp, 'w', encoding='utf-8') as f:
                f.write(yaml_frontmatter + md_text)
            
            logger.debug(f"已生成临时MD: {md_path_temp}")
            
            # 移动整个文件夹到目标位置
            final_folder = os.path.join(output_dir, basename_for_output)
            if os.path.exists(final_folder):
                shutil.rmtree(final_folder)
            
            shutil.move(temp_output_folder, final_folder)
            logger.info(f"已移动到: {final_folder}")
            
            final_md_path = os.path.join(final_folder, md_filename)
            
            return {
                'md_path': final_md_path,
                'folder_path': final_folder,
                'image_count': image_count,
                'ocr_count': 0
            }
        
        # 情况3：文本+OCR（❌图片 ✅OCR）
        elif not extract_images and extract_ocr:
            logger.info("模式: 文本+OCR（无图片保存）")
            
            # 启动图片提取监控线程
            stop_monitor = threading.Event()
            monitor_thread = None
            if progress_callback:
                monitor_thread = threading.Thread(
                    target=_monitor_image_extraction,
                    args=(temp_output_folder, stop_monitor, progress_callback),
                    daemon=True
                )
                monitor_thread.start()
                logger.debug("已启动图片提取监控线程")
            
            try:
                # 在临时文件夹中提取图片（用于OCR）
                md_text = pymupdf4llm.to_markdown(
                    pdf_path,
                    write_images=True,
                    image_path=temp_output_folder,
                    image_format="png"
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
            
            # 统计图片数量并提示
            temp_image_count = _count_images_in_folder(temp_output_folder)
            logger.info(f"提取了{temp_image_count}张图片用于OCR")
            if progress_callback:
                progress_callback(t('conversion.progress.images_extracted_preparing_ocr', count=temp_image_count))
            
            # 对图片进行OCR并替换图片链接为OCR文本
            md_text, ocr_count = _replace_images_with_ocr(
                md_text, temp_output_folder, cancel_event, progress_callback
            )
            logger.info(f"OCR识别完成: {ocr_count}个结果")
            
            # 删除图片文件（只保留OCR文本）
            _delete_images_in_folder(temp_output_folder)
            
            # 生成YAML头部（使用原始文件名，不含扩展名）
            original_file_stem = os.path.splitext(os.path.basename(pdf_path))[0]
            yaml_frontmatter = generate_basic_yaml_frontmatter(original_file_stem)
            
            # 保存MD文件（YAML头部 + 正文内容）
            md_filename = f"{basename_for_output}.md"
            md_path_temp = os.path.join(temp_output_folder, md_filename)
            with open(md_path_temp, 'w', encoding='utf-8') as f:
                f.write(yaml_frontmatter + md_text)
            
            logger.debug(f"已生成临时MD: {md_path_temp}")
            
            # 移动整个文件夹到目标位置
            final_folder = os.path.join(output_dir, basename_for_output)
            if os.path.exists(final_folder):
                shutil.rmtree(final_folder)
            
            shutil.move(temp_output_folder, final_folder)
            logger.info(f"已移动到: {final_folder}")
            
            final_md_path = os.path.join(final_folder, md_filename)
            
            return {
                'md_path': final_md_path,
                'folder_path': final_folder,
                'image_count': 0,
                'ocr_count': ocr_count
            }
        
        # 情况4：文本+图片+OCR（✅图片 ✅OCR）
        else:  # extract_images and extract_ocr
            logger.info("模式: 文本+图片+OCR")
            
            # 启动图片提取监控线程
            stop_monitor = threading.Event()
            monitor_thread = None
            if progress_callback:
                monitor_thread = threading.Thread(
                    target=_monitor_image_extraction,
                    args=(temp_output_folder, stop_monitor, progress_callback),
                    daemon=True
                )
                monitor_thread.start()
                logger.debug("已启动图片提取监控线程")
            
            try:
                # pymupdf4llm直接输出到临时文件夹
                md_text = pymupdf4llm.to_markdown(
                    pdf_path,
                    write_images=True,
                    image_path=temp_output_folder,
                    image_format="png"
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
            image_count = _count_images_in_folder(temp_output_folder)
            logger.info(f"提取了{image_count}张图片")
            
            # 提示图片提取完成的总数
            if progress_callback:
                progress_callback(t('conversion.progress.images_extracted_preparing_ocr', count=image_count))
            
            # 在图片后添加OCR文本
            md_text, ocr_count = _add_ocr_after_images(
                md_text, temp_output_folder, cancel_event, progress_callback
            )
            logger.info(f"OCR识别完成: {ocr_count}个结果")
            
            # 替换MD中的绝对路径为相对路径
            md_text = _convert_to_simple_paths(md_text)
            
            # 生成YAML头部（使用原始文件名，不含扩展名）
            original_file_stem = os.path.splitext(os.path.basename(pdf_path))[0]
            yaml_frontmatter = generate_basic_yaml_frontmatter(original_file_stem)
            
            # 保存MD文件（YAML头部 + 正文内容）
            md_filename = f"{basename_for_output}.md"
            md_path_temp = os.path.join(temp_output_folder, md_filename)
            with open(md_path_temp, 'w', encoding='utf-8') as f:
                f.write(yaml_frontmatter + md_text)
            
            logger.debug(f"已生成临时MD: {md_path_temp}")
            
            # 移动整个文件夹到目标位置
            final_folder = os.path.join(output_dir, basename_for_output)
            if os.path.exists(final_folder):
                shutil.rmtree(final_folder)
            
            shutil.move(temp_output_folder, final_folder)
            logger.info(f"已移动到: {final_folder}")
            
            final_md_path = os.path.join(final_folder, md_filename)
            
            return {
                'md_path': final_md_path,
                'folder_path': final_folder,
                'image_count': image_count,
                'ocr_count': ocr_count
            }


def _count_images_in_folder(folder_path: str) -> int:
    """
    统计文件夹中的图片数量
    
    参数:
        folder_path: 文件夹路径
        
    返回:
        int: 图片数量
    """
    count = 0
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            count += 1
    return count


def _delete_images_in_folder(folder_path: str):
    """
    删除文件夹中的所有图片文件
    
    参数:
        folder_path: 文件夹路径
    """
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            file_path = os.path.join(folder_path, filename)
            os.remove(file_path)
            logger.debug(f"删除图片: {filename}")


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
    
    # 查找所有图片链接 ![alt](path)
    pattern = r'!\[([^\]]*)\]\(([^)]+)\)'
    
    def replace_path(match):
        alt_text = match.group(1)
        img_path = match.group(2)
        
        # 提取纯文件名
        img_filename = os.path.basename(img_path)
        
        # 使用配置的格式
        return format_image_link(img_filename, image_link_style)
    
    result = re.sub(pattern, replace_path, md_text)
    
    # 统一使用正斜杠
    result = result.replace('\\', '/')
    
    return result


def _create_image_md_file(
    image_path: str,
    image_filename: str,
    output_folder: str,
    include_image: bool,
    include_ocr: bool,
    cancel_event: Optional[threading.Event] = None
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
    # 生成md文件名（与图片同名，扩展名改为.md）
    base_name = os.path.splitext(image_filename)[0]
    md_filename = f"{base_name}.md"
    md_path = os.path.join(output_folder, md_filename)
    
    try:
        lines = []
        
        # 包含图片链接（使用配置的链接格式）
        if include_image:
            from docwen.config.config_manager import config_manager
            from docwen.utils.markdown_utils import format_image_link
            link_settings = config_manager.get_markdown_link_style_settings()
            image_link_style = link_settings.get("image_link_style", "wiki_embed")
            lines.append(format_image_link(image_filename, image_link_style))
            lines.append("")  # 空行
        
        # 包含OCR识别文本
        if include_ocr:
            # 检查取消
            if cancel_event and cancel_event.is_set():
                raise InterruptedError("操作已被取消")
            
            logger.info(f"开始OCR识别: {image_filename}")
            from docwen.utils.ocr_utils import extract_text_simple
            # 传递cancel_event给OCR函数
            ocr_text = extract_text_simple(image_path, cancel_event)
            
            if ocr_text:
                lines.append(ocr_text)
                logger.info(f"OCR识别完成: {image_filename}, 识别出 {len(ocr_text)} 个字符")
        
        # 写入文件
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))
        
        logger.info(f"创建图片md文件: {md_filename}")
        return md_filename
        
    except Exception as e:
        logger.error(f"创建图片md文件失败: {md_filename}, 错误: {e}", exc_info=True)
        # 失败时返回图片链接作为兜底
        return image_filename


def _replace_images_with_ocr(
    md_text: str,
    images_folder: str,
    cancel_event: Optional[threading.Event] = None,
    progress_callback: Optional[callable] = None
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
    import re
    
    logger.info("开始OCR识别并生成图片md文件...")
    
    # 先找出所有图片链接，统计总数
    pattern = r'!\[([^\]]*)\]\(([^)]+)\)'
    image_matches = list(re.finditer(pattern, md_text))
    total_images = len(image_matches)
    
    logger.info(f"找到{total_images}张图片需要OCR")
    
    ocr_count = 0
    current_index = 0
    
    def replace_with_md_link(match):
        nonlocal ocr_count, current_index
        
        current_index += 1
        alt_text = match.group(1)
        img_path = match.group(2)
        
        # 检查取消（在处理每张图片前）
        if cancel_event and cancel_event.is_set():
            logger.info(f"OCR识别被取消，已完成 {ocr_count}/{total_images} 张")
            raise InterruptedError("操作已被取消")
        
        # 报告进度
        if progress_callback:
            progress_callback(t('conversion.progress.ocr_recognizing', current=current_index, total=total_images))
        
        # 构建完整图片路径
        if os.path.isabs(img_path):
            full_img_path = img_path
        else:
            img_filename = os.path.basename(img_path)
            full_img_path = os.path.join(images_folder, img_filename)
        
        if not os.path.exists(full_img_path):
            logger.warning(f"图片不存在: {full_img_path}")
            return "\n"
        
        # 创建图片md文件（只含OCR，不含图片链接）
        try:
            img_filename = os.path.basename(full_img_path)
            md_filename = _create_image_md_file(
                full_img_path,
                img_filename,
                images_folder,
                include_image=False,  # 不包含图片链接
                include_ocr=True,      # 包含OCR文本
                cancel_event=cancel_event
            )
            
            ocr_count += 1
            logger.debug(f"已为图片创建md文件: {md_filename}")
            
            # 使用配置的MD文件链接格式
            from docwen.config.config_manager import config_manager
            from docwen.utils.markdown_utils import format_md_file_link
            
            link_settings = config_manager.get_markdown_link_style_settings()
            md_file_link_style = link_settings.get("md_file_link_style", "wiki_embed")
            
            return format_md_file_link(md_filename, md_file_link_style)
            
        except InterruptedError:
            raise
        except ImportError:
            logger.warning("OCR功能不可用，请确保已安装PaddleOCR")
            return "\n"
        except Exception as e:
            logger.error(f"处理图片失败: {e}")
            return "\n"
    
    result = re.sub(pattern, replace_with_md_link, md_text)
    
    logger.info(f"图片md文件生成完成，共{ocr_count}个")
    
    return result, ocr_count


def _add_ocr_after_images(
    md_text: str,
    images_folder: str,
    cancel_event: Optional[threading.Event] = None,
    progress_callback: Optional[callable] = None
) -> tuple[str, int]:
    """
    在Markdown图片链接后添加图片md文件链接（含图片链接+OCR文本）
    
    用于情况4：保留图片，同时添加OCR文本
    现在改为与docx转MD一致：生成图片.md文件并插入链接
    
    参数:
        md_text: Markdown文本
        images_folder: 图片文件夹路径
        cancel_event: 取消事件(可选)
        progress_callback: 进度回调函数(可选)
        
    返回:
        tuple: (更新后的Markdown文本, OCR识别数量)
    """
    import re
    
    logger.info("开始OCR识别并生成图片md文件...")
    
    # 先找出所有图片链接，统计总数
    pattern = r'(!\[([^\]]*)\]\(([^)]+)\))'
    image_matches = list(re.finditer(pattern, md_text))
    total_images = len(image_matches)
    
    logger.info(f"找到{total_images}张图片需要OCR")
    
    ocr_count = 0
    current_index = 0
    
    def add_md_link(match):
        nonlocal ocr_count, current_index
        
        current_index += 1
        full_match = match.group(1)  # 完整的图片链接
        alt_text = match.group(2)
        img_path = match.group(3)
        
        # 检查取消（在处理每张图片前）
        if cancel_event and cancel_event.is_set():
            logger.info(f"OCR识别被取消，已完成 {ocr_count}/{total_images} 张")
            raise InterruptedError("操作已被取消")
        
        # 报告进度
        if progress_callback:
            progress_callback(t('conversion.progress.ocr_recognizing', current=current_index, total=total_images))
        
        # 构建完整图片路径
        if os.path.isabs(img_path):
            full_img_path = img_path
        else:
            img_filename = os.path.basename(img_path)
            full_img_path = os.path.join(images_folder, img_filename)
        
        if not os.path.exists(full_img_path):
            logger.warning(f"图片不存在: {full_img_path}")
            return full_match
        
        # 创建图片md文件（包含图片链接+OCR文本）
        try:
            img_filename = os.path.basename(full_img_path)
            md_filename = _create_image_md_file(
                full_img_path,
                img_filename,
                images_folder,
                include_image=True,   # 包含图片链接
                include_ocr=True,     # 包含OCR文本
                cancel_event=cancel_event
            )
            
            ocr_count += 1
            logger.debug(f"已为图片创建md文件: {md_filename}")
            
            # 使用配置的MD文件链接格式
            from docwen.config.config_manager import config_manager
            from docwen.utils.markdown_utils import format_md_file_link
            
            link_settings = config_manager.get_markdown_link_style_settings()
            md_file_link_style = link_settings.get("md_file_link_style", "wiki_embed")
            
            return format_md_file_link(md_filename, md_file_link_style)
            
        except InterruptedError:
            raise
        except ImportError:
            logger.warning("OCR功能不可用，请确保已安装PaddleOCR")
            return full_match
        except Exception as e:
            logger.error(f"处理图片失败: {e}")
            return full_match
    
    result = re.sub(pattern, add_md_link, md_text)
    
    logger.info(f"图片md文件生成完成，共{ocr_count}个")
    
    return result, ocr_count


# ==================== 公开API ====================

__all__ = [
    'extract_pdf_with_pymupdf4llm',
]
