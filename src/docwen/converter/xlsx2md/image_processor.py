"""
XLSX图片提取模块

从XLSX表格文件中提取内嵌图片并保存到指定文件夹
与DOCX的image_processor模块保持一致的设计模式
"""

import os
import logging
from typing import List, Dict, Optional, Callable
from docwen.utils.path_utils import generate_output_path
from docwen.utils.markdown_utils import format_image_link, format_md_file_link
from docwen.i18n import t

logger = logging.getLogger(__name__)


def extract_images_from_xlsx(
    workbook, 
    output_folder: str, 
    original_file_path: str,
    progress_callback: Optional[Callable[[str], None]] = None
) -> List[dict]:
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
    os.makedirs(output_folder, exist_ok=True)
    
    try:
        # 使用传入的workbook对象，避免重复加载
        wb = workbook
        
        # 遍历所有工作表
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            # 检查工作表中的图片
            if hasattr(ws, '_images') and ws._images:
                logger.debug(f"工作表 '{sheet_name}' 包含 {len(ws._images)} 张图片")
                
                for img in ws._images:
                    try:
                        # 获取图片数据
                        image_data = img._data()
                        
                        # 确定图片格式
                        # openpyxl的图片对象可能有format属性
                        ext = 'png'  # 默认格式
                        if hasattr(img, 'format'):
                            ext = img.format.lower()
                        elif hasattr(img, '_format'):
                            ext = img._format.lower()
                        
                        # 使用统一的路径生成工具生成规范文件名
                        # 使用original_file_path进行命名，确保文件名基于原始文件而非临时副本
                        image_path = generate_output_path(
                            input_path=original_file_path,
                            output_dir=output_folder,
                            section=f"image{image_counter}",
                            add_timestamp=True,
                            description="fromXlsx",
                            file_type=ext
                        )
                        filename = os.path.basename(image_path)
                        
                        # 保存图片
                        with open(image_path, 'wb') as f:
                            f.write(image_data)
                        
                        # 获取图片位置
                        row, col = None, None
                        if hasattr(img, 'anchor') and img.anchor:
                            if hasattr(img.anchor, '_from'):
                                row = img.anchor._from.row + 1  # openpyxl从0开始，转为从1开始
                                col = img.anchor._from.col + 1
                        
                        logger.debug(f"提取图片: {filename} (工作表 '{sheet_name}', 行{row}, 列{col})")
                        
                        # 记录图片信息（包括位置）
                        img_info = {
                            'filename': filename,
                            'image_path': image_path,
                            'sheet_name': sheet_name
                        }
                        if row is not None and col is not None:
                            img_info['row'] = row
                            img_info['col'] = col
                        
                        images_info.append(img_info)
                        
                        # 进度回调
                        if progress_callback:
                            progress_callback(t('conversion.progress.extracting_images', count=image_counter))
                        
                        image_counter += 1
                        
                    except Exception as e:
                        logger.warning(f"提取图片失败 (工作表 '{sheet_name}'): {e}")
                        continue
        
        logger.info(f"XLSX图片提取完成，共提取 {len(images_info)} 张图片")
        
    except Exception as e:
        logger.error(f"提取XLSX图片时出错: {e}", exc_info=True)
    
    return images_info


def process_image_with_ocr(img: dict, keep_images: bool, enable_ocr: bool, output_folder: str, progress_callback: Optional[Callable[[str], None]] = None, current_index: int = 1, total_images: int = 1, cancel_event=None) -> str:
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
    
    filename = img['filename']
    image_path = img['image_path']
    
    logger.debug(f"处理表格图片: {filename}, 保留图片: {keep_images}, OCR: {enable_ocr}")
    
    # 获取链接格式配置
    link_settings = config_manager.get_markdown_link_style_settings()
    image_link_style = link_settings.get("image_link_style", "wiki_embed")
    md_file_link_style = link_settings.get("md_file_link_style", "wiki_embed")
    
    # 场景1：只勾选提取图片
    if keep_images and not enable_ocr:
        logger.info(f"场景1：只保留图片 - {filename}")
        return format_image_link(filename, image_link_style)
    
    # 场景2：只勾选OCR
    elif not keep_images and enable_ocr:
        logger.info(f"场景2：只OCR识别 - {filename}")
        # 创建图片md文件（只包含OCR文本）
        md_filename = create_image_md_file(image_path, filename, output_folder, 
                                            include_image=False, include_ocr=True, progress_callback=progress_callback,
                                            current_index=current_index, total_images=total_images, cancel_event=cancel_event)
        return format_md_file_link(md_filename, md_file_link_style)
    
    # 场景3：两者都勾选
    elif keep_images and enable_ocr:
        logger.info(f"场景3：图片 + OCR - {filename}")
        # 创建图片md文件（包含图片链接和OCR文本）
        md_filename = create_image_md_file(image_path, filename, output_folder, 
                                            include_image=True, include_ocr=True, progress_callback=progress_callback,
                                            current_index=current_index, total_images=total_images, cancel_event=cancel_event)
        return format_md_file_link(md_filename, md_file_link_style)
    
    # 都不勾选（不应该出现这种情况，但作为兜底）
    else:
        logger.warning(f"都不勾选时不应该调用此函数 - {filename}")
        return format_image_link(filename, image_link_style)


def create_image_md_file(
    image_path: str, 
    image_filename: str, 
    output_folder: str, 
    include_image: bool, 
    include_ocr: bool,
    progress_callback: Optional[Callable[[str], None]] = None,
    current_index: int = 1,
    total_images: int = 1,
    cancel_event=None
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
                logger.info("OCR识别被取消")
                return image_filename
            
            # 进度回调
            if progress_callback:
                progress_callback(t('conversion.progress.ocr_recognizing', current=current_index, total=total_images))
            
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


def replace_image_markers(
    markdown_text: str, 
    sheet_images: List[dict], 
    keep_images: bool, 
    enable_ocr: bool, 
    output_folder: str,
    progress_callback: Optional[Callable[[str], None]] = None,
    cancel_event=None
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
    import re
    
    # 计算需要OCR的图片总数
    total_images = len(sheet_images)
    current_index = 0
    
    # 查找所有图片标记: {{IMAGE:filename}}
    for img_info in sheet_images:
        # 检查取消（在处理每张图片前）
        if cancel_event and cancel_event.is_set():
            logger.info(f"图片标记替换被取消，已处理 {current_index}/{total_images} 张")
            break
        
        filename = img_info.get('filename')
        if not filename:
            continue
        
        # 创建标记模式
        marker = f"{{{{IMAGE:{filename}}}}}"
        
        # 如果Markdown中包含这个标记，替换为实际链接
        if marker in markdown_text:
            # 增加当前索引
            current_index += 1
            
            # 处理图片时传递progress_callback、索引信息和cancel_event
            image_link = process_image_with_ocr(img_info, keep_images, enable_ocr, output_folder, progress_callback, current_index, total_images, cancel_event)
            
            markdown_text = markdown_text.replace(marker, image_link)
            logger.debug(f"替换图片标记: {marker} -> {image_link}")
    
    return markdown_text


# 模块测试
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(name)s:%(lineno)d - %(message)s"
    )
    
    logger.info("=== XLSX图片提取模块测试 ===\n")
    
    # 测试文件
    test_xlsx = "test.xlsx"
    
    if os.path.exists(test_xlsx):
        import openpyxl
        import tempfile
        
        try:
            wb = openpyxl.load_workbook(test_xlsx, data_only=True)
            logger.info(f"成功加载工作簿: {test_xlsx}")
            
            with tempfile.TemporaryDirectory() as temp_dir:
                output_folder = os.path.join(temp_dir, "test_images")
                
                images_info = extract_images_from_xlsx(wb, output_folder, test_xlsx)
                
                logger.info(f"\n提取结果:")
                logger.info(f"  总计: {len(images_info)} 张图片")
                
                for img in images_info:
                    logger.info(f"  - {img['filename']} (工作表 '{img['sheet_name']}')")
                    logger.info(f"    路径: {img['image_path']}")
                    logger.info(f"    文件存在: {os.path.exists(img['image_path'])}")
        
        except Exception as e:
            logger.error(f"测试失败: {e}", exc_info=True)
    else:
        logger.warning(f"测试文件不存在: {test_xlsx}")
    
    logger.info("\n=== 测试结束 ===")
