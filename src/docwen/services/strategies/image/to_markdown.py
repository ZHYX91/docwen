"""
图片转Markdown策略模块

将图片文件转换为Markdown格式，支持OCR文字识别。

依赖：
- PIL: 图片处理
- ocr_utils: OCR识别
- yaml_utils: YAML头部生成
"""

import os
import logging
import tempfile
import shutil
from typing import Dict, Any, Callable, Optional

from ..base_strategy import BaseStrategy
from docwen.services.result import ConversionResult
from .. import register_conversion, CATEGORY_IMAGE
from docwen.utils.path_utils import generate_output_path
from docwen.utils.yaml_utils import generate_basic_yaml_frontmatter
from docwen.i18n import t
from .utils import get_image_format_description, is_multipage_tiff, extract_tiff_pages

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_IMAGE, 'md')
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
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
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
            actual_format = options.get('actual_format') if options else None
            
            from docwen.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            original_basename = os.path.splitext(os.path.basename(file_path))[0]
            
            with tempfile.TemporaryDirectory() as temp_dir:
                # === 步骤1：创建输入副本 input.{ext} ===
                if progress_callback:
                    progress_callback(t('conversion.progress.preparing'))
                
                from docwen.utils.workspace_manager import prepare_input_file
                temp_image = prepare_input_file(file_path, temp_dir, actual_format or 'jpg')
                logger.debug(f"已创建输入副本: {os.path.basename(temp_image)}")
                
                # === 步骤2：生成统一的basename ===
                description = get_image_format_description(actual_format) if actual_format else 'fromImage'
                output_path_with_ext = generate_output_path(
                    file_path, None, "", True, description, 'md'
                )
                basename = os.path.splitext(os.path.basename(output_path_with_ext))[0]
                logger.info(f"统一basename: {basename}")
                
                # === 步骤3：创建子文件夹 ===
                folder_path = os.path.join(temp_dir, basename)
                os.makedirs(folder_path)
                logger.debug(f"子文件夹: {folder_path}")
                
                # === 步骤4：处理图片到子文件夹 ===
                images_in_folder = []  # [(图片路径, 图片文件名), ...]
                
                if is_multipage_tiff(file_path):
                    # 情况1：多页TIFF - 使用标准化命名直接提取到子文件夹
                    logger.info("检测到多页TIFF，使用标准化命名拆分")
                    pages = extract_tiff_pages(
                        file_path=temp_image,
                        output_dir=folder_path,
                        actual_format=actual_format or 'tiff',
                        progress_callback=progress_callback
                    )
                    images_in_folder = [(path, os.path.basename(path)) for _, path in pages]
                    logger.debug(f"多页TIFF拆分完成，共 {len(pages)} 页")
                
                elif actual_format in ['heic', 'heif', 'bmp']:
                    # 情况2：需要转换格式 - 直接转换到子文件夹
                    if progress_callback:
                        progress_callback(t('conversion.progress.converting_format_to_png', format=actual_format.upper()))
                    
                    logger.info(f"{actual_format.upper()}转换为PNG")
                    image_filename = f'{original_basename}.png'
                    final_image_path = os.path.join(folder_path, image_filename)
                    
                    from docwen.converter.formats.image import heic_to_png, bmp_to_png
                    if actual_format in ['heic', 'heif']:
                        heic_to_png(temp_image, output_path=final_image_path)
                    else:  # bmp
                        bmp_to_png(temp_image, output_path=final_image_path)
                    
                    images_in_folder = [(final_image_path, image_filename)]
                
                else:
                    # 情况3：标准格式 - 移动到子文件夹
                    if progress_callback:
                        progress_callback(t('conversion.progress.processing_images'))
                    
                    logger.info(f"标准{actual_format.upper()}格式")
                    image_filename = f'{original_basename}.{actual_format}'
                    final_image_path = os.path.join(folder_path, image_filename)
                    shutil.move(temp_image, final_image_path)
                    
                    images_in_folder = [(final_image_path, image_filename)]
                
                # === 步骤5：统一OCR子文件夹的所有图片并生成MD内容 ===
                # 生成YAML头部（使用原始文件名，不含扩展名）
                md_content = generate_basic_yaml_frontmatter(original_basename)
                
                try:
                    from docwen.utils.ocr_utils import extract_text_simple
                    ocr_available = True
                except ImportError:
                    logger.warning("PaddleOCR未安装，跳过文字识别")
                    ocr_available = False
                
                # 获取导出选项
                extract_image = options.get('extract_image', True) if options else True
                extract_ocr = options.get('extract_ocr', True) if options else True
                
                logger.info(f"导出选项 - 提取图片: {extract_image}, OCR: {extract_ocr}")
                
                # 获取链接格式配置
                from docwen.config.config_manager import config_manager
                from docwen.utils.markdown_utils import format_image_link
                
                link_settings = config_manager.get_markdown_link_style_settings()
                image_link_style = link_settings.get("image_link_style", "wiki_embed")
                
                # 计算总图片数（用于进度显示）
                total_images = len(images_in_folder)
                current_index = 0
                
                # 获取cancel_event（如果有）
                cancel_event = options.get('cancel_event') if options else None
                
                for img_path, img_filename in images_in_folder:
                    # 检查取消事件（在处理每张图片前）
                    if cancel_event and cancel_event.is_set():
                        logger.info("用户取消操作，停止OCR识别")
                        return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                    
                    # 根据选项添加图片链接
                    if extract_image:
                        image_link = format_image_link(img_filename, image_link_style)
                        md_content += f"{image_link}\n\n"
                    
                    # 根据选项进行OCR识别（每一页都识别）
                    if extract_ocr and ocr_available:
                        # 增加当前索引
                        current_index += 1
                        
                        if progress_callback:
                            progress_callback(t('conversion.progress.ocr_recognizing', current=current_index, total=total_images))
                        
                        try:
                            # 传递cancel_event给OCR函数
                            ocr_text = extract_text_simple(img_path, cancel_event)
                            if ocr_text:
                                md_content += f"{ocr_text}\n\n"
                                logger.debug(f"OCR识别成功: {img_filename}")
                        except Exception as e:
                            logger.warning(f"OCR识别失败: {e}")
                
                # === 步骤6：保存MD文件 ===
                if progress_callback:
                    progress_callback(t('conversion.progress.writing_file'))
                
                md_filename = f'{basename}.md'
                md_path = os.path.join(folder_path, md_filename)
                with open(md_path, 'w', encoding='utf-8') as f:
                    f.write(md_content)
                logger.info(f"MD文件已生成: {md_path}")
                
                # === 步骤7：移动子文件夹到目标目录 ===
                final_folder = os.path.join(output_dir, basename)
                if os.path.exists(final_folder):
                    shutil.rmtree(final_folder)
                shutil.move(folder_path, final_folder)
                logger.info(f"子文件夹已移动到: {final_folder}")
                
                return ConversionResult(
                    success=True,
                    output_path=os.path.join(final_folder, md_filename),
                    message=t('conversion.messages.conversion_to_format_success', format='Markdown')
                )
        
        except Exception as e:
            logger.error(f"图片转Markdown失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=t('conversion.messages.conversion_failed_check_log'),
                error=e
            )
