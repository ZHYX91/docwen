"""
图片转PDF策略模块

将图片文件转换为PDF格式，支持多种质量模式。

依赖：
- img2pdf: PDF生成
- PIL: 图片处理
- formats.image: 格式转换
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
from .utils import get_image_format_description
from docwen.i18n import t

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_IMAGE, 'pdf')
class ImageToPdfStrategy(BaseStrategy):
    """
    将图片文件转换为PDF格式的策略
    
    支持三种质量模式：
    1. original - 原图嵌入（推荐）：保持原始分辨率，无损嵌入
    2. a4 - 适合A4纸：自动判断横向/纵向，适配210×297mm
    3. a3 - 适合A3纸：自动判断横向/纵向，适配297×420mm
    
    支持的输入格式：
    - 直接支持：JPG, PNG, TIFF, GIF, WEBP
    - 预处理后支持：BMP, HEIC (已通过批量列表自动转换为PNG)
    
    特殊处理：
    - 多页TIFF：自动检测并处理每一帧
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行图片到PDF的转换
        
        Args:
            file_path: 输入图片文件路径
            options: 转换选项，包含quality_mode参数
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 转换结果对象
        """
        try:
            if progress_callback:
                progress_callback(t('conversion.progress.converting_to_format', format='PDF'))
            
            import img2pdf
            from PIL import Image
            
            # 获取质量模式，默认为'original'
            quality_mode = options.get('quality_mode', 'original') if options else 'original'
            actual_format = options.get('actual_format') if options else None
            logger.info(f"图片转PDF质量模式: {quality_mode}")
            
            from docwen.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 使用标准的临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：创建输入副本 input.{ext}
                from docwen.utils.workspace_manager import prepare_input_file
                temp_input = prepare_input_file(file_path, temp_dir, actual_format or 'jpg')
                logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
                
                # 步骤2：检测并预处理特殊格式（从副本处理）
                file_to_convert = self._preprocess_image(temp_input, temp_dir, None, actual_format)
                
                # 记录中间文件（如果有格式转换）
                intermediate_file = None
                if actual_format in ['bmp', 'heic', 'heif'] and file_to_convert != temp_input:
                    intermediate_file = file_to_convert
                    logger.debug(f"记录中间文件: {os.path.basename(intermediate_file)}")
                
                # 读取图片信息，判断方向
                with Image.open(file_to_convert) as img:
                    width, height = img.size
                    is_landscape = width > height
                    logger.debug(f"图片尺寸: {width}x{height}, 横向: {is_landscape}")
                
                # 根据质量模式设置PDF布局
                layout_fun = self._create_layout(quality_mode, is_landscape)
                
                # 转换为PDF
                logger.debug(f"开始转换图片为PDF: {file_to_convert}")
                if layout_fun is None:
                    # 原图嵌入模式：不传递layout_fun参数
                    pdf_bytes = img2pdf.convert(file_to_convert)
                else:
                    # A4/A3模式：传递layout_fun参数
                    pdf_bytes = img2pdf.convert(file_to_convert, layout_fun=layout_fun)
                
                # 生成输出文件名
                description = get_image_format_description(actual_format) if actual_format else "fromImage"
                output_filename = os.path.basename(
                    generate_output_path(
                        file_path,
                        section="",
                        add_timestamp=True,
                        description=description,
                        file_type="pdf"
                    )
                )
                
                # 保存PDF到临时目录
                temp_output = os.path.join(temp_dir, output_filename)
                with open(temp_output, 'wb') as f:
                    f.write(pdf_bytes)
                
                # 准备最终输出路径
                output_path = os.path.join(output_dir, output_filename)
                
                # 处理中间文件保留
                should_keep = self._should_keep_intermediates()
                if should_keep:
                    # 保留中间文件（排除输入副本）
                    logger.info("保留中间文件，移动规范命名的文件到输出目录")
                    for filename in os.listdir(temp_dir):
                        # 排除输入副本文件
                        if filename.startswith('input.'):
                            logger.debug(f"跳过输入副本: {filename}")
                            continue
                        src = os.path.join(temp_dir, filename)
                        if os.path.isfile(src):
                            dst = os.path.join(output_dir, filename)
                            shutil.move(src, dst)
                            logger.debug(f"保留中间文件: {filename}")
                else:
                    # 只移动最终PDF文件
                    logger.debug("清理中间文件，只移动最终文件")
                    shutil.move(temp_output, output_path)
                
                logger.info(f"成功转换为PDF: {output_path}")
            
            return ConversionResult(
                success=True,
                output_path=output_path,
                message=t('conversion.messages.conversion_to_format_success', format='PDF')
            )
        
        except Exception as e:
            logger.error(f"图片转PDF失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=f"转换失败: {e}",
                error=e
            )
    
    def _preprocess_image(self, file_path: str, temp_dir: str, cancel_event=None, actual_format: str = None) -> str:
        """
        预处理图片：将非标准格式（BMP/HEIC/HEIF）转换为PNG
        
        Args:
            file_path: 原始文件路径
            temp_dir: 临时目录路径（必需）
            cancel_event: 取消事件（可选）
            actual_format: 实际文件格式（可选，如果不提供则自动检测）
            
        Returns:
            str: 处理后的文件路径
            
        说明:
            - 使用actual_format参数避免重复检测文件格式
            - 如果输入已经是标准格式（JPG/PNG/TIFF等），直接返回原路径
            - 如果是BMP/HEIC/HEIF，转换为PNG并保存到临时目录
        """
        # 如果没有提供actual_format，则检测
        if actual_format is None:
            from docwen.utils.file_type_utils import detect_actual_file_format
            actual_format = detect_actual_file_format(file_path)
            logger.debug(f"自动检测图片文件格式: {actual_format}")
        else:
            logger.debug(f"使用传入的文件格式: {actual_format}")
        
        # 标准格式：直接使用，无需转换
        if actual_format in ['jpeg', 'jpg', 'png', 'tiff', 'tif', 'gif', 'webp']:
            logger.debug(f"文件已是标准图片格式({actual_format})，无需转换: {file_path}")
            return file_path
        
        # BMP格式需要转换为PNG
        if actual_format == 'bmp':
            logger.info(f"检测到BMP格式，转换为PNG: {os.path.basename(file_path)}")
            
            try:
                from docwen.converter.formats.image import bmp_to_png
                
                converted_path = bmp_to_png(file_path, output_dir=temp_dir)
                logger.info(f"BMP转PNG成功: {os.path.basename(converted_path)}")
                return converted_path
                
            except Exception as e:
                logger.error(f"BMP转PNG失败: {e}")
                raise RuntimeError(f"BMP转PNG失败: {e}")
        
        # HEIC/HEIF格式需要转换为PNG
        elif actual_format in ['heic', 'heif']:
            logger.info(f"检测到{actual_format.upper()}格式，转换为PNG: {os.path.basename(file_path)}")
            
            try:
                from docwen.converter.formats.image import heic_to_png
                
                converted_path = heic_to_png(file_path, output_dir=temp_dir)
                logger.info(f"{actual_format.upper()}转PNG成功: {os.path.basename(converted_path)}")
                return converted_path
                
            except Exception as e:
                logger.error(f"{actual_format.upper()}转PNG失败: {e}")
                raise RuntimeError(f"{actual_format.upper()}转PNG失败: {e}")
        
        # 其他不支持的格式：直接返回原路径
        logger.warning(f"不支持的图片格式: {actual_format}，尝试直接使用")
        return file_path
    
    def _create_layout(self, quality_mode: str, is_landscape: bool):
        """
        根据质量模式创建PDF布局函数
        
        Args:
            quality_mode: 质量模式 ('original', 'a4', 'a3')
            is_landscape: 是否为横向图片
            
        Returns:
            布局函数或None
        """
        import img2pdf
        
        if quality_mode == 'original':
            # 原图嵌入模式：不设置页面尺寸
            logger.debug("使用原图嵌入模式")
            return None
        
        elif quality_mode == 'a4':
            # A4纸张尺寸：210×297mm
            if is_landscape:
                # 横向：宽297mm，高210mm
                pagesize = (img2pdf.mm_to_pt(297), img2pdf.mm_to_pt(210))
                logger.debug("使用A4横向布局")
            else:
                # 纵向：宽210mm，高297mm
                pagesize = (img2pdf.mm_to_pt(210), img2pdf.mm_to_pt(297))
                logger.debug("使用A4纵向布局")
            
            return img2pdf.get_layout_fun(pagesize)
        
        elif quality_mode == 'a3':
            # A3纸张尺寸：297×420mm
            if is_landscape:
                # 横向：宽420mm，高297mm
                pagesize = (img2pdf.mm_to_pt(420), img2pdf.mm_to_pt(297))
                logger.debug("使用A3横向布局")
            else:
                # 纵向：宽297mm，高420mm
                pagesize = (img2pdf.mm_to_pt(297), img2pdf.mm_to_pt(420))
                logger.debug("使用A3纵向布局")
            
            return img2pdf.get_layout_fun(pagesize)
        
        else:
            logger.warning(f"未知的质量模式: {quality_mode}，使用原图嵌入")
            return None
    
    @staticmethod
    def _should_keep_intermediates() -> bool:
        """判断是否应该保留中间文件"""
        try:
            from docwen.config.config_manager import config_manager
            return config_manager.get_save_intermediate_files()
        except Exception as e:
            logger.warning(f"读取中间文件配置失败: {e}，使用默认设置（不保存中间文件）")
            return False
