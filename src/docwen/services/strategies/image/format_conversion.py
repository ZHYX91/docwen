"""
图片格式互转策略模块

支持多种图片格式之间的相互转换。

依赖：
- PIL: 图片处理
- formats.image: 格式转换函数
"""

import os
import logging
import tempfile
from typing import Dict, Any, Callable, Optional

from PIL import Image

from ..base_strategy import BaseStrategy
from docwen.services.result import ConversionResult
from .. import register_conversion, CATEGORY_IMAGE
from docwen.utils.path_utils import generate_output_path
from .utils import get_image_format_description, is_multipage_tiff
from docwen.i18n import t

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_IMAGE, CATEGORY_IMAGE)
class ImageFormatConversionStrategy(BaseStrategy):
    """
    通用图片格式转换策略
    
    支持以下格式的相互转换：
    - JPEG/JPG ↔ PNG ↔ GIF ↔ BMP ↔ TIFF ↔ WebP
    
    使用 Pillow 库进行格式转换，自动处理：
    - 透明度转换（RGBA → RGB，白色背景）
    - 颜色模式转换（P, LA 等 → RGB）
    - 多页TIFF（仅保留第一页）
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行图片格式转换
        
        标准化工作流程：
        1. 在临时目录创建输入文件副本（使用 prepare_input_file）
        2. 从副本读取并转换格式
        3. 保存到最终输出目录
        4. 自动清理临时目录
        
        多页TIFF处理：
        - 检测到多页TIFF时，拆分每一页并分别转换
        - 每页使用标准化命名：原始名_page{N}_{时间戳}_from{格式}.{目标格式}
        - 所有页面共享同一时间戳
        
        Args:
            file_path: 输入图片文件路径
            options: 转换选项，包含actual_format和target_format
            progress_callback: 进度更新回调函数
            
        Returns:
            ConversionResult: 转换结果对象
        """
        try:
            # 获取实际格式
            actual_format = options.get('actual_format') if options else None
            if not actual_format:
                from docwen.utils.file_type_utils import detect_actual_file_format
                actual_format = detect_actual_file_format(file_path)
            
            # 从options中获取目标格式
            target_format = options.get('target_format') if options else None
            
            if not target_format:
                logger.warning("未能从选项中获取目标格式")
                return ConversionResult(
                    success=False,
                    message=t('conversion.messages.conversion_failed_no_target')
                )
            
            logger.info(f"图片格式转换: {actual_format.upper()} → {target_format.upper()}")
            
            if progress_callback:
                progress_callback(t('conversion.progress.converting_format_to_target', format=actual_format.upper(), target=target_format.upper()))
            
            from docwen.utils.workspace_manager import get_output_directory
            output_dir = get_output_directory(file_path)
            
            # 保存options到实例变量，供子方法使用
            self._current_options = options
            
            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 创建输入副本
                from docwen.utils.workspace_manager import prepare_input_file
                temp_input = prepare_input_file(file_path, temp_dir, actual_format)
                logger.debug(f"已创建输入副本: {os.path.basename(temp_input)}")
                
                # 检测是否为多页TIFF
                if is_multipage_tiff(temp_input):
                    # 多页TIFF：拆分每一页并分别转换
                    return self._convert_multipage_tiff(
                        temp_input, file_path, actual_format, target_format,
                        output_dir, progress_callback
                    )
                else:
                    # 单页图片：直接转换
                    return self._convert_single_image(
                        temp_input, file_path, actual_format, target_format,
                        output_dir, progress_callback
                    )
        
        except Exception as e:
            logger.error(f"图片格式转换失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=t('conversion.messages.conversion_failed_with_error', error=str(e)),
                error=e
            )
    
    def _convert_single_image(
        self,
        temp_input: str,
        original_path: str,
        actual_format: str,
        target_format: str,
        output_dir: str,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        转换单页图片（支持压缩选项）
        
        Args:
            temp_input: 临时输入文件路径
            original_path: 原始文件路径
            actual_format: 实际格式
            target_format: 目标格式
            output_dir: 输出目录
            progress_callback: 进度回调
            
        Returns:
            ConversionResult: 转换结果
        """
        # 使用新的convert_image函数，支持压缩选项
        from docwen.converter.formats.image import convert_image
        
        # 策略层负责生成输出路径（使用原始文件路径）
        description = get_image_format_description(actual_format)
        output_path = generate_output_path(
            original_path,  # 使用原始文件路径生成文件名
            output_dir=output_dir,
            section="",
            add_timestamp=True,
            description=description,
            file_type=target_format
        )
        
        # 准备压缩选项
        compress_options = {}
        
        # 从父方法的options中提取压缩选项（通过self临时存储）
        if hasattr(self, '_current_options') and self._current_options:
            compress_mode = self._current_options.get('compress_mode', 'lossless')
            compress_options['compress_mode'] = compress_mode
            
            if compress_mode == 'limit_size':
                compress_options['size_limit'] = self._current_options.get('size_limit')
                compress_options['size_unit'] = self._current_options.get('size_unit', 'KB')
                logger.info(f"使用压缩模式: {compress_options['size_limit']}{compress_options['size_unit']}")
            else:
                logger.info("使用最高质量模式")
        else:
            logger.info("未提供压缩选项，使用默认最高质量模式")
            compress_options['compress_mode'] = 'lossless'
        
        try:
            # 调用convert_image函数执行转换
            result_path = convert_image(
                source_path=temp_input,    # 临时副本用于读取
                target_format=target_format,
                output_path=output_path,   # 基于原始文件名的输出路径
                options=compress_options
            )
            
            logger.info(f"✓ 格式转换成功: {os.path.basename(result_path)}")
            
            # 标准化扩展名用于消息
            file_ext = self._normalize_extension(target_format)
            
            return ConversionResult(
                success=True,
                output_path=result_path,
                message=t('conversion.messages.conversion_to_format_success', format=file_ext.upper())
            )
        
        except Exception as e:
            logger.error(f"图片转换失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=t('conversion.messages.conversion_failed_with_error', error=str(e)),
                error=e
            )
    
    def _convert_multipage_tiff(
        self,
        temp_input: str,
        original_path: str,
        actual_format: str,
        target_format: str,
        output_dir: str,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        转换多页TIFF，拆分每一页并分别转换
        
        Args:
            temp_input: 临时输入文件路径
            original_path: 原始文件路径
            actual_format: 实际格式
            target_format: 目标格式
            output_dir: 输出目录
            progress_callback: 进度回调
            
        Returns:
            ConversionResult: 转换结果（返回第一页的路径）
        """
        output_files = []
        
        # 标准化文件扩展名
        file_ext = self._normalize_extension(target_format)
        pil_format = self._get_pil_format(target_format)
        
        try:
            with Image.open(temp_input) as img:
                n_frames = getattr(img, 'n_frames', 1)
                logger.info(f"检测到多页TIFF（共{n_frames}页），开始拆分转换")
                
                for i in range(n_frames):
                    img.seek(i)
                    page_num = i + 1
                    
                    # 更新进度
                    if progress_callback:
                        progress_callback(t('conversion.progress.converting_page_to_format', page=page_num, format=file_ext.upper()))
                    
                    # 生成当前页的输出路径
                    # 使用 generate_output_path，section参数标识页码
                    description = get_image_format_description(actual_format)
                    page_output_path = generate_output_path(
                        original_path,
                        output_dir=output_dir,
                        section=f'page{page_num}',
                        add_timestamp=True,
                        description=description,
                        file_type=file_ext
                    )
                    
                    logger.debug(f"第{page_num}页输出路径: {os.path.basename(page_output_path)}")
                    
                    # 复制当前帧
                    frame = img.copy()
                    
                    # 处理颜色模式
                    converted_frame = self._convert_image_mode(frame, target_format)
                    
                    # 保存为目标格式
                    save_kwargs = self._get_save_kwargs(target_format)
                    converted_frame.save(page_output_path, format=pil_format, **save_kwargs)
                    
                    output_files.append(page_output_path)
                    logger.info(f"✓ 第{page_num}页转换成功: {os.path.basename(page_output_path)}")
            
            logger.info(f"多页TIFF转换完成，共转换{len(output_files)}页")
            
            if progress_callback:
                progress_callback(t('conversion.progress.tiff_completed', count=len(output_files)))
            
            # 返回第一页的路径作为主输出
            return ConversionResult(
                success=True,
                output_path=output_files[0] if output_files else None,
                message=t('conversion.messages.conversion_to_format_pages_success', format=file_ext.upper(), count=len(output_files))
            )
        
        except Exception as e:
            logger.error(f"多页TIFF转换失败: {e}", exc_info=True)
            # 清理已创建的文件
            for output_file in output_files:
                try:
                    if os.path.exists(output_file):
                        os.remove(output_file)
                        logger.debug(f"已删除: {output_file}")
                except Exception as cleanup_error:
                    logger.warning(f"清理文件失败: {output_file}, 错误: {cleanup_error}")
            raise
    
    def _convert_image_mode(self, img, target_format: str):
        """
        根据目标格式转换图片颜色模式
        
        Args:
            img: PIL Image对象
            target_format: 目标格式（jpeg, png, gif等）
            
        Returns:
            转换后的PIL Image对象
        """
        target_format = target_format.lower()
        
        # JPEG不支持透明度，需要转为RGB
        if target_format in ['jpeg', 'jpg']:
            if img.mode in ('RGBA', 'LA', 'P'):
                logger.debug(f"JPEG不支持透明度，{img.mode} → RGB（白色背景）")
                # 创建白色背景
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                if img.mode in ('RGBA', 'LA'):
                    background.paste(img, mask=img.split()[-1])  # 使用alpha通道
                else:
                    background.paste(img)
                return background
            elif img.mode != 'RGB':
                logger.debug(f"{img.mode} → RGB")
                return img.convert('RGB')
            return img
        
        # GIF支持透明度但只有256色
        elif target_format == 'gif':
            if img.mode not in ('P', 'L'):
                logger.debug(f"{img.mode} → P（256色）")
                return img.convert('P')
            return img
        
        # BMP不支持透明度
        elif target_format == 'bmp':
            if img.mode in ('RGBA', 'LA', 'P'):
                logger.debug(f"BMP不支持透明度，{img.mode} → RGB（白色背景）")
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                if img.mode in ('RGBA', 'LA'):
                    background.paste(img, mask=img.split()[-1])
                else:
                    background.paste(img)
                return background
            elif img.mode != 'RGB':
                logger.debug(f"{img.mode} → RGB")
                return img.convert('RGB')
            return img
        
        # PNG和WebP支持透明度，保持RGBA
        elif target_format in ['png', 'webp']:
            if img.mode in ('P', 'L', 'LA'):
                logger.debug(f"{img.mode} → RGBA")
                return img.convert('RGBA')
            return img
        
        # TIFF支持多种模式，保持原样或转RGB
        elif target_format in ['tiff', 'tif']:
            if img.mode in ('P', 'LA'):
                logger.debug(f"{img.mode} → RGB")
                return img.convert('RGB')
            return img
        
        # 默认：保持原样
        return img
    
    def _normalize_extension(self, format_name: str) -> str:
        """
        标准化文件扩展名（统一使用简短格式）
        
        Args:
            format_name: 格式名称
            
        Returns:
            标准化的扩展名
        """
        format_map = {
            'jpeg': 'jpg',
            'jpg': 'jpg',
            'tiff': 'tif',
            'tif': 'tif',
        }
        return format_map.get(format_name.lower(), format_name.lower())
    
    def _get_pil_format(self, format_name: str) -> str:
        """
        获取PIL库使用的格式名称
        
        Args:
            format_name: 格式名称
            
        Returns:
            PIL格式名称（大写）
        """
        # PIL使用JPEG和TIFF作为格式名称
        format_map = {
            'jpeg': 'JPEG',
            'jpg': 'JPEG',
            'tiff': 'TIFF',
            'tif': 'TIFF',
            'png': 'PNG',
            'gif': 'GIF',
            'bmp': 'BMP',
            'webp': 'WEBP',
        }
        return format_map.get(format_name.lower(), format_name.upper())
    
    def _get_save_kwargs(self, target_format: str) -> dict:
        """
        获取保存图片时的参数
        
        Args:
            target_format: 目标格式
            
        Returns:
            保存参数字典
        """
        target_format = target_format.lower()
        
        if target_format in ['jpeg', 'jpg']:
            return {'quality': 95, 'optimize': True}
        elif target_format == 'png':
            return {'optimize': True}
        elif target_format == 'webp':
            return {'quality': 95, 'method': 6}  # method=6 最高质量
        elif target_format in ['tiff', 'tif']:
            return {'compression': 'tiff_lzw'}  # 使用LZW无损压缩
        else:
            return {}
