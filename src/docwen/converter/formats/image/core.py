"""
图片格式转换核心模块

提供通用图片格式转换功能，支持：
- 多种图片格式互转（PNG/JPEG/WebP/BMP/GIF/TIFF）
- 最高质量模式（无损或接近无损）
- 压缩模式（限制文件大小）
"""

import os
import logging
from PIL import Image
from .compression import (
    get_file_size,
    bytes_to_kb,
    should_compress,
    get_save_params,
    compress_to_size,
)

logger = logging.getLogger(__name__)


def convert_image(source_path: str, target_format: str, output_path: str, options: dict = None) -> str:
    """
    转换图片格式（支持压缩选项）
    
    参数:
        source_path: 源文件路径（用于读取）
        target_format: 目标格式（小写，如'jpeg', 'png', 'webp'等）
        output_path: 输出文件路径（完整路径，由调用方提供）
        options: 转换选项，包括：
            - compress_mode: 'lossless' (最高质量) 或 'limit_size' (限制大小)
            - size_limit: 文件大小上限（数值，仅compress_mode='limit_size'时有效）
            - size_unit: 单位 'KB' 或 'MB'（仅compress_mode='limit_size'时有效）
    
    返回:
        str: 转换后的文件路径
        
    异常:
        RuntimeError: 转换失败时抛出
    """
    try:
        options = options or {}
        compress_mode = options.get('compress_mode', 'lossless')
        
        logger.info(f"开始转换图片: {os.path.basename(source_path)} → {target_format.upper()}")
        logger.debug(f"压缩模式: {compress_mode}")
        
        # 打开图片
        img = Image.open(source_path)
        logger.debug(f"原始图片: 模式={img.mode}, 尺寸={img.size}")
        
        # 确保输出目录存在
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # 处理色彩模式（JPEG不支持透明通道）
        if target_format in ['jpg', 'jpeg'] and img.mode in ['RGBA', 'LA', 'P']:
            logger.debug(f"JPEG不支持透明通道，将{img.mode}转换为RGB")
            # 创建白色背景
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            if img.mode in ['RGBA', 'LA']:
                background.paste(img, mask=img.split()[-1])  # 使用alpha通道
            else:
                background.paste(img)
            img = background
        elif img.mode not in ['RGB', 'RGBA', 'L']:
            # 其他模式转换为RGB
            logger.debug(f"将{img.mode}模式转换为RGB")
            img = img.convert('RGB')
        
        # 根据压缩模式处理
        if compress_mode == 'lossless':
            # 最高质量模式
            logger.info("使用最高质量模式")
            
            if target_format in ['jpg', 'jpeg', 'webp']:
                # JPEG/WebP: 使用quality=95
                save_params = get_save_params(target_format, quality=95)
                logger.debug(f"保存参数: {save_params}")
                img.save(output_path, **save_params)
            else:
                # 其他格式: 无损保存
                save_params = get_save_params(target_format)
                logger.debug(f"保存参数: {save_params}")
                img.save(output_path, **save_params)
        
        elif compress_mode == 'limit_size':
            # 压缩模式
            size_limit = options.get('size_limit')
            size_unit = options.get('size_unit', 'KB')
            
            if not size_limit:
                raise ValueError("压缩模式下必须提供size_limit参数")
            
            logger.info(f"使用压缩模式: 目标大小={size_limit}{size_unit}")
            
            # 检查是否需要压缩
            source_size = get_file_size(source_path)
            logger.debug(f"源文件大小: {bytes_to_kb(source_size):.2f}KB")
            
            if should_compress(source_size, size_limit, size_unit):
                # 需要压缩
                logger.info("源文件超过目标大小，开始压缩")
                
                if target_format in ['jpg', 'jpeg', 'webp']:
                    # JPEG/WebP支持有损压缩
                    compress_to_size(img, output_path, target_format, size_limit, size_unit)
                else:
                    # 其他格式不支持有效压缩，使用无损保存
                    logger.warning(f"{target_format.upper()}格式不支持有效压缩，使用无损保存")
                    save_params = get_save_params(target_format)
                    img.save(output_path, **save_params)
            else:
                # 无需压缩，直接无损转换
                logger.info("源文件已满足目标大小，使用无损转换")
                save_params = get_save_params(target_format)
                img.save(output_path, **save_params)
        
        else:
            raise ValueError(f"未知的压缩模式: {compress_mode}")
        
        # 记录最终文件大小
        final_size = get_file_size(output_path)
        logger.info(f"转换完成: {os.path.basename(output_path)} ({bytes_to_kb(final_size):.2f}KB)")
        
        img.close()
        return output_path
        
    except Exception as e:
        logger.error(f"图片转换失败: {e}", exc_info=True)
        raise RuntimeError(f"图片转换失败: {e}")


