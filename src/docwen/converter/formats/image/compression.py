"""
图片压缩相关功能模块

提供图片压缩和文件大小计算功能：
- 文件大小工具函数
- 图片压缩到指定大小
- 保存参数生成
"""

import os
import io
import logging
from PIL import Image

logger = logging.getLogger(__name__)


# ============================================
# 文件大小工具函数
# ============================================

def get_file_size(file_path: str) -> int:
    """
    获取文件大小（字节）
    
    参数:
        file_path: 文件路径
        
    返回:
        int: 文件大小（字节）
    """
    return os.path.getsize(file_path)


def bytes_to_kb(size_bytes: int) -> float:
    """
    字节转KB
    
    参数:
        size_bytes: 字节数
        
    返回:
        float: KB数
    """
    return size_bytes / 1024


def bytes_to_mb(size_bytes: int) -> float:
    """
    字节转MB
    
    参数:
        size_bytes: 字节数
        
    返回:
        float: MB数
    """
    return size_bytes / (1024 * 1024)


def should_compress(source_size: int, target_size: int, unit: str) -> bool:
    """
    判断是否需要压缩
    
    参数:
        source_size: 源文件大小（字节）
        target_size: 目标大小（数值）
        unit: 单位（KB或MB）
    
    返回:
        bool: True需要压缩，False无需压缩（直接无损）
    """
    if unit == 'KB':
        target_bytes = target_size * 1024
    else:  # MB
        target_bytes = target_size * 1024 * 1024
    
    return source_size > target_bytes


# ============================================
# 压缩相关函数
# ============================================

def get_save_params(target_format: str, quality: int = 95) -> dict:
    """
    获取保存参数
    
    参数:
        target_format: 目标格式（小写，如'jpeg', 'webp', 'png'等）
        quality: 质量参数（15-95，仅对JPEG/WebP有效）
    
    返回:
        dict: 保存参数字典
    """
    fmt_upper = target_format.upper()
    
    if target_format in ['jpg', 'jpeg']:
        return {
            'format': 'JPEG',
            'quality': quality,
            'optimize': True
        }
    elif target_format == 'webp':
        return {
            'format': 'WEBP',
            'quality': quality,
            'method': 6  # 最佳压缩方法
        }
    elif target_format == 'png':
        return {
            'format': 'PNG',
            'compress_level': 9,  # 最高压缩级别
            'optimize': True
        }
    elif target_format == 'bmp':
        return {
            'format': 'BMP'
        }
    elif target_format == 'gif':
        return {
            'format': 'GIF',
            'optimize': True
        }
    elif target_format in ['tif', 'tiff']:
        return {
            'format': 'TIFF',
            'compression': 'tiff_lzw'  # LZW压缩
        }
    else:
        # 默认参数
        return {'format': fmt_upper}


def compress_to_size(img: Image.Image, output_path: str, target_format: str, 
                     target_size: int, unit: str) -> bool:
    """
    压缩图片到指定大小（使用二分查找）
    
    仅适用于JPEG和WebP格式，这些格式支持有损压缩quality参数
    
    参数:
        img: PIL Image对象
        output_path: 输出路径
        target_format: 目标格式（小写）
        target_size: 目标大小
        unit: 单位（KB/MB）
    
    返回:
        bool: 是否成功
    """
    # 转换目标大小为字节
    if unit == 'KB':
        target_bytes = target_size * 1024
    else:
        target_bytes = target_size * 1024 * 1024
    
    logger.info(f"开始压缩图片到目标大小: {target_size}{unit} ({target_bytes}字节)")
    
    # 二分查找最优质量（15-95之间）
    quality_min = 15
    quality_max = 95
    best_quality = quality_max
    best_buffer = None
    
    while quality_min <= quality_max:
        quality = (quality_min + quality_max) // 2
        
        # 保存到临时缓冲区
        buffer = io.BytesIO()
        save_params = get_save_params(target_format, quality)
        
        try:
            img.save(buffer, **save_params)
            current_size = buffer.tell()
            
            logger.debug(f"尝试quality={quality}, 大小={bytes_to_kb(current_size):.2f}KB")
            
            if current_size <= target_bytes:
                # 满足要求，尝试提高质量
                best_quality = quality
                best_buffer = buffer
                quality_min = quality + 1
            else:
                # 超出大小，降低质量
                quality_max = quality - 1
        except Exception as e:
            logger.warning(f"quality={quality}时保存失败: {e}")
            quality_max = quality - 1
            continue
    
    # 使用最优质量保存
    if best_buffer:
        # 使用缓冲区保存（避免重新压缩）
        with open(output_path, 'wb') as f:
            f.write(best_buffer.getvalue())
        final_size = best_buffer.tell()
        logger.info(f"压缩完成: quality={best_quality}, 最终大小={bytes_to_kb(final_size):.2f}KB")
    else:
        # 降级方案：使用最低质量保存
        logger.warning("无法满足目标大小，使用最低质量(15)保存")
        save_params = get_save_params(target_format, 15)
        img.save(output_path, **save_params)
    
    return True
