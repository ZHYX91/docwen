"""
图片格式转换支持模块
支持HEIC/HEIF等特殊图片格式转换为通用格式
支持通用图片格式之间的转换和压缩
"""

import os
import io
import logging
from gongwen_converter.utils.path_utils import generate_output_path
from typing import Optional, Dict
from PIL import Image

# 配置日志
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


# ============================================
# 主转换函数
# ============================================

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


# ============================================
# 特殊格式转换函数（兼容旧代码）
# ============================================

def heic_to_png(
    heic_path: str, 
    output_path: str = None,
    output_dir: Optional[str] = None
) -> str:
    """
    将HEIC/HEIF格式转换为PNG格式
    
    HEIC (High Efficiency Image Container) 是苹果设备使用的图片格式，
    Windows系统默认不支持，需要转换为通用格式。
    
    参数:
        heic_path: HEIC/HEIF文件路径
        output_path: 输出文件的完整路径（优先使用）
        output_dir: 输出目录（可选）。如果output_path未指定，则使用此参数生成路径
        
    返回:
        str: 转换后的PNG文件路径
        
    异常:
        RuntimeError: 转换失败时抛出
    """
    try:
        # 导入必要的库
        from pillow_heif import register_heif_opener
        
        logger.info(f"开始转换HEIC/HEIF文件: {os.path.basename(heic_path)}")
        
        # 注册HEIF格式支持
        register_heif_opener()
        
        # 确定输出路径
        if output_path:
            # 使用指定的完整路径
            png_path = output_path
        else:
            # 使用旧逻辑生成路径（兼容性）
            _, ext = os.path.splitext(heic_path)
            description = "fromHeif" if ext.lower() == '.heif' else "fromHeic"
            png_path = generate_output_path(
                heic_path,
                output_dir=output_dir,
                section="",
                add_timestamp=True,
                description=description,
                file_type="png"
            )
        
        # 确保输出目录存在
        os.makedirs(os.path.dirname(png_path), exist_ok=True)
        
        # 打开并转换图片
        with Image.open(heic_path) as img:
            # 记录原始图片信息
            logger.debug(f"原始图片模式: {img.mode}, 尺寸: {img.size}")
            
            # 如果是RGBA模式，转为RGB（PNG支持透明度，但统一转为RGB）
            if img.mode == 'RGBA':
                # 创建白色背景
                background = Image.new('RGB', img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[3])  # 使用alpha通道作为mask
                img = background
                logger.debug("已将RGBA模式转换为RGB（白色背景）")
            elif img.mode != 'RGB':
                # 其他模式也转为RGB
                img = img.convert('RGB')
                logger.debug(f"已将{img.mode}模式转换为RGB")
            
            # 保存为PNG格式
            img.save(png_path, 'PNG', optimize=True)
        
        logger.info(f"成功转换为PNG: {os.path.basename(png_path)}")
        return png_path
        
    except ImportError as e:
        logger.error(f"缺少必要的库: {e}")
        raise RuntimeError(f"HEIC转换失败：缺少pillow-heif或Pillow库。请确保已安装相关依赖。")
    except Exception as e:
        logger.error(f"HEIC转PNG失败: {e}", exc_info=True)
        raise RuntimeError(f"HEIC转PNG失败: {e}")


def bmp_to_png(
    bmp_path: str,
    output_path: str = None,
    output_dir: Optional[str] = None
) -> str:
    """
    将BMP格式转换为PNG格式
    
    BMP是Windows位图格式，某些转换工具（如img2pdf）不直接支持，
    需要转换为通用格式。
    
    参数:
        bmp_path: BMP文件路径
        output_path: 输出文件的完整路径（优先使用）
        output_dir: 输出目录（可选）。如果output_path未指定，则使用此参数生成路径
        
    返回:
        str: 转换后的PNG文件路径
        
    异常:
        RuntimeError: 转换失败时抛出
    """
    try:
        logger.info(f"开始转换BMP文件: {os.path.basename(bmp_path)}")
        
        # 确定输出路径
        if output_path:
            # 使用指定的完整路径
            png_path = output_path
        else:
            # 使用旧逻辑生成路径（兼容性）
            png_path = generate_output_path(
                bmp_path,
                output_dir=output_dir,
                section="",
                add_timestamp=True,
                description="fromBmp",
                file_type="png"
            )
        
        # 确保输出目录存在
        os.makedirs(os.path.dirname(png_path), exist_ok=True)
        
        # 打开并转换图片
        with Image.open(bmp_path) as img:
            # 记录原始图片信息
            logger.debug(f"原始图片模式: {img.mode}, 尺寸: {img.size}")
            
            # 如果是RGBA模式，转为RGB
            if img.mode == 'RGBA':
                background = Image.new('RGB', img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[3])
                img = background
                logger.debug("已将RGBA模式转换为RGB（白色背景）")
            elif img.mode != 'RGB':
                img = img.convert('RGB')
                logger.debug(f"已将{img.mode}模式转换为RGB")
            
            # 保存为PNG格式
            img.save(png_path, 'PNG', optimize=True)
        
        logger.info(f"成功转换为PNG: {os.path.basename(png_path)}")
        return png_path
        
    except Exception as e:
        logger.error(f"BMP转PNG失败: {e}", exc_info=True)
        raise RuntimeError(f"BMP转PNG失败: {e}")


# 模块测试
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(name)s:%(lineno)d - %(message)s"
    )
    
    logger.info("图片格式转换模块测试开始")
    
    # 测试文件路径（需要实际存在的文件）
    test_image = "test.png"
    
    if os.path.exists(test_image):
        try:
            # 测试1: 最高质量模式
            logger.info("=" * 50)
            logger.info("测试1: 最高质量模式 PNG->JPEG")
            output1 = generate_output_path(
                test_image,
                section="",
                add_timestamp=True,
                description="test1",
                file_type="jpeg"
            )
            convert_image(
                test_image,
                'jpeg',
                output1,
                {'compress_mode': 'lossless'}
            )
            logger.info(f"测试成功! 输出: {output1}")
            
            # 测试2: 压缩模式
            logger.info("=" * 50)
            logger.info("测试2: 压缩模式 PNG->JPEG (限制200KB)")
            output2 = generate_output_path(
                test_image,
                section="",
                add_timestamp=True,
                description="test2",
                file_type="jpeg"
            )
            convert_image(
                test_image,
                'jpeg',
                output2,
                {
                    'compress_mode': 'limit_size',
                    'size_limit': 200,
                    'size_unit': 'KB'
                }
            )
            logger.info(f"测试成功! 输出: {output2}")
            
        except Exception as e:
            logger.error(f"测试失败: {str(e)}")
    else:
        logger.warning(f"测试文件不存在: {test_image}")
    
    logger.info("图片格式转换模块测试结束")
