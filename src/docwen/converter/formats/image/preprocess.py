"""
特殊图片格式预处理模块

将特殊图片格式（HEIC/HEIF/BMP等）转换为通用格式（PNG）：
- HEIC/HEIF → PNG（苹果设备图片格式）
- BMP → PNG（某些工具不直接支持BMP）
"""

import os
import logging
from typing import Optional
from PIL import Image
from docwen.utils.path_utils import generate_output_path

logger = logging.getLogger(__name__)


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
