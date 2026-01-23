"""
图片格式转换模块

支持的功能：
- 通用图片格式转换（PNG/JPEG/WebP/BMP/GIF/TIFF）
- 特殊格式预处理（HEIC/HEIF → PNG, BMP → PNG）
- 图片压缩（有损/无损）

模块:
- core: 通用图片格式转换
- preprocess: 特殊格式预处理
- compression: 压缩相关功能
- external: 外部软件支持【预留】
"""

from .core import convert_image
from .preprocess import heic_to_png, bmp_to_png
from .compression import (
    compress_to_size,
    get_save_params,
    should_compress,
    get_file_size,
    bytes_to_kb,
    bytes_to_mb,
)

__all__ = [
    # 核心转换
    'convert_image',
    # 特殊格式预处理
    'heic_to_png',
    'bmp_to_png',
    # 压缩功能
    'compress_to_size',
    'get_save_params',
    'should_compress',
    'get_file_size',
    'bytes_to_kb',
    'bytes_to_mb',
]
