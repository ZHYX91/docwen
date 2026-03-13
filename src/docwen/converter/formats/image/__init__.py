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

import importlib

from .compression import (
    bytes_to_kb,
    bytes_to_mb,
    compress_to_size,
    get_file_size,
    get_save_params,
    should_compress,
)


def __getattr__(name: str):
    if name == "convert_image":
        return importlib.import_module(f"{__name__}.core").convert_image
    if name == "bmp_to_png":
        return importlib.import_module(f"{__name__}.preprocess").bmp_to_png
    if name == "heic_to_png":
        return importlib.import_module(f"{__name__}.preprocess").heic_to_png
    raise AttributeError(name)


__all__ = [
    "bmp_to_png",
    "bytes_to_kb",
    "bytes_to_mb",
    # 压缩功能
    "compress_to_size",
    # 核心转换
    "convert_image",
    "get_file_size",
    "get_save_params",
    # 特殊格式预处理
    "heic_to_png",
    "should_compress",
]
