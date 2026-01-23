"""
图片转换策略子包

提供图片文件转换相关的策略类：
- ImageToPdfStrategy: 图片转PDF
- ImageToMarkdownStrategy: 图片转Markdown（含OCR）
- ImageFormatConversionStrategy: 图片格式互转
- MergeImagesToTiffStrategy: 合并图片为多页TIFF

辅助函数：
- get_image_format_description: 获取格式描述
- is_multipage_tiff: 检测多页TIFF
- extract_tiff_pages: 提取TIFF页面
"""

# 导入辅助函数
from .utils import (
    get_image_format_description,
    is_multipage_tiff,
    extract_tiff_pages,
)

# 导入策略类（自动注册）
from .to_pdf import ImageToPdfStrategy
from .to_markdown import ImageToMarkdownStrategy
from .format_conversion import ImageFormatConversionStrategy
from .merge import MergeImagesToTiffStrategy

__all__ = [
    # 辅助函数
    'get_image_format_description',
    'is_multipage_tiff',
    'extract_tiff_pages',
    # 策略类
    'ImageToPdfStrategy',
    'ImageToMarkdownStrategy',
    'ImageFormatConversionStrategy',
    'MergeImagesToTiffStrategy',
]
