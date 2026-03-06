"""
简化模式转换器模块

实现普通Word文档的基础转换逻辑。

特性：
- 基于Word样式（Title、Subtitle、Heading 1-6）转换
- 提取Title/Subtitle到YAML元数据
- 不做公文元素识别
- 支持图片提取和OCR
- 支持文本框和表格处理

模块结构：
- converter.py: 主转换入口
- yaml_builder.py: YAML头部生成
- paragraph_handler.py: 段落处理器

使用示例：
    from docwen.converter.docx2md.simple import convert_docx_to_md_simple

    result = convert_docx_to_md_simple('document.docx')
    if result['success']:
        print(result['main_content'])
"""

from .converter import convert_docx_to_md_simple
from .paragraph_handler import ParagraphHandler
from .yaml_builder import build_yaml_header, build_yaml_header_string

__all__ = [
    # 段落处理器
    "ParagraphHandler",
    # YAML生成
    "build_yaml_header",
    "build_yaml_header_string",
    # 主转换函数
    "convert_docx_to_md_simple",
]
