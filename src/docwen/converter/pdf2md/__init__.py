"""
PDF转Markdown模块

提供PDF文件的多种处理功能：
- PDF转Markdown（基于pymupdf4llm，支持4种组合）
- 从DOCX提取内容（用于PDF灵活提取流程）

注意：PDF转DOCX功能已移至 formats.layout 模块

模块结构：
- core: PDF→MD 核心转换（pymupdf4llm实现）
- docx_extractor: 从DOCX提取内容（文本、图片、OCR）

使用方式：
    from docwen.converter.pdf2md import extract_pdf_with_pymupdf4llm

    result = extract_pdf_with_pymupdf4llm(
        pdf_path="input.pdf",
        extract_images=True,
        extract_ocr=False,
        output_dir="output/"
    )
"""

from .core import extract_pdf_with_pymupdf4llm
from .docx_extractor import extract_content_from_docx

__all__ = [
    "extract_content_from_docx",
    # 核心功能
    "extract_pdf_with_pymupdf4llm",
]
