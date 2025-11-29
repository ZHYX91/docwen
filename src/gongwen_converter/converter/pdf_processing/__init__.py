"""
PDF处理模块

提供PDF文件的多种处理功能：
- PDF内容灵活提取（文本、图片、OCR，支持7种组合）
- PDF转DOCX（支持Word/LibreOffice/pdf2docx三级容错）
- PDF转Markdown（基于pymupdf4llm）
"""

from .pdf_converter_utils import convert_pdf_to_docx, extract_content_from_docx
from .pdf_pymupdf4llm import extract_pdf_with_pymupdf4llm

__all__ = [
    # 核心功能
    'convert_pdf_to_docx',
    'extract_content_from_docx',
    'extract_pdf_with_pymupdf4llm',
]
