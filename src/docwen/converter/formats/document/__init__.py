"""
文档格式转换模块

支持的转换：
- DOC/WPS → DOCX（预处理）
- RTF/ODT ↔ DOCX（格式互转）
- DOCX → DOC/ODT/RTF（后处理）

模块:
- external: 使用外部软件（WPS/Office/LibreOffice）进行转换
- builtin: 使用内置方法（python-docx等）进行转换【预留】
"""

from .external import (
    docx_to_doc,
    docx_to_odt,
    docx_to_rtf,
    odt_to_docx,
    office_to_docx,
    rtf_to_docx,
)

__all__ = [
    "docx_to_doc",
    "docx_to_odt",
    "docx_to_rtf",
    "odt_to_docx",
    "office_to_docx",
    "rtf_to_docx",
]
