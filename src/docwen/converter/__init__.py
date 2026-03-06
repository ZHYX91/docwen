"""
转换器包

提供各种文档格式转换功能：
- Markdown ↔ DOCX/XLSX（公文处理核心）
- PDF 处理和转换（内容提取、格式转换）
- 多格式支持（Office/版式/图片）
- 智能转换链（自动规划转换路径）
"""

import importlib

from .smart_converter import OfficeSoftwareNotFoundError, SmartConverter


def __getattr__(name: str):
    if name in {"docx2md", "formats", "md2docx", "md2xlsx", "pdf2md", "xlsx2md"}:
        return importlib.import_module(f"{__name__}.{name}")
    if name == "csv_to_xlsx":
        from .formats.spreadsheet import csv_to_xlsx

        return csv_to_xlsx
    if name == "xlsx_to_csv":
        from .formats.spreadsheet import xlsx_to_csv

        return xlsx_to_csv
    if name == "convert_md_to_docx":
        from .md2docx.core import convert as convert_md_to_docx

        return convert_md_to_docx
    if name == "convert_md_to_xlsx":
        from .md2xlsx.core import convert as convert_md_to_xlsx

        return convert_md_to_xlsx
    raise AttributeError(name)

__all__ = [
    # 异常
    "OfficeSoftwareNotFoundError",
    # 智能转换
    "SmartConverter",
    # Markdown转换
    "convert_md_to_docx",
    "convert_md_to_xlsx",
    # 表格转换
    "csv_to_xlsx",
    # 子包
    "docx2md",
    "formats",
    "md2docx",
    "md2xlsx",
    "pdf2md",
    "xlsx2md",
    "xlsx_to_csv",
]
