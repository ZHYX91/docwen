"""
转换器包

提供各种文档格式转换功能：
- Markdown ↔ DOCX/XLSX（公文处理核心）
- PDF 处理和转换（内容提取、格式转换）
- 多格式支持（Office/版式/图片）
- 智能转换链（自动规划转换路径）
"""

# 主要转换子包
from . import docx2md
from . import md2docx
from . import md2xlsx
from . import xlsx2md
from . import pdf2md
from . import formats

# 核心功能
from .md2docx.core import convert as convert_md_to_docx
from .md2xlsx.core import convert as convert_md_to_xlsx
from .smart_converter import SmartConverter, OfficeSoftwareNotFoundError
from .formats.spreadsheet import csv_to_xlsx, xlsx_to_csv

__all__ = [
    # 子包
    'docx2md',
    'md2docx',
    'md2xlsx',
    'xlsx2md',
    'pdf2md',
    'formats',
    # Markdown转换
    'convert_md_to_docx',
    'convert_md_to_xlsx',
    # 智能转换
    'SmartConverter',
    # 表格转换
    'csv_to_xlsx',
    'xlsx_to_csv',
    # 异常
    'OfficeSoftwareNotFoundError',
]
