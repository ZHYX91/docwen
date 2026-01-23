"""
格式转换模块

提供各种文件格式之间的转换功能。

子模块：
- common: 公共模块（软件检测、容错框架）
- document: 文档格式转换（DOC/DOCX/RTF/ODT）
- spreadsheet: 表格格式转换（XLS/XLSX/ODS）
- layout: 版式文件转换（PDF/OFD/XPS/CAJ）
- image: 图片格式转换
- pdf_export: PDF 导出

架构说明：
    本模块遵循三层架构设计：
    1. strategies（策略层）→ 协调转换流程
    2. formats（格式转换层）→ 具体格式转换实现（本模块）
    3. docx2md/md2docx 等（核心转换层）→ MD 互转

使用方式:
    # 推荐：从子模块导入
    from docwen.converter.formats.document import office_to_docx
    from docwen.converter.formats.layout import ofd_to_pdf
    
    # 兼容：从本模块导入（保持向后兼容）
    from docwen.converter.formats import office_to_docx, ofd_to_pdf
"""

# ==================== 公共模块 ====================
from .common import (
    check_office_availability,
    OfficeSoftwareNotFoundError,
    convert_with_fallback,
    convert_with_libreoffice,
    try_com_conversion,
    build_com_converters,
)

# ==================== 文档格式转换 ====================
from .document import (
    office_to_docx,
    rtf_to_docx,
    odt_to_docx,
    docx_to_odt,
    docx_to_doc,
    docx_to_rtf,
)

# ==================== 表格格式转换 ====================
from .spreadsheet import (
    office_to_xlsx,
    xlsx_to_ods,
    ods_to_xlsx,
    xlsx_to_xls,
    csv_to_xlsx,
    xlsx_to_csv,
)

# ==================== 版式文件转换 ====================
from .layout import (
    ofd_to_pdf,
    xps_to_pdf,
    caj_to_pdf,
    pdf_to_docx,
)

# ==================== PDF 导出 ====================
from .pdf_export import (
    docx_to_pdf,
    xlsx_to_pdf,
)

# ==================== 图片格式转换 ====================
from .image import (
    convert_image,
    heic_to_png,
    bmp_to_png,
    compress_to_size,
    get_save_params,
)

__all__ = [
    # common
    'check_office_availability',
    'OfficeSoftwareNotFoundError',
    'convert_with_fallback',
    'convert_with_libreoffice',
    'try_com_conversion',
    'build_com_converters',
    # document
    'office_to_docx',
    'rtf_to_docx',
    'odt_to_docx',
    'docx_to_odt',
    'docx_to_doc',
    'docx_to_rtf',
    # spreadsheet
    'office_to_xlsx',
    'xlsx_to_ods',
    'ods_to_xlsx',
    'xlsx_to_xls',
    'csv_to_xlsx',
    'xlsx_to_csv',
    # layout
    'ofd_to_pdf',
    'xps_to_pdf',
    'caj_to_pdf',
    'pdf_to_docx',
    # pdf_export
    'docx_to_pdf',
    'xlsx_to_pdf',
    # image
    'convert_image',
    'heic_to_png',
    'bmp_to_png',
    'compress_to_size',
    'get_save_params',
]
