"""
格式转换支持模块

提供多种文档格式的转换支持：
- Office格式：DOC/DOCX/ODT/RTF, XLS/XLSX/ODS（需Office软件）
- 版式文件：CAJ/XPS/OFD → PDF
- 图片格式：HEIC/BMP → PNG
"""

# Office格式支持
from .office import (
    # 表格格式
    office_to_xlsx,
    xlsx_to_ods,
    ods_to_xlsx,
    convert_xlsx_to_xls,
    # 文档格式
    office_to_docx,
    docx_to_odt,
    odt_to_docx,
    rtf_to_docx,
    convert_docx_to_doc,
    convert_docx_to_rtf,
    # 工具函数
    check_office_availability,
    # 异常
    OfficeSoftwareNotFoundError,
)

# 版式文件支持
from .layout import (
    caj_to_pdf,
    xps_to_pdf,
    ofd_to_pdf,
)

# 图片格式支持
from .image import (
    heic_to_png,
    bmp_to_png,
)

__all__ = [
    # Office: 表格
    'office_to_xlsx',
    'xlsx_to_ods',
    'ods_to_xlsx',
    'convert_xlsx_to_xls',
    # Office: 文档
    'office_to_docx',
    'docx_to_odt',
    'odt_to_docx',
    'rtf_to_docx',
    'convert_docx_to_doc',
    'convert_docx_to_rtf',
    # Office: 工具
    'check_office_availability',
    # 版式文件
    'caj_to_pdf',
    'xps_to_pdf',
    'ofd_to_pdf',
    # 图片
    'heic_to_png',
    'bmp_to_png',
    # 异常
    'OfficeSoftwareNotFoundError',
]
