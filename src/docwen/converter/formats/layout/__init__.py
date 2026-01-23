"""
版式文件格式转换模块

支持的转换：
- OFD → PDF（预处理）
- XPS → PDF（预处理）
- CAJ → PDF（预处理，待实现）
- PDF → DOCX（外部软件，三级容错）

模块:
- preprocess: OFD/XPS/CAJ → PDF 预处理
- external: PDF → DOCX 转换（Word/LibreOffice/pdf2docx 三级容错）
"""

from .preprocess import (
    ofd_to_pdf,
    xps_to_pdf,
    caj_to_pdf,
)

from .external import (
    pdf_to_docx,
)

__all__ = [
    # 预处理
    'ofd_to_pdf',
    'xps_to_pdf',
    'caj_to_pdf',
    # PDF → DOCX
    'pdf_to_docx',
]
