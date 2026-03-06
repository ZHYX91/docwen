"""
Markdown 策略子包

提供 Markdown 文件转换为其他格式的策略：
- to_document: MD → DOCX/DOC/ODT/RTF
- to_spreadsheet: MD → XLSX/XLS/ODS/CSV

使用方式：
    from docwen.services.strategies import get_strategy

    strategy_class = get_strategy(source_format='md', target_format='docx')
"""

from .to_spreadsheet import (
    BaseMdToSpreadsheetStrategy,
    MdToCsvStrategy,
    MdToOdsStrategy,
    MdToXlsStrategy,
    MdToXlsxStrategy,
)

try:
    from .to_document import (
        BaseMdToDocumentStrategy,
        MdToDocStrategy,
        MdToDocxStrategy,
        MdToOdtStrategy,
        MdToRtfStrategy,
    )
except Exception:
    BaseMdToDocumentStrategy = None
    MdToDocStrategy = None
    MdToDocxStrategy = None
    MdToOdtStrategy = None
    MdToRtfStrategy = None

__all__ = [
    # 文档转换策略
    "BaseMdToDocumentStrategy",
    # 表格转换策略
    "BaseMdToSpreadsheetStrategy",
    "MdToCsvStrategy",
    "MdToDocStrategy",
    "MdToDocxStrategy",
    "MdToOdsStrategy",
    "MdToOdtStrategy",
    "MdToRtfStrategy",
    "MdToXlsStrategy",
    "MdToXlsxStrategy",
]
