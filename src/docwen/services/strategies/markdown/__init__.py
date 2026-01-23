"""
Markdown 策略子包

提供 Markdown 文件转换为其他格式的策略：
- to_document: MD → DOCX/DOC/ODT/RTF
- to_spreadsheet: MD → XLSX/XLS/ODS/CSV

使用方式：
    from docwen.services.strategies import get_strategy
    
    strategy_class = get_strategy(source_format='md', target_format='docx')
"""

# 导入子模块以触发策略注册
from . import to_document
from . import to_spreadsheet

# 导出所有策略类（便于直接导入）
from .to_document import (
    BaseMdToDocumentStrategy,
    MdToDocxStrategy,
    MdToDocStrategy,
    MdToOdtStrategy,
    MdToRtfStrategy,
)
from .to_spreadsheet import (
    BaseMdToSpreadsheetStrategy,
    MdToXlsxStrategy,
    MdToXlsStrategy,
    MdToOdsStrategy,
    MdToCsvStrategy,
)

__all__ = [
    # 文档转换策略
    'BaseMdToDocumentStrategy',
    'MdToDocxStrategy',
    'MdToDocStrategy',
    'MdToOdtStrategy',
    'MdToRtfStrategy',
    # 表格转换策略
    'BaseMdToSpreadsheetStrategy',
    'MdToXlsxStrategy',
    'MdToXlsStrategy',
    'MdToOdsStrategy',
    'MdToCsvStrategy',
]
