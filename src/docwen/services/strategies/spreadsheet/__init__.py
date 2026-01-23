"""
表格转换策略子包

提供表格文件转换相关的策略类：
- SpreadsheetToMarkdownStrategy: 表格转Markdown
- SpreadsheetToTxtStrategy: 表格转TXT
- SpreadsheetToPdfStrategy: 表格转PDF
- CsvToXlsxStrategy: CSV转XLSX
- XlsxToCsvStrategy: XLSX转CSV
- 动态格式转换策略（通过工厂批量注册）

辅助函数：
- _preprocess_table_file: 表格文件预处理
"""

# 导入辅助函数
from .utils import _preprocess_table_file

# 导入策略类（自动注册）
from .to_markdown import SpreadsheetToMarkdownStrategy, SpreadsheetToTxtStrategy
from .to_pdf import SpreadsheetToPdfStrategy
from .csv_xlsx import CsvToXlsxStrategy, XlsxToCsvStrategy

# 导入格式转换模块（触发动态策略批量注册）
from . import format_conversion

__all__ = [
    # 辅助函数
    '_preprocess_table_file',
    # 策略类
    'SpreadsheetToMarkdownStrategy',
    'SpreadsheetToTxtStrategy',
    'SpreadsheetToPdfStrategy',
    'CsvToXlsxStrategy',
    'XlsxToCsvStrategy',
]
