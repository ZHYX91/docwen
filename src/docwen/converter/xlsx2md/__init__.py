"""
XLSX转MD子包
表格文件到Markdown的转换功能

支持格式：xlsx, csv（标准格式）
旧格式（xls, et）应在策略层预处理
"""

from .core import convert_spreadsheet_to_md

__all__ = ["convert_spreadsheet_to_md"]
