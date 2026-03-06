"""
表格格式转换模块

支持的转换：
- XLS/ET → XLSX（预处理）
- ODS ↔ XLSX（格式互转）
- XLSX → XLS/ODS（后处理）
- CSV ↔ XLSX（内置方法）

模块:
- external: 使用外部软件（WPS/Excel/LibreOffice）进行转换
- csv_convert: CSV与XLSX互转（使用pandas/openpyxl）
- builtin: 使用内置方法（openpyxl等）进行转换【预留】
"""

from .csv_convert import (
    csv_to_xlsx,
    xlsx_to_csv,
)
from .external import (
    ods_to_xlsx,
    office_to_xlsx,
    xlsx_to_ods,
    xlsx_to_xls,
)

__all__ = [
    # CSV转换（内置方法）
    "csv_to_xlsx",
    "ods_to_xlsx",
    # 外部软件转换
    "office_to_xlsx",
    "xlsx_to_csv",
    "xlsx_to_ods",
    "xlsx_to_xls",
]
