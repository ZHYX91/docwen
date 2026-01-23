"""
表格汇总包

提供表格汇总功能，支持将多个表格文件合并到一个基准表格中。

功能特性：
- 支持三种汇总模式：按行汇总、按列汇总、按单元格汇总
- 使用滑块算法自动对齐表格（处理行列偏移问题）
- 支持多种表格格式（xlsx、xls、csv等）

使用方式：
    from docwen.table_merger import TableMerger
    
    merger = TableMerger()
    success, message, output_path = merger.merge_tables(
        base_file="base.xlsx",
        collect_files=["file1.xlsx", "file2.xlsx"],
        mode=TableMerger.MODE_BY_ROW
    )
"""

from .core import TableMerger

__all__ = ['TableMerger']

# 包初始化日志
import logging
logger = logging.getLogger(__name__)
logger.info("table_merger包初始化完成 - 提供表格汇总功能")
