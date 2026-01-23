"""
操作类策略子包

与转换类策略不同，操作类策略执行的是文件处理操作而非格式转换：
- merge_tables: 表格汇总（多个表格合并）
- md_numbering: Markdown序号处理（添加/清理序号）

这些策略使用 @register_action 注册，通过 get_strategy(action_type='xxx') 查找。
"""

# 导入策略类（自动注册）
from .merge_tables import MergeTablesStrategy
from .md_numbering import MdNumberingStrategy

__all__ = [
    'MergeTablesStrategy',
    'MdNumberingStrategy',
]
