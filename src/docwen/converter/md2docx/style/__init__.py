"""
md2docx 样式系统模块

包含样式相关的处理器：
- injector: 样式注入逻辑
- templates: 样式XML模板定义
- helper: 样式查找辅助函数
"""

# injector - 样式注入
# helper - 样式查找
from .helper import (
    get_code_block_style_name,
    get_formula_block_style_name,
    get_heading_style_name,
    get_horizontal_rule_style_name,
    get_inline_code_style_name,
    get_list_block_style_name,
    get_quote_style_name,
)
from .injector import (
    ensure_styles,
)

# templates - 样式模板
from .templates import (
    CODE_BLOCK_STYLE_TEMPLATE,
    FORMULA_BLOCK_STYLE_TEMPLATE,
    HEADING_STYLE_TEMPLATES,
    INLINE_CODE_STYLE_TEMPLATE,
    INLINE_FORMULA_STYLE_TEMPLATE,
    QUOTE_STYLE_TEMPLATES,
    TABLE_CONTENT_STYLE_TEMPLATE,
    THREE_LINE_TABLE_STYLE_TEMPLATE,
)

__all__ = [
    "CODE_BLOCK_STYLE_TEMPLATE",
    "FORMULA_BLOCK_STYLE_TEMPLATE",
    # templates
    "HEADING_STYLE_TEMPLATES",
    "INLINE_CODE_STYLE_TEMPLATE",
    "INLINE_FORMULA_STYLE_TEMPLATE",
    "QUOTE_STYLE_TEMPLATES",
    "TABLE_CONTENT_STYLE_TEMPLATE",
    "THREE_LINE_TABLE_STYLE_TEMPLATE",
    # injector
    "ensure_styles",
    "get_code_block_style_name",
    "get_formula_block_style_name",
    # helper
    "get_heading_style_name",
    "get_horizontal_rule_style_name",
    "get_inline_code_style_name",
    "get_list_block_style_name",
    "get_quote_style_name",
]
