"""
MD转DOCX子包入口

模块结构:
├── core.py: 转换核心逻辑（入口）
├── processors/: 核心处理器
│   ├── docx_processor.py: Word文档处理（主处理流程）
│   ├── md_processor.py: Markdown解析和处理
│   ├── placeholder_handler.py: 占位符替换
│   └── xml_processor.py: XML深度处理
├── handlers/: 元素处理器
│   ├── break_handler.py: 分页/分节/分隔线处理
│   ├── formula_handler.py: 公式处理
│   ├── list_handler.py: 列表处理
│   ├── note_handler.py: 脚注/尾注处理
│   ├── notes_part_handler.py: 脚注/尾注Part处理
│   ├── numbering_handler.py: 编号部件处理
│   ├── table_handler.py: 表格创建与样式处理
│   └── text_handler.py: 文本格式化处理
└── style/: 样式系统
    ├── injector.py: 样式注入逻辑
    ├── templates.py: 样式XML模板定义
    └── helper.py: 样式查找辅助函数
"""

# 脚注/尾注处理（从 handlers 模块导入）
# 分页/分节/分隔线处理（从 handlers 模块导入）
from .handlers.break_handler import (
    append_horizontal_rule_to_paragraph,
    append_page_break_to_paragraph,
    append_section_break_to_paragraph,
    insert_horizontal_rule,
    insert_page_break,
    insert_section_break,
)
from .handlers.note_handler import (
    NOTE_REF_REGEX,
    NoteContext,
    extract_notes,
    find_note_references,
    get_clean_endnote_id,
    is_endnote_id,
)
from .handlers.notes_part_handler import (
    check_notes_parts,
    ensure_notes_parts,
)

# 表格处理（从 handlers 模块导入）
from .handlers.table_handler import (
    create_word_table,
    get_table_content_style_name,
    get_table_style_name,
)

# 样式查找辅助函数（从 style 模块导入）
from .style.helper import (
    get_code_block_style_name,
    get_formula_block_style_name,
    get_heading_style_name,
    get_horizontal_rule_style_name,
    get_list_block_style_name,
    get_quote_style_name,
)

__all__ = [
    "NOTE_REF_REGEX",
    "NoteContext",
    "append_horizontal_rule_to_paragraph",
    "append_page_break_to_paragraph",
    "append_section_break_to_paragraph",
    "check_notes_parts",
    "create_word_table",
    "ensure_notes_parts",
    "extract_notes",
    "find_note_references",
    "get_clean_endnote_id",
    "get_code_block_style_name",
    "get_formula_block_style_name",
    "get_heading_style_name",
    "get_horizontal_rule_style_name",
    "get_list_block_style_name",
    "get_quote_style_name",
    "get_table_content_style_name",
    "get_table_style_name",
    "insert_horizontal_rule",
    "insert_page_break",
    "insert_section_break",
    "is_endnote_id",
]
