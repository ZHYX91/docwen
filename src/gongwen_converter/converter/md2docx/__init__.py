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
from .handlers.note_handler import (
    extract_notes,
    find_note_references,
    is_endnote_id,
    get_clean_endnote_id,
    NOTE_REF_REGEX,
    NoteContext,
)

from .handlers.notes_part_handler import (
    ensure_notes_parts,
    check_notes_parts,
)

# 样式查找辅助函数（从 style 模块导入）
from .style.helper import (
    get_heading_style_name,
    get_code_block_style_name,
    get_quote_style_name,
    get_formula_block_style_name,
    get_list_block_style_name,
    get_horizontal_rule_style_name,
)

# 表格处理（从 handlers 模块导入）
from .handlers.table_handler import (
    create_word_table,
    get_or_inject_table_style,
    inject_table_style,
    inject_table_content_style,
)

# 分页/分节/分隔线处理（从 handlers 模块导入）
from .handlers.break_handler import (
    insert_page_break,
    insert_section_break,
    insert_horizontal_rule,
    append_page_break_to_paragraph,
    append_section_break_to_paragraph,
    append_horizontal_rule_to_paragraph,
)
