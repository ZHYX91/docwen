"""
md2docx 元素处理器模块

包含各种内容元素的处理器：
- break_handler: 分页/分节/分隔线处理
- formula_handler: 公式处理
- list_handler: 列表处理
- note_handler: 脚注/尾注逻辑
- notes_part_handler: notes.xml部件处理
- numbering_handler: 编号部件处理
- table_handler: 表格处理
- text_handler: 文本格式化处理
"""

# break_handler - 分页/分节/分隔线处理
from .break_handler import (
    append_horizontal_rule_to_paragraph,
    append_page_break_to_paragraph,
    append_section_break_to_paragraph,
    insert_horizontal_rule,
    insert_page_break,
    insert_section_break,
)

# formula_handler - 公式处理
from .formula_handler import (
    has_latex_formulas,
    is_formula_supported,
    process_paragraph_formulas,
)

# list_handler - 列表处理
from .list_handler import (
    BASE_INDENT,
    INDENT_INCREMENT,
    ListFormatManager,
    analyze_list_structure,
    apply_list_to_paragraph,
    group_consecutive_list_items,
)

# note_handler - 脚注/尾注逻辑
from .note_handler import (
    ENDNOTE_PREFIX,
    NOTE_REF_REGEX,
    NoteContext,
    extract_notes,
    find_note_references,
    get_clean_endnote_id,
    is_endnote_id,
    process_text_with_note_references,
    write_notes_to_docx,
)

# notes_part_handler - notes.xml部件处理
from .notes_part_handler import (
    check_notes_parts,
    ensure_notes_parts,
)

# numbering_handler - 编号部件处理
from .numbering_handler import (
    ensure_numbering_part,
)

# table_handler - 表格处理
from .table_handler import (
    create_word_table,
    get_table_content_style_name,
    get_table_style_name,
)

# text_handler - 文本格式化处理
from .text_handler import (
    add_formatted_text_to_heading,
    add_formatted_text_to_paragraph,
    apply_formats_to_run,
    parse_markdown_formatting,
)

__all__ = [
    "BASE_INDENT",
    "ENDNOTE_PREFIX",
    "INDENT_INCREMENT",
    "NOTE_REF_REGEX",
    # list_handler
    "ListFormatManager",
    # note_handler
    "NoteContext",
    "add_formatted_text_to_heading",
    "add_formatted_text_to_paragraph",
    "analyze_list_structure",
    "append_horizontal_rule_to_paragraph",
    "append_page_break_to_paragraph",
    "append_section_break_to_paragraph",
    "apply_formats_to_run",
    "apply_list_to_paragraph",
    "check_notes_parts",
    # table_handler
    "create_word_table",
    # notes_part_handler
    "ensure_notes_parts",
    # numbering_handler
    "ensure_numbering_part",
    "extract_notes",
    "find_note_references",
    "get_clean_endnote_id",
    "get_table_content_style_name",
    "get_table_style_name",
    "group_consecutive_list_items",
    "has_latex_formulas",
    "insert_horizontal_rule",
    # break_handler
    "insert_page_break",
    "insert_section_break",
    "is_endnote_id",
    "is_formula_supported",
    # text_handler
    "parse_markdown_formatting",
    # formula_handler
    "process_paragraph_formulas",
    "process_text_with_note_references",
    "write_notes_to_docx",
]
