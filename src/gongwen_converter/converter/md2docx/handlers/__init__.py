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
    insert_page_break,
    insert_section_break,
    insert_horizontal_rule,
    append_page_break_to_paragraph,
    append_section_break_to_paragraph,
    append_horizontal_rule_to_paragraph,
)

# formula_handler - 公式处理
from .formula_handler import (
    process_paragraph_formulas,
    is_formula_supported,
    has_latex_formulas,
)

# list_handler - 列表处理
from .list_handler import (
    ListFormatManager,
    apply_list_to_paragraph,
    analyze_list_structure,
    group_consecutive_list_items,
    BASE_INDENT,
    INDENT_INCREMENT,
)

# note_handler - 脚注/尾注逻辑
from .note_handler import (
    NoteContext,
    extract_notes,
    find_note_references,
    is_endnote_id,
    get_clean_endnote_id,
    write_notes_to_docx,
    process_text_with_note_references,
    NOTE_REF_REGEX,
    ENDNOTE_PREFIX,
)

# notes_part_handler - notes.xml部件处理
from .notes_part_handler import (
    ensure_notes_parts,
    check_notes_parts,
)

# numbering_handler - 编号部件处理
from .numbering_handler import (
    ensure_numbering_part,
)

# table_handler - 表格处理
from .table_handler import (
    create_word_table,
    get_or_inject_table_style,
    inject_table_style,
    inject_table_content_style,
)

# text_handler - 文本格式化处理
from .text_handler import (
    parse_markdown_formatting,
    apply_formats_to_run,
    add_formatted_text_to_paragraph,
    add_formatted_text_to_heading,
    add_formatted_text_to_paragraph_with_breaks,
)

__all__ = [
    # break_handler
    'insert_page_break',
    'insert_section_break',
    'insert_horizontal_rule',
    'append_page_break_to_paragraph',
    'append_section_break_to_paragraph',
    'append_horizontal_rule_to_paragraph',
    # formula_handler
    'process_paragraph_formulas',
    'is_formula_supported',
    'has_latex_formulas',
    # list_handler
    'ListFormatManager',
    'apply_list_to_paragraph',
    'analyze_list_structure',
    'group_consecutive_list_items',
    'BASE_INDENT',
    'INDENT_INCREMENT',
    # note_handler
    'NoteContext',
    'extract_notes',
    'find_note_references',
    'is_endnote_id',
    'get_clean_endnote_id',
    'write_notes_to_docx',
    'process_text_with_note_references',
    'NOTE_REF_REGEX',
    'ENDNOTE_PREFIX',
    # notes_part_handler
    'ensure_notes_parts',
    'check_notes_parts',
    # numbering_handler
    'ensure_numbering_part',
    # table_handler
    'create_word_table',
    'get_or_inject_table_style',
    'inject_table_style',
    'inject_table_content_style',
    # text_handler
    'parse_markdown_formatting',
    'apply_formats_to_run',
    'add_formatted_text_to_paragraph',
    'add_formatted_text_to_heading',
    'add_formatted_text_to_paragraph_with_breaks',
]
