"""
md2docx 核心处理器模块

包含主流程相关的处理器：
- docx_processor: DOCX文档处理
- md_processor: Markdown解析
- placeholder_handler: 占位符替换
- xml_processor: XML级别操作
"""

# docx_processor - DOCX文档处理
from .docx_processor import (
    load_document,
    replace_placeholders,
    process_main_content,
    remove_marked_elements,
    save_and_process_temp_file,
)

# md_processor - Markdown解析
from .md_processor import (
    process_md_body,
    process_md_body_with_notes,
)

# placeholder_handler - 占位符替换
from .placeholder_handler import (
    mark_special_placeholders,
    is_special_marked,
    remove_special_mark,
    process_paragraph_placeholders,
    process_table_cell_placeholders,
    process_attachment_description_placeholder,
    process_image_placeholders,
    try_remove_element,
    extract_placeholder_keys,
    SPECIAL_PLACEHOLDERS,
)

# xml_processor - XML级别操作
from .xml_processor import (
    process_docx_file,
)

__all__ = [
    # docx_processor
    'load_document',
    'replace_placeholders',
    'process_main_content',
    'remove_marked_elements',
    'save_and_process_temp_file',
    # md_processor
    'process_md_body',
    'process_md_body_with_notes',
    # placeholder_handler
    'mark_special_placeholders',
    'is_special_marked',
    'remove_special_mark',
    'process_paragraph_placeholders',
    'process_table_cell_placeholders',
    'process_attachment_description_placeholder',
    'process_image_placeholders',
    'try_remove_element',
    'extract_placeholder_keys',
    'SPECIAL_PLACEHOLDERS',
    # xml_processor
    'process_docx_file',
]
