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
    process_main_content,
    remove_marked_elements,
    replace_placeholders,
    save_and_process_temp_file,
)

# md_processor - Markdown解析
from .md_processor import (
    process_md_body,
    process_md_body_with_notes,
)

# placeholder_handler - 占位符替换
from .placeholder_handler import (
    extract_placeholder_keys,
    is_special_marked,
    mark_special_placeholders,
    process_image_placeholders,
    process_paragraph_placeholders,
    process_table_cell_placeholders,
    remove_special_mark,
    try_remove_element,
)

# xml_processor - XML级别操作
from .xml_processor import (
    process_docx_file,
)

__all__ = [
    "extract_placeholder_keys",
    "is_special_marked",
    # docx_processor
    "load_document",
    # placeholder_handler
    "mark_special_placeholders",
    # xml_processor
    "process_docx_file",
    "process_image_placeholders",
    "process_main_content",
    # md_processor
    "process_md_body",
    "process_md_body_with_notes",
    "process_paragraph_placeholders",
    "process_table_cell_placeholders",
    "remove_marked_elements",
    "remove_special_mark",
    "replace_placeholders",
    "save_and_process_temp_file",
    "try_remove_element",
]
