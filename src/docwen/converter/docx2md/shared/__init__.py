"""
共享处理器模块

包含所有文种通用的处理器：
- 表格处理（table_processor）
- 图片处理（image_processor）
- 公式处理（formula_processor）
- 脚注/尾注处理（note_processor）
- 文本框处理（textbox_processor）
- 内容注入（content_injector）
- 列表处理（list_processor）
- 边框组/分隔符处理（break_processor）
- 样式检测（style_detector）
- Markdown转换（markdown_converter）
- 转换上下文（conversion_context）
"""

# 表格处理
from .table_processor import (
    extract_tables_from_document,
    extract_table_content,
    convert_table_to_md,
    extract_tables_with_structure,
    convert_table_to_md_with_images,
    replace_image_markers_in_table,
)

# 图片处理
from .image_processor import (
    extract_images_from_docx,
    process_image_with_ocr,
    get_paragraph_images,
)

# 公式处理
from .formula_processor import (
    has_formulas_in_paragraph,
    process_paragraph_with_formulas,
    is_formula_supported,
)

# 脚注/尾注处理
from .note_processor import (
    NoteExtractor,
    ENDNOTE_PREFIX,
)

# 文本框处理
from .textbox_processor import (
    extract_textboxes_from_document,
)

# 内容注入
from .content_injector import (
    process_document_with_special_content,
)

# 列表处理
from .list_processor import (
    detect_list_item,
    get_list_marker,
    preprocess_list_ranges,
    ListCounterManager,
    ListContextManager,
    LIST_PARAGRAPH_STYLE_NAMES,
)

# 边框组/分隔符处理
from .break_processor import (
    extract_paragraph_border_info,
    detect_horizontal_rule_in_paragraph,
    detect_page_or_section_break,
    detect_page_break_in_run,
    detect_section_break_in_paragraph,
    detect_all_breaks_in_paragraph,
    BorderGroupTracker,
    is_valid_border,
)

# 样式检测
from .style_detector import (
    detect_paragraph_style_type,
    detect_run_style_type,
    is_full_paragraph_code_style,
    merge_consecutive_runs,
)

# Markdown转换
from .markdown_converter import (
    convert_paragraph_to_markdown,
    convert_paragraph_to_markdown_skip_prefix,
    convert_paragraph_to_markdown_with_styles,
    apply_format_markers,
    has_gray_shading,
    has_paragraph_gray_shading,
    detect_note_reference_in_run,
)

# 转换上下文
from .conversion_context import (
    ConversionContext,
)

__all__ = [
    # 表格处理
    'extract_tables_from_document',
    'extract_table_content',
    'convert_table_to_md',
    'extract_tables_with_structure',
    'convert_table_to_md_with_images',
    'replace_image_markers_in_table',
    # 图片处理
    'extract_images_from_docx',
    'process_image_with_ocr',
    'get_paragraph_images',
    # 公式处理
    'has_formulas_in_paragraph',
    'process_paragraph_with_formulas',
    'is_formula_supported',
    # 脚注/尾注处理
    'NoteExtractor',
    'ENDNOTE_PREFIX',
    # 文本框处理
    'extract_textboxes_from_document',
    # 内容注入
    'process_document_with_special_content',
    # 列表处理
    'detect_list_item',
    'get_list_marker',
    'preprocess_list_ranges',
    'ListCounterManager',
    'ListContextManager',
    'LIST_PARAGRAPH_STYLE_NAMES',
    # 边框组/分隔符处理
    'extract_paragraph_border_info',
    'detect_horizontal_rule_in_paragraph',
    'detect_page_or_section_break',
    'detect_page_break_in_run',
    'detect_section_break_in_paragraph',
    'detect_all_breaks_in_paragraph',
    'BorderGroupTracker',
    'is_valid_border',
    # 样式检测
    'detect_paragraph_style_type',
    'detect_run_style_type',
    'is_full_paragraph_code_style',
    'merge_consecutive_runs',
    # Markdown转换
    'convert_paragraph_to_markdown',
    'convert_paragraph_to_markdown_skip_prefix',
    'convert_paragraph_to_markdown_with_styles',
    'apply_format_markers',
    'has_gray_shading',
    'has_paragraph_gray_shading',
    'detect_note_reference_in_run',
    # 转换上下文
    'ConversionContext',
]
