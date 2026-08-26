"""Core Markdown link parsing, resolution, embedding, and rendering utilities.

The package is independent of plugin and runtime composition layers.
"""

from docwen_core.links._anchor import (
    extract_block_by_id,
    extract_section_by_heading,
    parse_anchor,
    split_yaml_front_matter_source,
    strip_yaml_front_matter,
)
from docwen_core.links._data_uri import (
    is_data_uri_image,
    resolve_data_uri_image_to_temp_file,
)
from docwen_core.links._embed_dispatch import (
    process_single_embed,
    resolve_embedded_links,
)
from docwen_core.links._embed_image import (
    EmbeddedImageMode,
    format_image_placeholder,
    process_embedded_image,
    split_alt_text_and_size,
)
from docwen_core.links._embed_md import (
    TABLE_CELL_BR_TOKEN,
    EmbeddedMdMode,
    make_table_safe,
    process_embedded_md_file,
    restore_table_safe_breaks,
)
from docwen_core.links._error_semantics import (
    LinkErrorKind,
    NotFoundAction,
    dispatch_error_output,
    make_error_placeholder,
    make_keep_link,
)
from docwen_core.links._markdown_inline import escape_unescaped_pipes
from docwen_core.links._markdown_orchestrator import (
    process_markdown_links,
)
from docwen_core.links._non_embed import (
    _process_non_embed_links,
    split_markdown_block_segments,
    split_markdown_inline_segments,
)
from docwen_core.links._patterns import (
    IMAGE_EXTENSIONS,
    MD_EXTENSIONS,
    WIKI_EMBED_PATTERN,
    WIKI_EMBED_SIZE_PATTERN,
)
from docwen_core.links._resolver import (
    get_file_type,
    normalize_link_target,
    resolve_file_path,
)
from docwen_core.links.declared_resources import (
    DeclaredResourceError,
    DeclaredResourceResolver,
    bind_declared_markdown_images,
    reject_declared_input_link_lookups,
)

__all__ = [
    "IMAGE_EXTENSIONS",
    "MD_EXTENSIONS",
    "TABLE_CELL_BR_TOKEN",
    "WIKI_EMBED_PATTERN",
    "WIKI_EMBED_SIZE_PATTERN",
    "DeclaredResourceError",
    "DeclaredResourceResolver",
    "EmbeddedImageMode",
    "EmbeddedMdMode",
    "LinkErrorKind",
    "NotFoundAction",
    "_process_non_embed_links",
    "bind_declared_markdown_images",
    "dispatch_error_output",
    "escape_unescaped_pipes",
    "extract_block_by_id",
    "extract_section_by_heading",
    "format_image_placeholder",
    "get_file_type",
    "is_data_uri_image",
    "make_error_placeholder",
    "make_keep_link",
    "make_table_safe",
    "normalize_link_target",
    "parse_anchor",
    "process_embedded_image",
    "process_embedded_md_file",
    "process_markdown_links",
    "process_single_embed",
    "reject_declared_input_link_lookups",
    "resolve_data_uri_image_to_temp_file",
    "resolve_embedded_links",
    "resolve_file_path",
    "restore_table_safe_breaks",
    "split_alt_text_and_size",
    "split_markdown_block_segments",
    "split_markdown_inline_segments",
    "split_yaml_front_matter_source",
    "strip_yaml_front_matter",
]
