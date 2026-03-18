from __future__ import annotations

from docwen.docx_spell.anchor_report import (
    build_anchor_report_markdown,
    extract_comment_texts_from_comments_xml,
    extract_occurrences_from_document_xml,
    read_docx_part,
    redact_text,
)

__all__ = [
    "build_anchor_report_markdown",
    "extract_comment_texts_from_comments_xml",
    "extract_occurrences_from_document_xml",
    "read_docx_part",
    "redact_text",
]
