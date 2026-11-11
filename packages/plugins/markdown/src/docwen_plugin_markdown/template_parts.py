"""Template story owners shared by all placeholder consumers."""

from __future__ import annotations

from typing import Any

from docwen_core.docx_package_parts import DocxPartError, document_story_parts
from docwen_plugin_markdown.template_errors import TemplatePackageError


def template_story_parts(document: Any) -> tuple[Any, ...]:
    try:
        return document_story_parts(document)
    except DocxPartError as exc:
        raise TemplatePackageError(str(exc)) from exc
