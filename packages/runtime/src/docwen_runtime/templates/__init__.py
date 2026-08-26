"""Template discovery registry."""

from docwen_runtime.templates.registry import (
    TemplateIdentityConflictError,
    TemplateInfo,
    TemplateNotFoundError,
    TemplateRegistry,
    TemplateResolutionError,
    is_canonical_template_id,
    validate_template_path,
)

__all__ = [
    "TemplateIdentityConflictError",
    "TemplateInfo",
    "TemplateNotFoundError",
    "TemplateRegistry",
    "TemplateResolutionError",
    "is_canonical_template_id",
    "validate_template_path",
]
