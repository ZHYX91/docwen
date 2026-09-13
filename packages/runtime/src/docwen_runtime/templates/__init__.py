"""Template discovery and user-management APIs."""

from docwen_runtime.templates.manager import TemplateManagementError, TemplateManager
from docwen_runtime.templates.registry import (
    TemplateIdentityConflictError,
    TemplateInfo,
    TemplateNotFoundError,
    TemplateRegistry,
    TemplateResolutionError,
    is_canonical_template_id,
    validate_template_path,
)
from docwen_runtime.templates.state import TemplateStateStore, template_state_path, user_templates_dir

__all__ = [
    "TemplateIdentityConflictError",
    "TemplateInfo",
    "TemplateManagementError",
    "TemplateManager",
    "TemplateNotFoundError",
    "TemplateRegistry",
    "TemplateResolutionError",
    "TemplateStateStore",
    "is_canonical_template_id",
    "template_state_path",
    "user_templates_dir",
    "validate_template_path",
]
