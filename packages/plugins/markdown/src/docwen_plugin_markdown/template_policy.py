"""Template filling policy shared by Word and spreadsheet conversion routes."""


def template_list_separator(config: object) -> str:
    """Read the immutable request config without stripping intentional spaces."""
    getter = getattr(config, "get", None)
    value = getter("template_fill.list_separator", None) if callable(getter) else None
    return "、" if value is None else str(value)


def template_gongwen_fields_enabled(config: object) -> bool:
    """Read the single validated request-scoped switch for built-in fields."""
    getter = getattr(config, "get", None)
    return bool(getter("template_fill.gongwen_fields_enabled", True)) if callable(getter) else True
