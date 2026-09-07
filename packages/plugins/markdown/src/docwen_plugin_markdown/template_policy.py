"""Template filling policy shared by Word and spreadsheet conversion routes."""


def template_list_separator(config: object) -> str:
    """Read the immutable request config without stripping intentional spaces."""
    getter = getattr(config, "get", None)
    value = getter("template_fill.list_separator", None) if callable(getter) else None
    return "、" if value is None else str(value)
