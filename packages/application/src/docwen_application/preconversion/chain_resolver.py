"""Format chain resolution for an already-admitted source format.

Some document-family Markdown conversions require a hub ``docx`` step.
Plugin-owned document format routes and print routes stay direct so
runtime manifests remain the source of truth.

Examples:
    ("doc", "md") → ["docx", "md"]   # .doc → .docx → .md
    ("doc", "docx") → ["docx"]       # direct plugin route, no app pre-conversion
    ("doc", "docx", action="validate") → ["docx", "docx"]
                                          # .doc → .docx → DOCX action
    ("doc", "pdf") → ["pdf"]         # direct print plugin route
    ("docx", "md") → ["md"]          # directly supported
    ("csv", "xlsx") → ["xlsx"]       # direct plugin route, no Office bridge
    ("xls", "md") → ["md"]           # spreadsheet plugin owns its bridge
"""

from __future__ import annotations

# Category → (hub_format, {non-hub source formats that need application pre-conversion})
_CATEGORY_RULES: dict[str, tuple[str, frozenset[str]]] = {
    "document": ("docx", frozenset({"doc", "wps", "rtf", "odt"})),
}

_DOCUMENT_PRECONVERSION_TARGETS: frozenset[str] = frozenset({"md", "markdown"})


def resolve_chain(
    source_format: str,
    target_format: str,
    *,
    action_name: str = "",
) -> list[str]:
    """Return the format conversion chain.

    Args:
        source_format: The source format (e.g. ``"doc"``, ``"docx"``).
        target_format: The desired target format (e.g. ``"md"``, ``"pdf"``).
        action_name: Optional named action.  A category-level action whose
            target is the document hub must receive a real DOCX, so legacy
            document formats are pre-converted before Runtime dispatch.

    Returns:
        A list of formats representing the conversion chain.
        - ``[]`` — no conversion needed (source == target).
        - ``["md"]`` — direct conversion (source is hub or unknown format).
        - ``["docx"]`` — direct plugin route when target is the hub format.
        - ``["docx", "md"]`` — pre-convert to hub, then to target.

    Raises:
        ValueError: If either side is not a concrete admitted format.
    """
    if not source_format or source_format == "unknown":
        raise ValueError("source_format must be a concrete admitted format")
    if not target_format or target_format == "unknown":
        raise ValueError("target_format must be concrete")
    if source_format == target_format:
        return []

    for _category, (hub, sources) in _CATEGORY_RULES.items():
        if source_format in sources or source_format == hub:
            if source_format == hub:
                return [target_format]
            if target_format == hub and action_name:
                return [hub, target_format]
            if target_format == hub:
                return [target_format]
            if target_format not in _DOCUMENT_PRECONVERSION_TARGETS:
                return [target_format]
            return [hub, target_format]

    # Other admitted formats use their plugin-owned direct route.
    return [target_format]
