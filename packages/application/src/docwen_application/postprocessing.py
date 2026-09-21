"""Application-owned post-processing option contracts.

Front ends normalize user-facing switches, then call these helpers before
route-scoped option validation. Plugins never import or know about this module.
"""

from __future__ import annotations

from typing import Any

from docwen_core.models.request import POSTPROCESS_PROOFREAD_OPTION
from docwen_core.options import PROOFREAD_OPTIONS_SCHEMA

_PROOFREAD_OPTION_KEYS = frozenset(PROOFREAD_OPTIONS_SCHEMA.get("properties", {}))
_MARKDOWN_SOURCE_FORMATS = frozenset({"md", "markdown", "txt"})


def prepare_postprocess_options(
    options: dict[str, Any],
    *,
    source_format: str,
    target_format: str,
    action_name: str,
    proofread_requested: bool = False,
) -> dict[str, Any]:
    """Move canonical proofreading switches into the Application pipeline.

    Named actions such as ``validate`` keep their route-owned options. Ordinary
    Markdown→DOCX conversion consumes proofreading switches as a second-stage
    Application intent. Other routes retain the keys so their normal route
    validation can reject unsupported combinations rather than silently
    accepting them.
    """

    prepared = dict(options)
    source = source_format.strip().lower()
    target = target_format.strip().lower()
    if proofread_requested and (action_name or source not in _MARKDOWN_SOURCE_FORMATS or target != "docx"):
        raise ValueError("--proofread requires an ordinary Markdown-to-DOCX conversion")
    if action_name:
        return prepared
    if source not in _MARKDOWN_SOURCE_FORMATS:
        return prepared
    if target != "docx":
        return prepared

    proofread: dict[str, bool] = {}
    for key in tuple(prepared):
        if key not in _PROOFREAD_OPTION_KEYS:
            continue
        value = prepared.pop(key)
        if type(value) is not bool:
            raise ValueError(f"Proofreading option {key!r} must be boolean")
        proofread[key] = value
    if proofread_requested or any(proofread.values()):
        prepared[POSTPROCESS_PROOFREAD_OPTION] = proofread
    return prepared


def postprocess_proofread_options(request: Any) -> dict[str, bool] | None:
    """Validate and return the DOCX post-processing payload on one request."""

    raw = getattr(request, "options", {}).get(POSTPROCESS_PROOFREAD_OPTION)
    if raw is None:
        return None
    if getattr(request, "action_name", "") or str(getattr(request, "target_format", "")).lower() != "docx":
        raise ValueError("Post-conversion proofreading is supported only for ordinary DOCX conversion")
    source = next(
        (ref for ref in getattr(request, "input_refs", ()) if getattr(ref, "input_role", "") == "source"),
        None,
    )
    if source is None or str(getattr(source, "format", "")).lower() not in _MARKDOWN_SOURCE_FORMATS:
        raise ValueError("Post-conversion proofreading requires an admitted Markdown source")
    if not isinstance(raw, dict):
        raise ValueError("Post-conversion proofreading options must be an object")

    unknown = sorted(set(raw) - _PROOFREAD_OPTION_KEYS)
    if unknown:
        raise ValueError(f"Unsupported post-conversion proofreading option(s): {', '.join(unknown)}")

    normalized: dict[str, bool] = {}
    for key, value in raw.items():
        if type(value) is not bool:
            raise ValueError(f"Post-conversion proofreading option {key!r} must be boolean")
        normalized[key] = value
    return normalized


__all__ = ["postprocess_proofread_options", "prepare_postprocess_options"]
