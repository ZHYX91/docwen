"""Shared utilities for the Proofread plugin."""

from __future__ import annotations

import uuid
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


def new_artifact_id() -> str:
    """Return a new unique artifact identifier."""
    return f"proofread-{uuid.uuid4().hex[:12]}"


def file_size(path: str | Path) -> int:
    """Return file size in bytes, or 0 if the file does not exist."""
    try:
        return Path(path).stat().st_size
    except OSError:
        return 0


def request_source_format(context: ConverterContext) -> str:
    """Return the concrete source format frozen at file admission."""
    refs = context.request.input_refs
    if not refs:
        return "unknown"
    return str(refs[0].format or "unknown").strip().lower()


def resolve_proofread_options(
    context: ConverterContext,
    *,
    extra_options: dict[str, Any] | None = None,
) -> dict[str, bool]:
    """Resolve proofread check enable/disable flags.

    Reads defaults from ``context.config`` (the runtime config wiring),
    then overlays ``context.request.options``, and finally any explicit
    *extra_options*.  Request options take precedence over config defaults.

    This is the single place where the proofread plugin consumes config
    from the runtime wiring — satisfying the "at least one real subsystem
    consumes the new wiring" requirement.

    Returns:
        Dict with keys ``enable_symbol_pairing``, ``enable_symbol_correction``,
        ``enable_typos_rule``, ``enable_sensitive_word`` (all ``bool``).
    """
    from docwen_core.options.proofread import resolve_proofread_switches

    config = _safe_get_attr(context, "config", {})
    request = _safe_get_attr(context, "request", None)
    request_options = getattr(request, "options", {}) or {}
    options = dict(request_options) if isinstance(request_options, dict) else {}
    options.update(extra_options or {})
    return resolve_proofread_switches(config, options)


def _safe_get_attr(obj: object, attr: str, default: Any = None) -> Any:
    """Safely get an attribute from an object, returning *default* on failure."""
    return getattr(obj, attr, default)
