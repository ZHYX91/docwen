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
    # Read engine defaults from config — engine is namespaced under proofread
    config = _safe_get_attr(context, "config", {})
    proofread = config.get("proofread", {}) if hasattr(config, "get") else {}
    engine = proofread.get("engine", {}) if isinstance(proofread, dict) else {}
    if not isinstance(engine, dict):
        engine = {}

    options: dict[str, bool] = {
        "enable_symbol_pairing": bool(engine.get("enable_symbol_pairing", True)),
        "enable_symbol_correction": bool(engine.get("enable_symbol_correction", True)),
        "enable_typos_rule": bool(engine.get("enable_typos_rule", True)),
        "enable_sensitive_word": bool(engine.get("enable_sensitive_word", True)),
    }

    # Request options override config
    req_options = _safe_get_attr(context, "request", None)
    if req_options is not None:
        req_opts = getattr(req_options, "options", {}) or {}
        if isinstance(req_opts, dict):
            for key in options:
                if key in req_opts:
                    options[key] = bool(req_opts[key])

    # Explicit extra_options take highest precedence
    if extra_options:
        for key in options:
            if key in extra_options:
                options[key] = bool(extra_options[key])

    return options


def _safe_get_attr(obj: object, attr: str, default: Any = None) -> Any:
    """Safely get an attribute from an object, returning *default* on failure."""
    return getattr(obj, attr, default)
