"""Policy for preserving intermediate pre-conversion files.

Controlled by config key ``output.intermediate_files.save_to_output``.
"""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import Any


def should_save_intermediates(*, config_snapshot: Mapping[str, Any]) -> bool:
    """Return whether intermediate pre-conversion files should be
    saved to the output directory.

    Config key: ``output.intermediate_files.save_to_output`` (default ``false``).
    The admitted request snapshot is authoritative, including an empty mapping.

    """
    output = config_snapshot.get("output", {})
    if not isinstance(output, Mapping):
        return False
    intermediate_files = output.get("intermediate_files", {})
    if not isinstance(intermediate_files, Mapping):
        return False
    return bool(intermediate_files.get("save_to_output", False))


def build_intermediate_record_if_enabled(
    temp_path: str,
    original_stem: str,
    source_format: str,
    *,
    target_format: str,
    backend: str,
    config_snapshot: Mapping[str, Any],
) -> dict[str, Any] | None:
    """Build a serializable request record for a preserved intermediate.

    The runtime finalizer consumes this record and places the file alongside
    the primary output, keeping final output writes in one layer.
    """
    if not should_save_intermediates(config_snapshot=config_snapshot):
        return None

    path = Path(temp_path)
    return {
        "staging_path": str(path),
        "suggested_name": f"{original_stem}_from{source_format.capitalize()}{path.suffix}",
        "source_format": source_format,
        "target_format": target_format,
        "backend": backend,
    }
