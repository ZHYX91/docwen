"""Build child-process import environments for one source checkout."""

from __future__ import annotations

import os
from collections.abc import Mapping
from pathlib import Path


def source_import_environment(repo: Path, environment: Mapping[str, str]) -> dict[str, str]:
    """Prepend this checkout's package sources without mutating caller state."""
    result = dict(environment)
    sources = [str(path) for path in sorted((repo / "packages").rglob("src")) if path.is_dir()]
    existing = result.get("PYTHONPATH", "")
    inherited = [entry for entry in existing.split(os.pathsep) if entry not in sources] if existing else []
    result["PYTHONPATH"] = os.pathsep.join([*sources, *inherited])
    return result
