from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[4]

LOCAL_SRC_PATHS = [
    PROJECT_ROOT,
    PROJECT_ROOT / "packages" / "core" / "src",
    PROJECT_ROOT / "packages" / "application" / "src",
    PROJECT_ROOT / "packages" / "runtime" / "src",
    PROJECT_ROOT / "packages" / "bundle" / "src",
    PROJECT_ROOT / "packages" / "apps" / "cli" / "src",
]
LOCAL_SRC_PATHS.extend(
    sorted(path / "src" for path in (PROJECT_ROOT / "packages" / "plugins").iterdir() if (path / "src").is_dir())
)

for path in reversed(LOCAL_SRC_PATHS):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))
