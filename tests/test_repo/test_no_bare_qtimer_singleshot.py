from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def test_repo_disallows_bare_qtimer_singleshot_outside_safe_timing() -> None:
    project_root = Path(__file__).resolve().parents[2]
    src_root = project_root / "src" / "docwen"
    allowed_files = {
        src_root / "gui" / "core" / "safe_timing.py",
    }
    offenders: list[str] = []

    for path in src_root.rglob("*.py"):
        if path in allowed_files:
            continue
        content = path.read_text(encoding="utf-8")
        if "QTimer.singleShot(" in content:
            offenders.append(path.relative_to(project_root).as_posix())

    assert offenders == []
