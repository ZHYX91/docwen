"""仓库门禁测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def test_no_egg_info_under_src() -> None:
    root = Path(__file__).resolve().parents[2]
    src_dir = root / "src"
    assert src_dir.is_dir()

    egg_infos = sorted(p for p in src_dir.rglob("*.egg-info") if p.is_dir())
    assert not egg_infos, f"Found build artifacts under src/: {[str(p) for p in egg_infos]}"

