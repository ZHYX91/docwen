"""仓库门禁测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def test_no_egg_info_under_src() -> None:
    root = Path(__file__).resolve().parents[2]
    src_dir = root / "src"

    egg_infos = sorted(p for p in src_dir.rglob("*.egg-info") if p.is_dir() and p.name != "docwen.egg-info")
    assert not egg_infos, (
        "Found setuptools build metadata under src/ (repo hygiene failure, not a product-code regression): "
        f"{[str(p) for p in egg_infos]}. "
        "These directories are usually transient local artifacts from editable installs or build/package commands. "
        "Clean them with `python scripts/clean/clean_build.py` and rerun repo hygiene tests. "
        "If this repeats, inspect recent local packaging/install commands rather than treating it as a core conversion bug."
    )
