"""Cross-platform link parser and runtime isolation contracts."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]


def test_shipped_link_config_does_not_advertise_legacy_error_templates() -> None:
    config = (ROOT / "configs/link.toml").read_text(encoding="utf-8")

    assert "file_not_found_text" not in config
    assert "circular_text" not in config
    assert "max_depth_text" not in config


def test_link_embed_runtime_has_no_network_client_import() -> None:
    link_dir = ROOT / "packages/core/src/docwen_core/links"
    forbidden_roots = {"httpx", "requests", "urllib3", "aiohttp"}
    forbidden_urllib = {"request"}

    for path in link_dir.glob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                assert not ({alias.name.split(".", 1)[0] for alias in node.names} & forbidden_roots), path
            elif isinstance(node, ast.ImportFrom):
                assert (node.module or "").split(".", 1)[0] not in forbidden_roots, path
                if node.module == "urllib":
                    assert not ({alias.name for alias in node.names} & forbidden_urllib), path
