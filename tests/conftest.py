from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest


def _ensure_docwen_importable() -> None:
    if importlib.util.find_spec("docwen") is not None:
        return

    project_root = Path(__file__).resolve().parents[1]
    src_path = project_root / "src"
    pkg_init = src_path / "docwen" / "__init__.py"
    if not pkg_init.is_file():
        raise ModuleNotFoundError("docwen is not importable and src/docwen is missing")

    src_path_str = str(src_path)
    if src_path_str not in sys.path:
        sys.path.insert(0, src_path_str)

    if importlib.util.find_spec("docwen") is None:
        raise ModuleNotFoundError("docwen is not importable after adding src to sys.path")


_ensure_docwen_importable()


@pytest.fixture(scope="session")
def project_root() -> Path:
    return Path(__file__).resolve().parents[1]
