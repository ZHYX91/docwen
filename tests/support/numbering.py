"""Explicit repository numbering fixtures for tests."""

from __future__ import annotations

import tomllib
from pathlib import Path
from typing import Any

from docwen_runtime.numbering import NumberingSchemeRegistry

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def repository_numbering_snapshot(*, locale: str = "zh_CN") -> dict[str, Any]:
    """Load a fresh numbering snapshot from the checked-in canonical config."""

    add_config = tomllib.loads((PROJECT_ROOT / "configs" / "numbering" / "add.toml").read_text(encoding="utf-8"))
    return {
        "gui": {"language": {"locale": locale}},
        "numbering": {"add": add_config},
    }


def repository_numbering_registry(*, locale: str = "zh_CN") -> NumberingSchemeRegistry:
    """Build a registry without process-global config or resource discovery."""

    return NumberingSchemeRegistry.from_config_snapshot(
        repository_numbering_snapshot(locale=locale),
        locale_path=PROJECT_ROOT / "i18n" / "locales" / f"{locale}.toml",
    )
