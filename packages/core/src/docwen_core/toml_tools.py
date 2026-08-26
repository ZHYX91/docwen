"""Shared TOML parsing and editable-document primitives.

Stateless wrappers used by runtime, apps, and plugins. For configuration-
stateful TOML I/O (file registry, config dir, reload), use
``docwen_runtime.config.ConfigLoader`` — those concerns stay in runtime.
This module is the ``yaml_tools`` analogue for TOML: thin tomlkit wrappers
with no dependency on runtime/app/plugin state.
"""

from __future__ import annotations

import logging
import tomllib
from typing import Any

import tomlkit

logger = logging.getLogger(__name__)


def parse_toml_text(text: str) -> dict[str, Any]:
    """Parse TOML text into a plain dictionary.

    Read-only dict semantics (mirrors ``docwen_runtime.toml_io.read_toml_file``
    but for an in-memory string). Use :func:`read_toml_text` when you need a
    mutable document with comment preservation.
    """
    return tomllib.loads(text)


def read_toml_text(text: str) -> Any:
    """Parse a TOML *text* string into a mutable tomlkit document.

    Mutable-document semantics (mirrors ``docwen_runtime.toml_io.load_toml_document``
    but for an in-memory string). The returned document preserves comments and
    key order; mutate it and re-serialize with ``tomlkit.dumps``.
    """
    return tomlkit.parse(text)


def toml_value(value: Any, comment: str = "") -> Any:
    """Construct a tomlkit value with an optional inline *comment*.

    For per-value inline comment scenarios (e.g. proofread dictionary entries
    with user remarks). *value* can be a scalar, list, or dict.
    """
    item = tomlkit.item(value)
    if comment:
        item.comment(comment)
    return item


def toml_table() -> Any:
    """Construct an empty tomlkit table (a section container for per-value
    comment scenarios)."""
    return tomlkit.table()
