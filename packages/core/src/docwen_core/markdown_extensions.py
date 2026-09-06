"""Explicit, request-scoped Markdown dialect choices shared by all frontends."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, fields
from typing import Literal


@dataclass(frozen=True, slots=True)
class MarkdownExtensions:
    structural_tables: bool = False
    captions_references: bool = False
    extended_headings: bool = False
    typed_endnotes: bool = False

    @classmethod
    def obsidian(cls) -> MarkdownExtensions:
        return cls(True, True, True, True)

    def to_dict(self) -> dict[str, bool]:
        return {field.name: getattr(self, field.name) for field in fields(self)}


EXTENSION_NAMES = tuple(MarkdownExtensions().to_dict())
MARKDOWN_EXTENSIONS_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        direction: {
            "type": "object",
            "additionalProperties": False,
            "properties": {name: {"type": "boolean"} for name in EXTENSION_NAMES},
        }
        for direction in ("input", "output")
    },
}


def resolve_markdown_extensions(
    options: Mapping[str, object],
    config: object,
    *,
    direction: Literal["input", "output"],
) -> MarkdownExtensions:
    """Explicit request values override the immutable configuration snapshot."""

    getter = getattr(config, "get", None)
    raw_config = getter(f"conversion.markdown_extensions.{direction}", {}) if callable(getter) else {}
    configured = raw_config if isinstance(raw_config, Mapping) else {}
    raw_options = options.get("markdown_extensions", {})
    if not isinstance(raw_options, Mapping) or set(raw_options) - {"input", "output"}:
        raise ValueError("markdown_extensions must contain only input and output objects")
    selected = raw_options.get(direction, {})
    if not isinstance(selected, Mapping) or set(selected) - set(EXTENSION_NAMES):
        raise ValueError(f"markdown_extensions.{direction} contains an unknown extension")
    values: dict[str, bool] = {}
    for name in EXTENSION_NAMES:
        value = selected.get(name, configured.get(name, False))
        if type(value) is not bool:
            raise ValueError(f"markdown_extensions.{direction}.{name} must be a boolean")
        values[name] = value
    return MarkdownExtensions(**values)


__all__ = [
    "EXTENSION_NAMES",
    "MARKDOWN_EXTENSIONS_OPTIONS_SCHEMA",
    "MarkdownExtensions",
    "resolve_markdown_extensions",
]
