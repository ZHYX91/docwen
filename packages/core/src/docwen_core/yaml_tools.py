"""YAML front matter extraction and generation for Markdown documents.

F-I2b-001: ``extract_yaml`` — extract YAML front matter and Markdown body.
"""

from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from yaml.nodes import MappingNode

logger = logging.getLogger(__name__)

_YAML_OPEN = re.compile(r"(?:\ufeff)?---[ \t]*(?:\r\n|[\r\n])")
_YAML_CLOSE = re.compile(r"(?<=[\r\n])(?:---|\.\.\.)[ \t]*(?:\r\n|[\r\n]|\Z)")

_YAML_PLAIN_STRING_PATTERN = re.compile(r"^[^\W\d_][\w .()'/+-]*$")
_YAML_IMPLICIT_STRING_VALUES = frozenset(
    {
        "false",
        "no",
        "null",
        "off",
        "on",
        "true",
        "yes",
    }
)


def _quoted_yaml_string(value: str) -> str:
    """Return a YAML-compatible JSON string with non-printing code points escaped."""
    encoded = json.dumps(value, ensure_ascii=False)
    return "".join(
        f"\\u{codepoint:04x}"
        if 0x007F <= codepoint <= 0x009F or codepoint in {0x2028, 0x2029, 0xFFFE, 0xFFFF}
        else character
        for character in encoded
        for codepoint in (ord(character),)
    )


def format_yaml_scalar(value: str | bool) -> str:
    """Render *value* as one deterministic YAML scalar.

    Ordinary human-readable names retain the existing plain-scalar shape.
    Everything else uses JSON's double-quoted string encoding, which is also
    valid YAML and preserves the value as a string across YAML loaders.
    """
    if isinstance(value, bool):
        return str(value)
    if (
        value
        and value == value.strip()
        and _YAML_PLAIN_STRING_PATTERN.fullmatch(value)
        and value.casefold() not in _YAML_IMPLICIT_STRING_VALUES
    ):
        return value
    return _quoted_yaml_string(value)


@dataclass(frozen=True, slots=True)
class _FrontMatter:
    payload: str
    end: int
    mapping: MappingNode | None


def _read_front_matter(content: str) -> _FrontMatter | None:
    """Retain the parsed mapping used to recognize this exact source prefix."""
    opening = _YAML_OPEN.match(content)
    if opening is None:
        return None
    closing = _YAML_CLOSE.search(content, opening.end())
    if closing is None:
        return None
    payload = content[opening.end() : closing.start()]
    mapping = None
    if payload.strip(" \t\r\n"):
        from docwen_core._yaml_properties import compose_properties

        mapping = compose_properties(payload)
        if mapping is None:
            return None
    return _FrontMatter(payload, closing.end(), mapping)


def yaml_front_matter_end(content: str) -> int:
    """Return the exact prefix length, or zero for ordinary Markdown."""
    front = _read_front_matter(content)
    return front.end if front else 0


def split_yaml_payload(content: str) -> tuple[str | None, str]:
    """Return the exact payload and body, preserving BOM and EOL offsets."""
    front = _read_front_matter(content)
    if front is None:
        return None, content
    return front.payload, content[front.end :]


def extract_yaml_properties(content: str) -> tuple[dict[str, Any], str]:
    """Read template fields while preserving the exact authored body."""
    front = _read_front_matter(content)
    if front is None:
        return {}, content
    body = content[front.end :]
    if front.mapping is None:
        return {}, body
    from docwen_core._yaml_properties import property_values

    return property_values(front.mapping), body


def extract_yaml(content: str) -> tuple[str, str]:
    """Return the normalized YAML payload and the stripped Markdown body."""
    payload, body = split_yaml_payload(content)
    if payload is None:
        return "", content.removeprefix("\ufeff").strip()
    payload = payload.replace("\r\n", "\n").replace("\r", "\n").removesuffix("\n")
    logger.debug("Extracted YAML front matter (%d chars)", len(payload))
    return payload, body.strip()


def generate_basic_yaml_frontmatter(
    file_stem: str,
    *,
    extra: dict[str, str | bool] | None = None,
    yaml_key_labels: object | None = None,
) -> str:
    """Generate a basic YAML front matter block for a Markdown document.

    Produces::

        ---
        title: {file_stem}
        aliases:
          - {file_stem}
        ---

    Extra key/value pairs are appended after the ``title`` line when *extra*
    is provided.

    Args:
        file_stem: File name without extension, used for both ``title``
                   and ``aliases``.
        extra: Optional additional YAML fields (e.g. ``{"source_format": "png"}``).
        yaml_key_labels: Optional pre-resolved labels from an application edge,
                         e.g. ``{"title": "Titel"}``.
    """
    title_key = "title"
    if isinstance(yaml_key_labels, dict):
        label_title = yaml_key_labels.get("title")
        if isinstance(label_title, str) and label_title.strip():
            title_key = label_title.strip()
    safe_file_stem = format_yaml_scalar(file_stem)
    lines = ["---", f"{title_key}: {safe_file_stem}", "aliases:", f"  - {safe_file_stem}"]
    if extra:
        for key, value in extra.items():
            lines.append(f"{key}: {format_yaml_scalar(value)}")
    lines.append("---")
    lines.append("")
    lines.append("")
    return "\n".join(lines)
