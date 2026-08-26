"""YAML front matter extraction and generation for Markdown documents.

F-I2b-001: ``extract_yaml`` — extract YAML front matter and Markdown body.
"""

from __future__ import annotations

import json
import logging
import re

logger = logging.getLogger(__name__)

_YAML_FRONT_PATTERN = re.compile(
    r"^---\s*\r?\n(.*?)\r?\n---\s*(?:\r?\n|$)(.*)$",
    re.DOTALL,
)

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


def _format_yaml_scalar(value: str | bool) -> str:
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


def extract_yaml(content: str) -> tuple[str, str]:
    """Extract YAML front matter and Markdown body from *content*.

    Returns ``(yaml_content, md_body)``.  When no YAML front matter is found,
    returns ``("", content.strip())``.

    Handles BOM stripping and both LF/CRLF line endings.
    """
    if content.startswith("﻿"):
        content = content.lstrip("﻿")

    match = _YAML_FRONT_PATTERN.search(content)
    if match:
        yaml_content = match.group(1).replace("\r\n", "\n").replace("\r", "\n")
        md_body = match.group(2).strip()
        logger.debug("Extracted YAML front matter (%d chars)", len(yaml_content))
        return yaml_content, md_body

    logger.debug("No YAML front matter found; returning full content as body")
    return "", content.strip()


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
    safe_file_stem = _format_yaml_scalar(file_stem)
    lines = ["---", f"{title_key}: {safe_file_stem}", "aliases:", f"  - {safe_file_stem}"]
    if extra:
        for key, value in extra.items():
            lines.append(f"{key}: {_format_yaml_scalar(value)}")
    lines.append("---")
    lines.append("")
    lines.append("")
    return "\n".join(lines)
