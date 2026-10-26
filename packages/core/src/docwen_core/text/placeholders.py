"""Collision-checked text placeholders with a single nonrecursive restoration pass."""

from __future__ import annotations

import re
import uuid
from collections.abc import Mapping


def placeholder_namespace(authored_text: str, *, prefix: str) -> str:
    namespace = f"{prefix}{uuid.uuid4().hex.upper()}X"
    while namespace in authored_text:
        namespace += "X"
    return namespace


def restore_placeholders(text: str, replacements: Mapping[str, str]) -> str:
    if not replacements:
        return text
    pattern = re.compile("|".join(re.escape(token) for token in sorted(replacements, key=len, reverse=True)))
    return pattern.sub(lambda match: replacements[match.group()], text)
