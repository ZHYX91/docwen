"""Persist the historical effective heading renderer once per user profile.

The retired template-fill switch forced text rendering even when the heading
renderer said word_native. A version in the sparse text layer distinguishes
that legacy conflict from a subsequent deliberate native-renderer choice.
"""

from __future__ import annotations

from copy import deepcopy
from typing import Any

MIGRATION_VERSION_KEY = "heading_numbering_migration_version"


def migrate_heading_numbering_preferences(text: dict[str, Any], template_fill: dict[str, Any]) -> dict[str, Any]:
    result = deepcopy(text)
    if MIGRATION_VERSION_KEY in result or "heading_numbering_render_mode" not in result:
        return result
    if (
        result.get("heading_numbering_render_mode") == "word_native"
        and template_fill.get("gongwen_numbering_text", True) is not False
    ):
        result["heading_numbering_render_mode"] = "text"
    result[MIGRATION_VERSION_KEY] = 1
    return result
