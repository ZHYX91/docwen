"""Build immutable heading-cleanup rules from a request config snapshot."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any

from docwen_core.text.heading_numbering import (
    HeadingCleanupRules,
    compile_clean_rules_from_data,
)


def build_heading_cleanup_rules(
    config_snapshot: Mapping[str, Any] | None,
) -> HeadingCleanupRules:
    """Return ordered compiled rules owned by one conversion request."""

    snapshot = config_snapshot if isinstance(config_snapshot, Mapping) else {}
    numbering = snapshot.get("numbering", {})
    cleanup = numbering.get("cleanup", {}) if isinstance(numbering, Mapping) else {}
    if not isinstance(cleanup, Mapping):
        return ()

    raw_rules = cleanup.get("rules", [])
    rules = [rule for rule in raw_rules if isinstance(rule, Mapping)] if isinstance(raw_rules, list) else []
    settings = cleanup.get("settings", {})
    raw_order = settings.get("order", []) if isinstance(settings, Mapping) else []
    order = [str(rule_id) for rule_id in raw_order] if isinstance(raw_order, list) else []

    rules_by_id = {str(rule.get("id", "")): rule for rule in rules}
    ordered = [rules_by_id[rule_id] for rule_id in order if rule_id in rules_by_id]
    ordered.extend(rule for rule in rules if str(rule.get("id", "")) not in order)
    return compile_clean_rules_from_data(ordered)


__all__ = ["build_heading_cleanup_rules"]
