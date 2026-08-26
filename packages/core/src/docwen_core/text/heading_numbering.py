"""Heading numbering — detection, stripping, and formatting.

This module is the **single source of truth** for heading numbering
detection, stripping, and formatting within the docwen project.

It has **zero dependencies** on any plugin (document, markdown, gongwen)
or on docwen-runtime.  The only intra-project import allowed is
``docwen_core.text.numbering``.

Typical usage::

    from docwen_core.text.heading_numbering import (
        compile_clean_rules_from_data,
        detect_heading_prefix,
        strip_heading_prefix,
        HeadingFormatter,
    )

    rules = compile_clean_rules_from_data(
        [
            {"id": "chapter", "pattern": r"^第一章　", "level": 1},
            {"id": "section", "pattern": r"^（一）", "level": 2},
        ]
    )

    # Detect / strip a heading prefix with request-owned rules
    info = detect_heading_prefix("第一章　总则", rules=rules)
    assert info.rule_id == "legal_unit"
    assert info.prefix == "第一章　"
    assert info.clean_text == "总则"

    prefix, clean = strip_heading_prefix("（一）背景", rules=rules)
    assert prefix == "（一）"

    # Apply a numbering scheme
    config = {"level_1": {"format": "{1.chinese_lower}、"}}
    fmt = HeadingFormatter(config)
    result = fmt.format_heading("概述", 1)   # → "一、概述"
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import Any

from docwen_core.text.numbering import (
    number_to_arabic_full,
    number_to_chinese,
    number_to_chinese_upper,
    number_to_circled,
    number_to_letter_lower,
    number_to_letter_upper,
    number_to_roman_lower,
    number_to_roman_upper,
)

# ═══════════════════════════════════════════════════════════════════════════
# HeadingPrefixInfo
# ═══════════════════════════════════════════════════════════════════════════


@dataclass(frozen=True)
class HeadingPrefixInfo:
    """Describes a detected heading-numbering prefix.

    Attributes:
        prefix: The matched prefix text as it appears in the source
            (e.g. ``"一、"``, ``"（1）"``, ``"1.1 "``).
        clean_text: The remainder of the input after removing the prefix.
        numbering_level: Implied heading level (1–5), or ``0`` when no
            prefix was detected.
        rule_id: Identifier of the rule that matched
            (e.g. ``"chinese_顿号"``, ``"circled"``, ``"hierarchical"``).
    """

    prefix: str
    clean_text: str
    numbering_level: int
    rule_id: str


class NumberingSchemeResolutionError(ValueError):
    """Typed failure raised when an explicitly requested scheme is unusable."""

    def __init__(
        self,
        *,
        error_type: str,
        diagnostic_code: str,
        message: str,
    ) -> None:
        super().__init__(message)
        self.error_type = error_type
        self.diagnostic_code = diagnostic_code


# ═══════════════════════════════════════════════════════════════════════════
# Strip rules — immutable request data (ordered — first match wins)
# ═══════════════════════════════════════════════════════════════════════════

HeadingCleanupRule = tuple[str, re.Pattern[str], int]
HeadingCleanupRules = tuple[HeadingCleanupRule, ...]


def compile_clean_rules_from_data(
    rules_data: Sequence[Mapping[str, Any]],
) -> HeadingCleanupRules:
    """Compile immutable heading-cleanup rules from ordered config data."""

    compiled: list[HeadingCleanupRule] = []
    for rule in rules_data:
        if not rule.get("enabled", True):
            continue
        rule_id = str(rule.get("id", ""))
        pattern_str = str(rule.get("pattern", ""))
        level = int(rule.get("level", 0))
        try:
            compiled_pattern = re.compile(pattern_str)
        except re.error:
            continue
        compiled.append((rule_id, compiled_pattern, level))
    return tuple(compiled)


# ═══════════════════════════════════════════════════════════════════════════
# Public detection / stripping API
# ═══════════════════════════════════════════════════════════════════════════


def detect_heading_prefix(
    text: str,
    *,
    rules: Sequence[HeadingCleanupRule],
) -> HeadingPrefixInfo | None:
    """Detect a heading-numbering prefix in *text*.

    Tries each request-owned clean rule in order. The first match wins.

    Args:
        text: The heading text to inspect.
        rules: Immutable rules compiled from the active request snapshot.

    Returns:
        A :class:`HeadingPrefixInfo` describing the match, or ``None``
        when no rule matches.
    """
    for rule_id, pattern, level in rules:
        m = pattern.match(text)
        if m:
            prefix = m.group(0)
            clean_text = text[m.end() :]
            return HeadingPrefixInfo(
                prefix=prefix,
                clean_text=clean_text,
                numbering_level=level,
                rule_id=rule_id,
            )
    return None


def strip_heading_prefix(
    text: str,
    *,
    rules: Sequence[HeadingCleanupRule],
) -> tuple[str, str]:
    """Convenience wrapper around :func:`detect_heading_prefix`.

    Returns a ``(prefix, clean_text)`` pair.  When no prefix is detected,
    returns ``("", text)``.

    Args:
        text: The heading text to strip.
        rules: Immutable rules compiled from the active request snapshot.

    Returns:
        A two-tuple ``(prefix, clean_text)``.
    """
    info = detect_heading_prefix(text, rules=rules)
    if info is None:
        return ("", text)
    return (info.prefix, info.clean_text)


# ═══════════════════════════════════════════════════════════════════════════
# Style converters  (same mapping as the original formatter)
# ═══════════════════════════════════════════════════════════════════════════

_STYLE_CONVERTERS: dict[str, Any] = {
    "chinese_lower": number_to_chinese,
    "chinese_upper": number_to_chinese_upper,
    "arabic_half": str,
    "arabic_full": number_to_arabic_full,
    "arabic_circled": number_to_circled,
    "letter_upper": number_to_letter_upper,
    "letter_lower": number_to_letter_lower,
    "roman_upper": number_to_roman_upper,
    "roman_lower": number_to_roman_lower,
}

_PLACEHOLDER_RE = re.compile(r"\{(\d+)\.(\w+)\}")


# ═══════════════════════════════════════════════════════════════════════════
# HeadingFormatter  (migrated from docwen_plugin_markdown.numbering_formatter)
# ═══════════════════════════════════════════════════════════════════════════


class HeadingFormatter:
    """Stateful formatter that applies a TOML numbering scheme to headings.

    Maintains per-level counters (1–9) that auto-increment on each call
    to :meth:`format_heading` and reset deeper levels when a shallower
    level is incremented.

    Usage::

        config = {"level_1": {"format": "第{1.chinese_lower}章　"}, ...}
        f = HeadingFormatter(config)
        prefix = f.format_heading("概述", 1)  # → "第一章　概述"
    """

    def __init__(
        self,
        scheme_config: dict[str, Any],
        max_level: int = 9,
    ) -> None:
        """Initialise with a numbering scheme's level config dict.

        Args:
            scheme_config: Dict like ``{"level_1": {"format": "..."},
                "level_2": {"format": "..."}, ...}``.
            max_level: Maximum heading level (default 9).
        """
        self._counters = [0] * max_level
        self._templates: dict[int, str] = {}

        for key, value in scheme_config.items():
            if key.startswith("level_"):
                try:
                    lvl = int(key.split("_", 1)[1])
                except (ValueError, IndexError):
                    continue
                if isinstance(value, dict):
                    self._templates[lvl] = value.get("format", "")

    def format_heading(self, text: str, level: int) -> str:
        """Increment counter for *level*, reset deeper levels, return prefixed text.

        Args:
            text: The raw heading text (without any existing numbering).
            level: The heading level (1‑based).

        Returns:
            The heading text with the numbering prefix prepended, or
            *text* unchanged if no template is defined for this level.
        """
        # Increment counter and reset deeper levels
        idx = level - 1
        if 0 <= idx < len(self._counters):
            self._counters[idx] += 1
            for i in range(idx + 1, len(self._counters)):
                self._counters[i] = 0

        # Fetch template for this level
        template = self._templates.get(level, "")
        if not template:
            return text

        # Resolve placeholders
        prefix = self._resolve_template(template)
        prefix = prefix.replace("&nbsp;", "\u00a0")

        return prefix + text

    def _resolve_template(self, template: str) -> str:
        """Replace ``{级别.样式}`` placeholders with actual counter values.

        Args:
            template: Format string like ``"{1.chinese_lower}、"``.

        Returns:
            Resolved string like ``"一、"``.
        """

        def _replacer(m: re.Match) -> str:
            ref_level = int(m.group(1))
            style = m.group(2)

            if not (1 <= ref_level <= len(self._counters)):
                return str(ref_level)

            counter_value = self._counters[ref_level - 1]
            if counter_value == 0:
                counter_value = 1

            converter = _STYLE_CONVERTERS.get(style)
            if converter is None:
                return str(counter_value)

            try:
                return str(converter(counter_value))
            except Exception:
                return str(counter_value)

        return _PLACEHOLDER_RE.sub(_replacer, template)

    def reset_counters(self) -> None:
        """Reset all counters to zero."""
        self._counters = [0] * len(self._counters)


def resolve_heading_numbering_scheme(
    scheme_id: object,
    registry: Any,
) -> dict[str, dict[str, str]]:
    """Resolve one exact, enabled numbering scheme into formatter config.

    The caller decides whether numbering is requested.  Once requested, this
    function never selects a default and never treats an unusable scheme as a
    successful no-op.  ``registry`` is intentionally duck-typed so Core stays
    independent of the Runtime registry implementation.
    """

    requested = scheme_id.strip() if isinstance(scheme_id, str) else ""
    if not requested:
        raise NumberingSchemeResolutionError(
            error_type="invalid_input",
            diagnostic_code="NUMBERING-SCHEME-REQUIRED",
            message="add_numbering requires a non-empty numbering_scheme",
        )
    if registry is None:
        raise NumberingSchemeResolutionError(
            error_type="capability_unavailable",
            diagnostic_code="NUMBERING-REGISTRY-UNAVAILABLE",
            message=f"Numbering registry is unavailable for scheme '{requested}'",
        )

    try:
        scheme = registry.get_scheme(requested)
    except LookupError as exc:
        raise NumberingSchemeResolutionError(
            error_type="resource_not_found",
            diagnostic_code="NUMBERING-SCHEME-NOT-FOUND",
            message=f"Numbering scheme '{requested}' was not found",
        ) from exc

    if not bool(getattr(scheme, "enabled", True)):
        raise NumberingSchemeResolutionError(
            error_type="capability_unavailable",
            diagnostic_code="NUMBERING-SCHEME-DISABLED",
            message=f"Numbering scheme '{requested}' is disabled",
        )

    levels = getattr(scheme, "levels", None)
    scheme_config: dict[str, dict[str, str]] = {}
    if isinstance(levels, Mapping):
        for key, value in levels.items():
            level_key = str(key)
            format_text = str(value)
            if level_key.startswith("level_") and format_text:
                scheme_config[level_key] = {"format": format_text}
    if not scheme_config:
        raise NumberingSchemeResolutionError(
            error_type="invalid_input",
            diagnostic_code="NUMBERING-SCHEME-NO-LEVELS",
            message=f"Numbering scheme '{requested}' has no usable heading levels",
        )
    return scheme_config
