"""Scheme format string → Word multi-level list translation layer.

Translates heading numbering scheme format strings (e.g. ``{1.chinese_lower}、``)
into Word ``numbering.xml`` level definitions and analyzes compatibility
constraints arising from Word's single-numFmt-per-level limitation.

This module has zero GUI dependencies and zero I/O — pure functions only.

**Architectural placement:** this is a Word/Office *adaptation* layer for the
scheme format owned by :mod:`docwen_core.text.heading_numbering`. It lives in
core for the same reason :mod:`docwen_core.office_bridge` does — core already
carries Word/Office interop knowledge, and this module is the numbering.xml
analogue. It must not depend on any plugin, runtime, or app package; plugins
(markdown → docx) and GUI (compatibility preview) consume it as a shared core
utility so that GUI does not import plugin internals.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

# ── Public data structures ──────────────────────────────────────────────────


@dataclass(frozen=True)
class WordNumberingLevel:
    """One level of a Word multi-level list definition, translated from a scheme.

    Attributes:
        ilvl: 0-based level index (0 = Heading1, 1 = Heading2, …).
        num_fmt: Word ``numFmt`` value (e.g. ``"chineseCounting"``, ``"decimal"``).
        lvl_text: ``lvlText`` value (e.g. ``"%1、"``, ``"（%2）"``, ``"%1.%2"``).
        suff:  ``"nothing"`` | ``"space"`` | ``"tab"``.
        p_style: ``"Heading1"``, ``"Heading2"``, etc.
        start: Always ``"1"``.
    """

    ilvl: int
    num_fmt: str
    lvl_text: str
    suff: str
    p_style: str
    start: str


@dataclass(frozen=True)
class LevelCompatibility:
    """Compatibility verdict for a single level."""

    verdict: str  # "full" | "unsupported"
    reason: str  # empty if full, explanation if unsupported


@dataclass(frozen=True)
class TranslationResult:
    """Complete translation result for a scheme.

    Attributes:
        verdict: ``"full"`` | ``"approximate"`` | ``"unsupported"``.
        levels: Translated levels (only for compatible levels).
        per_level: 1-based level → compatibility.
        reason: Overall reason for a non-full verdict.
    """

    verdict: str
    levels: list[WordNumberingLevel]
    per_level: dict[int, LevelCompatibility]
    reason: str


# ── Style → numFmt mapping (authoritative) ──────────────────────────────────

STYLE_TO_NUMFMT: dict[str, str] = {
    "chinese_lower": "chineseCounting",
    "chinese_upper": "chineseCountingThousand",
    "arabic_half": "decimal",
    "arabic_full": "decimalFullWidth",
    "arabic_circled": "decimalEnclosedCircleChinese",
    "letter_upper": "upperLetter",
    "letter_lower": "lowerLetter",
    "roman_upper": "upperRoman",
    "roman_lower": "lowerRoman",
}

# Regex for ``{level.style}`` placeholders
_PLACEHOLDER_RE = re.compile(r"\{(\d+)\.(\w+)\}")


# ── Intermediate representation ─────────────────────────────────────────────


@dataclass(frozen=True)
class _PlaceholderRef:
    """A reference to another level's counter in a format string."""

    level: int  # the referenced level (1-based)
    style: str  # the number style identifier


@dataclass(frozen=True)
class _FixedText:
    """Literal text segment in a format string."""

    text: str


_Segment = _PlaceholderRef | _FixedText


# ═════════════════════════════════════════════════════════════════════════════
# Internal helpers
# ═════════════════════════════════════════════════════════════════════════════


def _parse_format_string(fmt: str) -> list[_Segment]:
    """Parse a format string into a list of segments.

    For example, ``"{1.chinese_lower}、"`` becomes
    ``[_PlaceholderRef(1, "chinese_lower"), _FixedText("、")]``.
    """
    segments: list[_Segment] = []
    pos = 0
    for match in _PLACEHOLDER_RE.finditer(fmt):
        start, end = match.span()
        # Text before this placeholder
        if start > pos:
            segments.append(_FixedText(fmt[pos:start]))
        level = int(match.group(1))
        style = match.group(2)
        segments.append(_PlaceholderRef(level, style))
        pos = end
    # Trailing text
    if pos < len(fmt):
        segments.append(_FixedText(fmt[pos:]))
    return segments


def _build_lvl_text(segments: list[_Segment]) -> str:
    """Build the ``lvlText`` value by replacing ``{M.style}`` with ``%M``."""
    parts: list[str] = []
    for seg in segments:
        if isinstance(seg, _PlaceholderRef):
            parts.append(f"%{seg.level}")
        else:
            parts.append(seg.text)
    return "".join(parts)


def _determine_suff(lvl_text: str) -> tuple[str, str]:
    """Determine ``suff`` and the stripped ``lvlText``.

    Returns ``(stripped_lvl_text, suff)``.
    """
    # Strip trailing whitespace (regular space or full-width space)
    stripped = lvl_text.rstrip(" \u3000")
    if len(stripped) < len(lvl_text):
        return stripped, "space"

    # Empty lvlText → no suffix needed
    if not stripped:
        return stripped, "nothing"

    # If ends with a typical separator character → nothing needed
    if stripped[-1] in ("、", "）", ")", "。", "．", ".", ":", "：", ",", "，"):
        return stripped, "nothing"

    # Default: add space before heading text
    return stripped, "space"


def _num_fmt_for_level(
    level: int,
    parsed_formats: dict[int, list[_Segment]],
) -> str | None:
    """Determine the numFmt for *level* based on its self-reference.

    Returns the numFmt string, or ``None`` if the self-referenced style
    is not in ``STYLE_TO_NUMFMT`` (unsupported).
    """
    segments = parsed_formats.get(level, [])
    for seg in segments:
        if isinstance(seg, _PlaceholderRef) and seg.level == level:
            return STYLE_TO_NUMFMT.get(seg.style)
    # Level doesn't reference itself → default fallback
    return "decimal"


def _get_cross_refs(
    level: int,
    parsed_formats: dict[int, list[_Segment]],
) -> list[tuple[int, str]]:
    """Return ``(referenced_level, style)`` pairs for cross-references in *level*."""
    refs: list[tuple[int, str]] = []
    for seg in parsed_formats.get(level, []):
        if isinstance(seg, _PlaceholderRef) and seg.level != level:
            refs.append((seg.level, seg.style))
    return refs


def _get_self_refs(
    level: int,
    parsed_formats: dict[int, list[_Segment]],
) -> list[str]:
    """Return the styles of all self-references in *level*'s format."""
    styles: list[str] = []
    for seg in parsed_formats.get(level, []):
        if isinstance(seg, _PlaceholderRef) and seg.level == level:
            styles.append(seg.style)
    return styles


def _resolve_num_fmt(style: str) -> str | None:
    """Resolve a style string to its Word ``numFmt`` value, or ``None`` if unknown."""
    return STYLE_TO_NUMFMT.get(style)


# ═════════════════════════════════════════════════════════════════════════════
# Public API
# ═════════════════════════════════════════════════════════════════════════════


def translate_scheme(scheme_config: dict) -> TranslationResult:
    """Translate a heading numbering scheme to Word multi-level list definitions.

    Args:
        scheme_config:
            Dict like::

                {
                    "level_1": {"format": "{1.chinese_lower}、"},
                    "level_2": {"format": "（{2.chinese_lower}）"},
                    ...
                }

            Keys are ``"level_1"`` … ``"level_9"`` (only defined levels matter).

    Returns:
        :class:`TranslationResult` with verdict, translated levels, and
        per-level compatibility.
    """
    # ── Parse all defined levels ───────────────────────────────────────
    parsed_formats: dict[int, list[_Segment]] = {}
    defined_levels: set[int] = set()
    for key, value in scheme_config.items():
        if not key.startswith("level_") or not isinstance(value, dict):
            continue
        try:
            level_num = int(key[len("level_") :])
        except ValueError:
            continue
        fmt = value.get("format", "")
        parsed_formats[level_num] = _parse_format_string(str(fmt))
        defined_levels.add(level_num)

    if not defined_levels:
        return TranslationResult(
            verdict="full",
            levels=[],
            per_level={},
            reason="",
        )

    # ── Determine each defined level's numFmt from self-reference ──────
    level_numfmt: dict[int, str | None] = {}
    for level_num in defined_levels:
        nf = _num_fmt_for_level(level_num, parsed_formats)
        level_numfmt[level_num] = nf

    # ── Check for unknown styles ───────────────────────────────────────
    unknown_style_levels: set[int] = set()
    for lvl in defined_levels:
        for seg in parsed_formats[lvl]:
            if isinstance(seg, _PlaceholderRef) and seg.style not in STYLE_TO_NUMFMT:
                unknown_style_levels.add(lvl)

    # ── Compatibility analysis per defined level ───────────────────────
    per_level: dict[int, LevelCompatibility] = {}
    reasons: list[str] = []

    for level_num in sorted(defined_levels):
        # Check 1: unknown self-style → unsupported
        if level_num in unknown_style_levels and level_numfmt.get(level_num) is None:
            per_level[level_num] = LevelCompatibility(
                verdict="unsupported",
                reason=f"Level {level_num} uses unknown number style.",
            )
            reasons.append(f"Level {level_num} uses unknown number style.")
            continue

        # Check 2: multiple self-references with conflicting styles
        self_styles = _get_self_refs(level_num, parsed_formats)
        if len(set(self_styles)) > 1:
            per_level[level_num] = LevelCompatibility(
                verdict="unsupported",
                reason=f"Level {level_num} references itself with multiple "
                f"conflicting styles ({', '.join(self_styles)}). "
                f"Word uses a single numFmt per level.",
            )
            reasons.append(f"Level {level_num}: conflicting self-references.")
            continue

        # Check 3: cross-references that don't match the target level's numFmt
        cross_refs = _get_cross_refs(level_num, parsed_formats)
        mismatches: list[str] = []
        for ref_level, style in cross_refs:
            target_nf = level_numfmt.get(ref_level)
            if target_nf is None:
                # Referenced level is not defined or has unknown style
                mismatches.append(f"references level {ref_level} which has no defined numFmt")
                continue
            ref_nf = _resolve_num_fmt(style)
            if ref_nf is None:
                mismatches.append(f"references level {ref_level} with unknown style '{style}'")
                continue
            if ref_nf != target_nf:
                mismatches.append(
                    f"level {level_num} references level {ref_level} with "
                    f"conflicting style {style} ({ref_nf}), but level {ref_level} "
                    f"uses {STYLE_TO_NUMFMT.get(_self_style(level_numfmt, parsed_formats, ref_level) or '?', '?')}. "
                    f"Word cannot render the same counter with different numFmt."
                )

        if mismatches:
            reason = "; ".join(mismatches)
            per_level[level_num] = LevelCompatibility(
                verdict="unsupported",
                reason=reason,
            )
            reasons.append(f"Level {level_num}: {reason}")
        else:
            per_level[level_num] = LevelCompatibility(
                verdict="full",
                reason="",
            )

    # ── Overall verdict ────────────────────────────────────────────────
    unsupported_1_to_5 = [
        lvl
        for lvl in sorted(defined_levels)
        if lvl <= 5 and per_level.get(lvl, LevelCompatibility("full", "")).verdict == "unsupported"
    ]
    unsupported_6_to_9 = [
        lvl
        for lvl in sorted(defined_levels)
        if lvl >= 6 and per_level.get(lvl, LevelCompatibility("full", "")).verdict == "unsupported"
    ]

    if unsupported_1_to_5:
        verdict = "unsupported"
    elif unsupported_6_to_9:
        verdict = "approximate"
    else:
        verdict = "full"

    reason = _build_reason(per_level, defined_levels)

    # ── Build translated levels (only compatible ones) ─────────────────
    levels: list[WordNumberingLevel] = []
    for level_num in sorted(defined_levels):
        comp = per_level.get(level_num, LevelCompatibility(verdict="full", reason=""))
        if comp.verdict != "full":
            continue

        segments = parsed_formats[level_num]
        lvl_text_raw = _build_lvl_text(segments)
        lvl_text_clean, suff = _determine_suff(lvl_text_raw)
        nf = level_numfmt.get(level_num) or "decimal"
        p_style = f"Heading{level_num}"

        levels.append(
            WordNumberingLevel(
                ilvl=level_num - 1,
                num_fmt=nf,
                lvl_text=lvl_text_clean,
                suff=suff,
                p_style=p_style,
                start="1",
            )
        )

    return TranslationResult(
        verdict=verdict,
        levels=levels,
        per_level=per_level,
        reason=reason,
    )


# ═════════════════════════════════════════════════════════════════════════════
# Helpers
# ═════════════════════════════════════════════════════════════════════════════


def _self_style(
    level_numfmt: dict[int, str | None],
    parsed_formats: dict[int, list[_Segment]],
    level_num: int,
) -> str | None:
    """Return the style name used by *level_num* to reference itself, if any."""
    segments = parsed_formats.get(level_num, [])
    for seg in segments:
        if isinstance(seg, _PlaceholderRef) and seg.level == level_num:
            return seg.style
    return None


def _build_reason(
    per_level: dict[int, LevelCompatibility],
    defined_levels: set[int],
) -> str:
    """Build a human-readable reason string from per-level issues."""
    parts: list[str] = []
    for lvl in sorted(defined_levels):
        comp = per_level.get(lvl)
        if comp and comp.verdict != "full" and comp.reason:
            parts.append(comp.reason)
    return " ".join(parts) if parts else ""
