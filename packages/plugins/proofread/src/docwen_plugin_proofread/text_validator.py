"""TextValidator — rule-based text proofreading engine.

Performs four checks on text:
1. Symbol pairing — unmatched brackets, quotes, etc.
2. Symbol correction — wrong symbol variants (fullwidth digits, etc.)
3. Typos — common Chinese typographical errors
4. Sensitive words — configurable sensitive word list (default: empty)

The engine is self-contained — all rules are supplied at construction
time. Defaults come from :mod:`docwen_plugin_proofread.rules` (the
``DEFAULT_*`` constants), which mirror the ConfigLoader-managed TOML
seeds. In production, validators load the TOML files via the
``load_*_rules`` functions and pass the parsed data in explicitly; the
in-code defaults are only used when no data is supplied (e.g. tests).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from docwen_plugin_proofread.rules import (
    DEFAULT_SENSITIVE_WORDS,
    DEFAULT_SYMBOL_MAP,
    DEFAULT_SYMBOL_PAIRS,
    DEFAULT_TYPOS_MAP,
)

if TYPE_CHECKING:
    from collections.abc import Iterable


@dataclass(frozen=True)
class TextError:
    """A single proofreading issue located in text."""

    start_pos: int
    end_pos: int
    error_text: str
    suggestion: str
    error_type: str
    source: str
    replacement: str | None = None
    """Explicit machine-applicable replacement, when the rule defines one.

    ``suggestion`` remains presentation text and must never be interpreted as
    an edit.  Only typo and symbol-correction rules populate this field.
    """


# ── Stable rule-key mapping ──────────────────────────────────────────

_RULE_KEY_MAP: dict[str, str] = {
    "typo": "typo",
    "symbol": "symbol_correct",
    "pairing": "symbol_pair",
    "sensitive": "sensitive",
}

# ── I18n error type labels ──────────────────────────────────────────────

_ERROR_TYPE_I18N: dict[str, dict[str, str]] = {
    "en": {"sensitive": "Sensitive Word", "typo": "Typo", "punctuation": "Punctuation", "symbol": "Unmatched Symbol"},
    "zh_CN": {"sensitive": "敏感词", "typo": "错别字", "punctuation": "标点错误", "symbol": "符号不匹配"},
}


def _get_error_label(error_type: str, lang: str = "en") -> str:
    """Return the i18n label for *error_type* in *lang*.

    Falls back to the English label, then to *error_type* itself.
    """
    return _ERROR_TYPE_I18N.get(lang, _ERROR_TYPE_I18N["en"]).get(error_type, error_type)


def rule_key(source: str) -> str:
    """Return a stable machine-readable key for an error source."""
    return _RULE_KEY_MAP.get(source, "unknown")


class TextValidator:
    """Rule-based text validator.

    Parameters
    ----------
    symbol_pairs:
        List of ``(open, close)`` tuples for pairing checks.
        Defaults to :data:`DEFAULT_SYMBOL_PAIRS`.
    symbol_map:
        Dict of ``correct → [wrong_variants]`` for symbol correction.
        Defaults to :data:`DEFAULT_SYMBOL_MAP`.
    typos_map:
        Dict of ``correct_word → [wrong_spellings]`` for typo detection.
        Defaults to :data:`DEFAULT_TYPOS_MAP` (empty — user-populated).
    sensitive_words:
        Dict of ``word → [exception_contexts]`` for sensitive word detection.
        Defaults to :data:`DEFAULT_SENSITIVE_WORDS` (empty — user-populated).
    enabled:
        Which checks are active.  Keys: ``symbol_pairing``, ``symbol_correction``,
        ``typos_rule``, ``sensitive_word``. Defaults to all-on except
        ``sensitive_word``.
    """

    def __init__(
        self,
        *,
        symbol_pairs: Iterable[tuple[str, str]] | None = None,
        symbol_map: dict[str, list[str]] | None = None,
        typos_map: dict[str, list[str]] | None = None,
        sensitive_words: dict[str, list[str]] | None = None,
        enabled: dict[str, bool] | None = None,
        lang: str = "en",
    ) -> None:
        if enabled is None:
            enabled = {
                "symbol_pairing": True,
                "symbol_correction": True,
                "typos_rule": True,
                "sensitive_word": False,
            }
        self._enabled = dict(enabled)
        self._lang = lang

        # Symbol pairing
        self._symbol_pairs: list[tuple[str, str]] = (
            list(symbol_pairs) if symbol_pairs is not None else list(DEFAULT_SYMBOL_PAIRS)
        )

        # Symbol correction
        self._symbol_lookup: dict[str, str] = {}
        self._symbol_pattern: re.Pattern[str] | None = None
        sm = DEFAULT_SYMBOL_MAP if symbol_map is None else symbol_map
        for correct, wrongs in sm.items():
            for w in wrongs:
                if w:
                    self._symbol_lookup[w] = correct
        self._symbol_pattern = _compile_alternation(self._symbol_lookup.keys())

        # Typos
        self._typos_lookup: dict[str, str] = {}
        self._typos_pattern: re.Pattern[str] | None = None
        tm = DEFAULT_TYPOS_MAP if typos_map is None else typos_map
        for correct, wrongs in tm.items():
            for w in wrongs:
                if w:
                    self._typos_lookup[w] = correct
        self._typos_pattern = _compile_alternation(self._typos_lookup.keys())

        # Sensitive words
        self._sensitive_lookup: dict[str, tuple[str, list[str] | None]] = {}
        self._sensitive_pattern: re.Pattern[str] | None = None
        sw = DEFAULT_SENSITIVE_WORDS if sensitive_words is None else sensitive_words
        for word, exceptions in sw.items():
            if word:
                self._sensitive_lookup[word.casefold()] = (word, exceptions)
        keys = [orig for orig, _ in self._sensitive_lookup.values()]
        self._sensitive_pattern = _compile_alternation(keys, flags=re.IGNORECASE)

    # ── Public API ─────────────────────────────────────────────────────

    def validate_text(self, text: str) -> list[TextError]:
        """Run all enabled checks on *text* and return a list of errors."""
        if not text:
            return []

        errors: list[TextError] = []

        # Keep user-facing issue order deterministic.
        if self._enabled.get("sensitive_word", False):
            errors.extend(self._sensitive_check(text))

        if self._enabled.get("typos_rule", True):
            errors.extend(self._typos_check(text))

        if self._enabled.get("symbol_correction", True):
            errors.extend(self._symbol_correction_check(text))

        if self._enabled.get("symbol_pairing", True):
            errors.extend(self._symbol_pairing_check(text))

        return errors

    def any_enabled(self) -> bool:
        """Return True if at least one check is enabled."""
        return any(
            self._enabled.get(k, True)
            for k in (
                "symbol_pairing",
                "symbol_correction",
                "typos_rule",
                "sensitive_word",
            )
        )

    # ── Individual checks ──────────────────────────────────────────────

    def _sensitive_check(self, text: str) -> list[TextError]:
        if not self._sensitive_pattern:
            return []
        errors: list[TextError] = []
        for match in self._sensitive_pattern.finditer(text):
            start, end = match.span()
            hit = match.group(0)
            _sensitive_word, exceptions = self._sensitive_lookup[hit.casefold()]
            if exceptions:
                # Check if hit is within an exception context
                max_ex_len = max(len(ex) for ex in exceptions)
                ctx_start = max(0, start - max_ex_len)
                ctx_end = min(len(text), end + max_ex_len)
                context = text[ctx_start:ctx_end]
                is_exception = False
                for exc in exceptions:
                    if exc not in context:
                        continue
                    exc_idx = context.find(exc)
                    exc_s = ctx_start + exc_idx
                    exc_e = exc_s + len(exc)
                    if exc_s <= start and end <= exc_e:
                        is_exception = True
                        break
                if is_exception:
                    continue
            errors.append(
                TextError(
                    start_pos=start,
                    end_pos=end,
                    error_text=hit,
                    suggestion=_get_error_label("sensitive", self._lang),
                    error_type=_get_error_label("sensitive", self._lang),
                    source="sensitive",
                )
            )
        return errors

    def _typos_check(self, text: str) -> list[TextError]:
        if not self._typos_pattern:
            return []
        return [
            TextError(
                start_pos=m.start(),
                end_pos=m.end(),
                error_text=m.group(0),
                suggestion=self._typos_lookup[m.group(0)],
                error_type=_get_error_label("typo", self._lang),
                source="typo",
                replacement=self._typos_lookup[m.group(0)],
            )
            for m in self._typos_pattern.finditer(text)
        ]

    def _symbol_correction_check(self, text: str) -> list[TextError]:
        if not self._symbol_pattern:
            return []
        return [
            TextError(
                start_pos=m.start(),
                end_pos=m.end(),
                error_text=m.group(0),
                suggestion=self._symbol_lookup[m.group(0)],
                error_type=_get_error_label("punctuation", self._lang),
                source="symbol",
                replacement=self._symbol_lookup[m.group(0)],
            )
            for m in self._symbol_pattern.finditer(text)
        ]

    def _symbol_pairing_check(self, text: str) -> list[TextError]:
        errors: list[TextError] = []
        if not self._symbol_pairs:
            return errors

        stacks: dict[str, list[int]] = {pair[0]: [] for pair in self._symbol_pairs}
        closing_map: dict[str, str] = {pair[1]: pair[0] for pair in self._symbol_pairs}
        symmetric_symbols = {opening for opening, closing in self._symbol_pairs if opening == closing}

        for i, char in enumerate(text):
            if char in symmetric_symbols:
                stack = stacks[char]
                if _is_apostrophe_usage(text, i, has_opening=bool(stack)):
                    continue
                if stack:
                    stack.pop()
                else:
                    stack.append(i)
            elif char in stacks:
                stacks[char].append(i)
            elif char in closing_map:
                opening_char = closing_map[char]
                if _is_apostrophe_usage(text, i, has_opening=bool(stacks[opening_char])):
                    continue
                if stacks[opening_char]:
                    stacks[opening_char].pop()
                else:
                    errors.append(
                        TextError(
                            start_pos=i,
                            end_pos=i + 1,
                            error_text=char,
                            suggestion=_get_error_label("symbol", self._lang),
                            error_type=_get_error_label("symbol", self._lang),
                            source="pairing",
                        )
                    )

        for opening_char, stack in stacks.items():
            for pos in stack:
                errors.append(
                    TextError(
                        start_pos=pos,
                        end_pos=pos + 1,
                        error_text=opening_char,
                        suggestion=_get_error_label("symbol", self._lang),
                        error_type=_get_error_label("symbol", self._lang),
                        source="pairing",
                    )
                )

        return errors


# ── Helpers ────────────────────────────────────────────────────────────


_DECADE_ELISION_RE = re.compile(r"\d{2}s?(?!\w)", re.IGNORECASE)
_LEADING_ELISION_RE = re.compile(r"[A-Za-z]+")
_LEADING_ELISION_WORDS = frozenset(
    {
        "bout",
        "cause",
        "cept",
        "em",
        "fore",
        "gainst",
        "n",
        "neath",
        "round",
        "scuse",
        "til",
        "tis",
        "twas",
        "twere",
        "twill",
        "twould",
    }
)


def _is_leading_elision(text: str, index: int) -> bool:
    """Return whether an apostrophe starts a narrowly recognized elision."""

    tail = text[index + 1 :]
    if _DECADE_ELISION_RE.match(tail):
        return True
    match = _LEADING_ELISION_RE.match(tail)
    return bool(match and match.group(0).casefold() in _LEADING_ELISION_WORDS)


def _is_apostrophe_usage(text: str, index: int, *, has_opening: bool) -> bool:
    """Return whether a straight/curly single quote is word punctuation.

    A quote between word characters is a contraction apostrophe even inside
    a quoted span.  A trailing apostrophe without a live opening delimiter is
    treated as possessive punctuation; when an opening delimiter exists, the
    same position closes that delimiter.
    """

    char = text[index]
    if char not in {"'", "’"}:
        return False
    previous_is_word = index > 0 and text[index - 1].isalnum()
    next_is_word = index + 1 < len(text) and text[index + 1].isalnum()

    # Preserve quote-opening semantics for arbitrary words/numbers.  Only
    # conventional two-digit decades and an explicit set of common English
    # elisions (for example '90s and ’tis) are exempted.
    if not has_opening and not previous_is_word and next_is_word and _is_leading_elision(text, index):
        return True
    if not previous_is_word:
        return False
    if next_is_word:
        return True
    return not has_opening


def _compile_alternation(words: Iterable[str], *, flags: int = 0) -> re.Pattern[str] | None:
    """Compile a regex that matches any of *words* (longest-match-first)."""
    items = sorted({w for w in words if w}, key=len, reverse=True)
    if not items:
        return None
    return re.compile("|".join(re.escape(w) for w in items), flags)
