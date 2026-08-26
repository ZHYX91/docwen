"""Proofread rule constants, schema, and diagnostic codes.

This module is intentionally pure data:

1. ``DEFAULT_*`` constants mirror the proofread TOML defaults seeded by
   ``docwen_runtime.config.loader`` and serve as in-code fallbacks for
   :class:`TextValidator`.
2. ``PROOFREAD_OPTIONS_SCHEMA`` describes request/config options.
3. ``DIAGNOSTIC_*`` constants keep proofread diagnostic codes centralized.

Runtime-side TOML loading and normalization now lives in
``docwen_runtime.config.proofread_rules``. GUI editors read/write TOML
through core/runtime helpers, while validators consume
``context.proofread_rules`` with zero filesystem I/O.
"""

from __future__ import annotations

# ── Symbol pairs (port from proofread_pairing.toml) ─────────────────────
# Each entry is [opening, closing]. Canonical default; ConfigLoader seeds
# proofread/pairs.toml with the same data. Users edit the
# TOML; this constant is the in-code fallback for TextValidator defaults.
DEFAULT_SYMBOL_PAIRS: list[tuple[str, str]] = [
    ("(", ")"),
    ("（", "）"),
    ("[", "]"),
    ("【", "】"),
    ("{", "}"),
    ("「", "」"),
    ("'", "'"),
    ('"', '"'),
    ("‘", "’"),
    ("“", "”"),
    ("《", "》"),
]

# ── Symbol correction map (port from proofread_symbols.toml) ───────────
# correct_symbol → [incorrect_variants]. Canonical default; ConfigLoader
# seeds proofread/symbol_map.toml with the same data.
DEFAULT_SYMBOL_MAP: dict[str, list[str]] = {
    "0": ["０"],
    "1": ["１"],
    "2": ["２"],
    "3": ["３"],
    "4": ["４"],
    "5": ["５"],
    "6": ["６"],
    "7": ["７"],
    "8": ["８"],
    "9": ["９"],
}

# ── Common Chinese typos (empty by default — user-populated) ───────────
# correct_word → [common_misspellings]. Users populate via typos_dict.toml.
DEFAULT_TYPOS_MAP: dict[str, list[str]] = {}
# Note: "的"/"地"/"得" grammatical particle disambiguation requires
# context-aware rules and is intentionally excluded from simple char mapping.

# ── Sensitive words (empty by default — user-populated) ────────────────
# word → [exception_contexts]. Users populate via sensitive_dict.toml.
DEFAULT_SENSITIVE_WORDS: dict[str, list[str]] = {}


# ── Proofread options schema ───────────────────────────────────────────

PROOFREAD_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "enable_symbol_pairing": {
            "type": "boolean",
            "default": True,
            "description": "Check for unmatched symbol pairs (brackets, quotes, etc.)",
        },
        "enable_symbol_correction": {
            "type": "boolean",
            "default": True,
            "description": "Check for incorrect symbol usage (e.g. fullwidth digits)",
        },
        "enable_typos_rule": {
            "type": "boolean",
            "default": True,
            "description": "Check for common Chinese typographical errors",
        },
        "enable_sensitive_word": {
            "type": "boolean",
            "default": True,
            "description": "Check for sensitive words",
        },
        "skip_code_blocks": {
            "type": "boolean",
            "default": True,
            "description": "Skip code-like DOCX paragraphs during proofreading",
        },
        "skip_quote_blocks": {
            "type": "boolean",
            "default": False,
            "description": "Skip quote-like DOCX paragraphs during proofreading",
        },
    },
    "required": [],
}

# ── Diagnostic codes ───────────────────────────────────────────────────
DIAGNOSTIC_OK = "PROOFREAD-OK"
DIAGNOSTIC_SKIPPED = "PROOFREAD-SKIPPED"
DIAGNOSTIC_INVALID_INPUT = "PROOFREAD-INVALID-INPUT"
DIAGNOSTIC_ERROR = "PROOFREAD-ERROR"
DIAGNOSTIC_CORRUPTED = "PROOFREAD-CORRUPTED-DOCX"
DIAGNOSTIC_CANCELLED = "PROOFREAD-CANCELLED"
