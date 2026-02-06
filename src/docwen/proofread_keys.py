from __future__ import annotations

from typing import Dict, Mapping, Optional, Tuple

SYMBOL_PAIRING = "symbol_pairing"
SYMBOL_CORRECTION = "symbol_correction"
TYPOS_RULE = "typos_rule"
SENSITIVE_WORD = "sensitive_word"

PROOFREAD_OPTION_KEYS: Tuple[str, str, str, str] = (
    SYMBOL_PAIRING,
    SYMBOL_CORRECTION,
    TYPOS_RULE,
    SENSITIVE_WORD,
)


def normalize_proofread_options(options: Optional[Mapping[str, bool]]) -> Dict[str, bool]:
    if not options:
        return {key: False for key in PROOFREAD_OPTION_KEYS}
    return {key: bool(options.get(key, False)) for key in PROOFREAD_OPTION_KEYS}

