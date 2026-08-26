"""Build proofread rule bundles from runtime config snapshots."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from typing import Any

from docwen_core.models.proofread import ProofreadRules


def build_proofread_rules(config_snapshot: Mapping[str, Any] | None) -> ProofreadRules:
    """Create a normalized proofread rule bundle from merged config data.

    ``config_snapshot`` is the full merged config dict (e.g. from
    ``ConfigLoader.config.as_dict()``).  Proofread data lives under the
    ``proofread`` namespace — each sub-key (pairs, symbol_map, typos,
    sensitive_words) has an ``items`` or ``entries`` wrapper from its
    base TOML file.
    """
    data = dict(config_snapshot or {})
    proofread = data.get("proofread", {}) if isinstance(data.get("proofread"), dict) else {}

    pairs = proofread.get("pairs", {}) if isinstance(proofread.get("pairs"), dict) else {}
    symbol_map = proofread.get("symbol_map", {}) if isinstance(proofread.get("symbol_map"), dict) else {}
    typos = proofread.get("typos", {}) if isinstance(proofread.get("typos"), dict) else {}
    sensitive = proofread.get("sensitive_words", {}) if isinstance(proofread.get("sensitive_words"), dict) else {}

    return ProofreadRules(
        symbol_pairs=_normalize_pairs(pairs.get("items", ())),
        symbol_map=_normalize_string_map(symbol_map.get("entries", {})),
        typos_map=_normalize_string_map(typos.get("entries", {})),
        sensitive_words=_normalize_string_map(sensitive.get("entries", {})),
    )


def _normalize_pairs(raw: Any) -> tuple[tuple[str, str], ...]:
    if not isinstance(raw, list):
        return ()

    pairs: list[tuple[str, str]] = []
    for item in raw:
        if isinstance(item, dict):
            opening = item.get("open", item.get("source"))
            closing = item.get("close", item.get("target"))
            if opening is None or closing is None:
                continue
            pairs.append((str(opening), str(closing)))
            continue
        if isinstance(item, (list, tuple)) and len(item) >= 2:
            pairs.append((str(item[0]), str(item[1])))
    return tuple(pairs)


def _normalize_string_map(raw: Any) -> dict[str, tuple[str, ...]]:
    if not isinstance(raw, Mapping):
        return {}

    normalized: dict[str, tuple[str, ...]] = {}
    for key, value in raw.items():
        normalized[str(key)] = _normalize_value_list(value)
    return normalized


def _normalize_value_list(raw: Any) -> tuple[str, ...]:
    if isinstance(raw, str):
        return (raw,)
    if isinstance(raw, Iterable):
        return tuple(str(item) for item in raw)
    return ()
