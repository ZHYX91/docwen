"""Resolve proofreading switches consistently for actions and composed workflows."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any

from docwen_core.options import PROOFREAD_OPTIONS_SCHEMA


def resolve_proofread_switches(config: Mapping[str, Any], options: Mapping[str, Any]) -> dict[str, bool]:
    proofread = config.get("proofread", {})
    engine = proofread.get("engine", {}) if isinstance(proofread, Mapping) else {}
    if not isinstance(engine, Mapping):
        engine = {}
    defaults = {
        key: bool(engine.get(key, spec["default"]))
        for key, spec in PROOFREAD_OPTIONS_SCHEMA["properties"].items()
        if key.startswith("enable_")
    }
    return {
        "enable_symbol_pairing": bool(options.get("enable_symbol_pairing", defaults["enable_symbol_pairing"])),
        "enable_symbol_correction": bool(options.get("enable_symbol_correction", defaults["enable_symbol_correction"])),
        "enable_typos_rule": bool(options.get("enable_typos_rule", defaults["enable_typos_rule"])),
        "enable_sensitive_word": bool(options.get("enable_sensitive_word", defaults["enable_sensitive_word"])),
    }
