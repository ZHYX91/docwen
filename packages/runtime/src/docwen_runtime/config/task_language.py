"""Task output language is independent of every application's display language."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any

DEFAULT_OUTPUT_LOCALE = "zh_CN"


def configured_output_locale(snapshot: Mapping[str, Any]) -> str:
    conversion = snapshot.get("conversion")
    value = conversion.get("output_locale") if isinstance(conversion, Mapping) else None
    return value.strip() if isinstance(value, str) and value.strip() else DEFAULT_OUTPUT_LOCALE


def output_yaml_labels(locale: str) -> dict[str, str]:
    from docwen_core.docx_styles import SHIPPED_STYLE_LOCALES
    from docwen_runtime.i18n import load_locale_table
    from docwen_runtime.resources import ResourceRegistry

    if locale not in SHIPPED_STYLE_LOCALES:
        return {"title": "title", "subtitle": "subtitle"}
    table = load_locale_table(ResourceRegistry.default().locales_dir() / f"{locale}.toml")
    keys = table.get("yaml_keys", {})
    return {key: str(keys.get(key) or key) if isinstance(keys, Mapping) else key for key in ("title", "subtitle")}
