"""Shared accessors for numbering schemes used by GUI surfaces."""

from __future__ import annotations

from collections.abc import Mapping


def get_numbering_scheme_items(
    locale: str | None = None,
    *,
    config_data: object,
) -> list[tuple[str, str]]:
    """Project ``(display_name, scheme_id)`` from an injected config value."""
    if not isinstance(config_data, Mapping):
        return []

    from docwen_gui.i18n import get_locale, t
    from docwen_runtime.i18n import locale_is_eligible

    active_locale = str(locale or get_locale())

    raw_settings = config_data.get("settings", {})
    raw_schemes = config_data.get("schemes", {})
    settings = raw_settings if isinstance(raw_settings, Mapping) else {}
    schemes = raw_schemes if isinstance(raw_schemes, Mapping) else {}
    raw_order = settings.get("order", [])
    order = [str(item) for item in raw_order] if isinstance(raw_order, (list, tuple)) else []
    scheme_ids = [*order, *(str(key) for key in schemes if str(key) not in order)]
    items: list[tuple[str, str]] = []
    for scheme_id in scheme_ids:
        raw = schemes.get(scheme_id)
        if not isinstance(raw, Mapping):
            continue
        if not locale_is_eligible(raw.get("locales", ["*"]), active_locale):
            continue
        name = str(raw.get("name") or "")
        name_key = str(raw.get("name_key") or "")
        if not name and name_key:
            name = t(f"editors.numbering_add.names.{name_key}", name_key)
        items.append((name or scheme_id, scheme_id))
    return items
