"""YAML front matter extraction for MD→DOCX conversion."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from typing import Any

from docwen_core.yaml_tools import extract_yaml_properties

extract_yaml_front_matter = extract_yaml_properties

# Bundled template placeholder contract.  These values are the title/body
# translations used by the 12 bundled DOCX templates.
# Keeping the contract at the plugin boundary avoids importing the
# runtime i18n service from conversion code.
TITLE_PLACEHOLDER_ALIASES: frozenset[str] = frozenset(
    {
        "Titel",
        "title",
        "título",
        "Titre",
        "タイトル",
        "제목",
        "Título",
        "Заголовок",
        "Tiêu đề",
        "标题",
        "標題",
        "العنوان",
    }
)
BODY_PLACEHOLDER_ALIASES: frozenset[str] = frozenset(
    {
        "Inhalt",
        "body",
        "cuerpo",
        "Corps",
        "本文",
        "본문",
        "Corpo",
        "Текст",
        "Nội dung",
        "正文",
        "المتن",
    }
)


def is_empty_field_value(value: Any) -> bool:
    """Treat empty nested YAML containers as empty without discarding zero/false."""
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    if isinstance(value, Mapping):
        return all(is_empty_field_value(item) for item in value.values())
    if isinstance(value, (list, tuple, set)):
        return all(is_empty_field_value(item) for item in value)
    return False


def ensure_title_fallback(
    yaml_data: dict[str, Any],
    *,
    placeholder_names: Iterable[str],
    source_stem: str,
) -> None:
    """Apply the maintained title fallback for active templates.

    Only title placeholders actually present in the selected template are
    populated.  An existing localized title wins; otherwise the first
    non-empty ``aliases`` item is used, followed by the input filename stem.
    """
    fallback = _title_fallback_value(yaml_data, source_stem)
    if not fallback:
        return
    for name in placeholder_names:
        if name not in TITLE_PLACEHOLDER_ALIASES:
            continue
        value = yaml_data.get(name)
        if value is None or (isinstance(value, str) and not value.strip()):
            yaml_data[name] = fallback


def _title_fallback_value(yaml_data: dict[str, Any], source_stem: str) -> str:
    aliases = yaml_data.get("aliases")
    if isinstance(aliases, str) and aliases.strip():
        return aliases.strip()
    if isinstance(aliases, (list, tuple)):
        for item in aliases:
            text = str(item).strip() if item is not None else ""
            if text:
                return text
    return str(source_stem or "").strip()
