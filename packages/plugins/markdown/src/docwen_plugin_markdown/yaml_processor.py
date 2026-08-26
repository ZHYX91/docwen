"""YAML front matter extraction for MD→DOCX conversion."""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

import yaml

from docwen_core.links import split_yaml_front_matter_source

# Bundled template placeholder contract.  These values are the title/body
# translations used by the 11 DOCX templates shared with both reference
# projects.  Keeping the contract at the plugin boundary avoids importing the
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
    }
)


def extract_yaml_front_matter(content: str) -> tuple[dict[str, Any], str]:
    """Extract YAML front matter from markdown content.

    Args:
        content: Raw markdown text, possibly starting with ``---\\n``.

    Returns:
        ``(yaml_dict, remaining_body)``. If no front matter is found or
        parsing fails, returns ``({}, content)``.
    """
    yaml_front, body = split_yaml_front_matter_source(content)
    if not yaml_front:
        return {}, content
    front_lines = yaml_front.removeprefix("\ufeff").splitlines()
    yaml_text = "\n".join(front_lines[1:-1])
    try:
        yaml_data = yaml.safe_load(yaml_text)
        if not isinstance(yaml_data, dict):
            return {}, content
        return yaml_data, body
    except yaml.YAMLError:
        return {}, content


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
