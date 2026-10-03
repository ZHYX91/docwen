"""Regression guards for reviewed Brazilian Portuguese UI copy."""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

_LOCALE = Path(__file__).resolve().parents[2] / "i18n" / "locales" / "pt_BR.toml"


def _value(section: str, key: str) -> str:
    data = tomllib.loads(_LOCALE.read_text(encoding="utf-8"))
    node = data
    for part in section.split("."):
        node = node[part]
    value = node[key]
    assert isinstance(value, str)
    return value


@pytest.mark.parametrize(
    ("section", "key", "expected"),
    [
        ("main_window", "settings_tooltip", "Configurações (Ctrl+,)"),
        ("main_window", "task_completed_status", "Concluído"),
        ("settings.reset", "success", "Todas as configurações foram redefinidas para o padrão"),
        ("settings.toml_editor", "config_label", "Configuração"),
        ("settings.formatting", "heading_merge_mode_label", "Modo de fusão entre título e corpo:"),
        ("components.file_drop", "unsupported_type_msg", "Tipo de arquivo não suportado: {filename}"),
    ],
)
def test_reviewed_portuguese_copy_keeps_native_diacritics(
    section: str,
    key: str,
    expected: str,
) -> None:
    assert _value(section, key) == expected
