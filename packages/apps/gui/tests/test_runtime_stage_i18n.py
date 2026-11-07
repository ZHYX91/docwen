"""All shipped locales must explain preparation, cancellation and workdir errors."""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    "locale",
    ["zh_CN", "en_US", "zh_TW", "de_DE", "es_ES", "fr_FR", "ja_JP", "ko_KR", "pt_BR", "ru_RU", "vi_VN", "ar_SA"],
)
def test_runtime_progress_and_folder_help_keys_are_translated(locale: str) -> None:
    repo = Path(__file__).resolve().parents[4]
    catalog = tomllib.loads((repo / "i18n" / "locales" / f"{locale}.toml").read_text(encoding="utf-8"))
    for key in ("runtime_folder_unavailable", "output_folder_unavailable"):
        assert isinstance(catalog["main_window"][key], str) and catalog["main_window"][key].strip()
    for key in ("phase_checking_output", "phase_preparing_workspace", "phase_converting", "phase_cancelling_hint"):
        assert isinstance(catalog["info_area"][key], str) and catalog["info_area"][key].strip()
