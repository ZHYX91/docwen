"""The language option describes its bounded effects in real CLI help."""

from __future__ import annotations

import os
import subprocess
import sys
import tomllib
from argparse import Namespace
from pathlib import Path

import pytest

from docwen_cli import i18n
from docwen_cli.commands.execution_options import build_execution_options
from docwen_cli.main import main
from docwen_cli.parser import get_available_locale_codes
from docwen_core.text import LOCALE_TO_OCR_LANGUAGE, resolve_ocr_language

pytestmark = pytest.mark.unit
ROOT = Path(__file__).resolve().parents[4]


def _catalog_help(locale: str) -> str:
    with (ROOT / "i18n" / "locales" / f"{locale}.toml").open("rb") as stream:
        return tomllib.load(stream)["cli"]["help"]["lang"]


@pytest.fixture(autouse=True)
def _restore_locale():
    previous = i18n.get_cli_locale()
    yield
    i18n.init_cli_locale(previous)


@pytest.mark.parametrize("locale", get_available_locale_codes())
def test_language_help_has_matching_runtime_and_standalone_copy(locale, monkeypatch):
    expected = _catalog_help(locale)
    i18n.init_cli_locale(locale)
    assert i18n.cli_t("cli.help.lang") == expected
    monkeypatch.setattr(i18n, "_get_runtime_i18n", lambda: None)
    assert i18n.cli_t("cli.help.lang") == expected
    assert "OCR" in expected


@pytest.mark.parametrize("locale", get_available_locale_codes())
@pytest.mark.parametrize("command", [[], ["convert"]], ids=["root", "convert"])
def test_real_cli_help_renders_the_requested_language(locale, command, tmp_path):
    # Use the actual entry point in a fresh interpreter, loading this checkout.
    sources = sorted(str(path) for path in (ROOT / "packages").rglob("src") if path.is_dir())
    env = {
        **os.environ,
        "PYTHONPATH": os.pathsep.join([str(ROOT), *sources]),
        "PYTHONIOENCODING": "utf-8",
        "DOCWEN_DATA_DIR": str(tmp_path / "profile"),
    }
    result = subprocess.run(
        [
            sys.executable,
            "-c",
            "from docwen_cli.main import main; raise SystemExit(main())",
            *command,
            "--lang",
            locale,
            "--help",
        ],
        cwd=ROOT,
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=30,
        check=False,
    )
    assert result.returncode == 0, result.stderr
    assert " ".join(_catalog_help(locale).split()) in " ".join(result.stdout.split())
    assert "Override interface language" not in result.stdout


def test_cached_parent_help_follows_repeated_entry_point_language_changes(capsys):
    for locale in ("en_US", "zh_CN", "ar_SA", "en_US"):
        assert main(["convert", "--lang", locale, "--help"]) == 0
        output = capsys.readouterr().out
        assert " ".join(_catalog_help(locale).split()) in " ".join(output.split())


@pytest.mark.parametrize("locale", get_available_locale_codes())
def test_language_option_reaches_supported_tasks_and_only_influences_auto_ocr(locale):
    args = Namespace(command="convert", lang=locale)
    aware = build_execution_options(args, route_options={"locale", "ocr_language"})
    unaware = build_execution_options(args, route_options={"ocr_language"})
    assert aware["locale"] == locale
    assert "locale" not in unaware
    assert "ocr_language" not in aware
    assert "ocr_language" not in unaware

    assert resolve_ocr_language("auto", aware["locale"]) == LOCALE_TO_OCR_LANGUAGE.get(locale, "chinese")
    assert resolve_ocr_language("english", aware["locale"]) == "english"
