"""Test that all docwen_cli modules are importable.

This enforces the import hygiene rule: CLI modules must not
deep-import plugin or runtime internals.
"""

import pytest

pytestmark = pytest.mark.unit


def test_cli_importable() -> None:
    """docwen_cli should be importable."""
    import docwen_cli

    assert docwen_cli.__version__ == "0.9.1"


def test_exit_codes_importable() -> None:
    from docwen_cli.exit_codes import ExitCode

    assert ExitCode.OK == 0
    assert ExitCode.INTERNAL_ERROR == 1
    assert ExitCode.INVALID_INPUT == 2
    assert ExitCode.CANCELLED == 130


def test_parser_importable() -> None:
    from docwen_cli.parser import get_available_locale_codes, get_common_parser

    assert callable(get_common_parser)
    codes = get_available_locale_codes()
    assert "zh_CN" in codes
    assert "en_US" in codes


def test_utils_importable() -> None:
    from docwen_cli.utils import expand_paths, validate_files

    assert callable(expand_paths)
    assert callable(validate_files)


def test_i18n_importable() -> None:
    from docwen_cli.i18n import cli_t, init_cli_locale

    assert callable(cli_t)
    assert callable(init_cli_locale)


def test_i18n_docstring_has_no_stale_phase_marker() -> None:
    import inspect

    from docwen_cli import i18n

    source = inspect.getsource(i18n)
    assert "Phase 6" not in source
    assert "CLI i18n adapter with inline fallbacks and an optional runtime bridge." in source


def test_presenters_importable() -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter
    from docwen_cli.presenters.text_presenter import TextPresenter

    assert TextPresenter
    assert JsonPresenter


def test_commands_importable() -> None:
    """All command modules should be importable."""
    # All imports succeeded
    assert True


def test_main_importable() -> None:
    from docwen_cli.main import (
        main,
        pre_parse_lang,
    )

    assert callable(main)
    assert callable(pre_parse_lang)


def test_cli_does_not_import_runtime() -> None:
    """CLI package must not directly import docwen_runtime.

    Checks each loaded ``docwen_cli.*`` submodule for references to
    ``docwen_runtime``, rather than global ``sys.modules`` which may
    be polluted by other test packages (bundle, IPC, etc.).
    """
    import sys

    # Collect all docwen_cli submodules currently loaded.
    cli_modules = sorted(k for k in sys.modules if k == "docwen_cli" or k.startswith("docwen_cli."))
    assert cli_modules, "No docwen_cli modules loaded — test ordering issue?"

    offenders: dict[str, list[str]] = {}
    for mod_name in cli_modules:
        mod = sys.modules[mod_name]
        refs = [
            v.__name__ if hasattr(v, "__name__") else str(v)
            for v in vars(mod).values()
            if hasattr(v, "__name__")
            and (
                v.__name__ == "docwen_runtime"
                or (isinstance(v.__name__, str) and v.__name__.startswith("docwen_runtime."))
            )
        ]
        if refs:
            offenders[mod_name] = refs

    assert not offenders, f"docwen_cli must not import docwen_runtime directly. Offending modules: {offenders}"


def test_cli_does_not_import_plugins() -> None:
    """CLI package must not import any docwen_plugin_*.

    Checks each loaded ``docwen_cli.*`` submodule for references to
    ``docwen_plugin_*``, rather than global ``sys.modules``.
    """
    import sys

    cli_modules = sorted(k for k in sys.modules if k == "docwen_cli" or k.startswith("docwen_cli."))
    assert cli_modules, "No docwen_cli modules loaded — test ordering issue?"

    offenders: dict[str, list[str]] = {}
    for mod_name in cli_modules:
        mod = sys.modules[mod_name]
        refs = [
            v.__name__ if hasattr(v, "__name__") else str(v)
            for v in vars(mod).values()
            if hasattr(v, "__name__") and (isinstance(v.__name__, str) and v.__name__.startswith("docwen_plugin_"))
        ]
        if refs:
            offenders[mod_name] = refs

    assert not offenders, f"docwen_cli must not import plugin internals. Offending modules: {offenders}"
