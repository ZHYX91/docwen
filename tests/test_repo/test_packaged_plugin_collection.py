from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def test_build_script_collects_all_default_plugin_packages() -> None:
    from docwen_bundle.runtime_factory import _DEFAULT_PLUGIN_IMPORTS

    source = Path("scripts/build/build.py").read_text(encoding="utf-8")

    for plugin_import in _DEFAULT_PLUGIN_IMPORTS:
        assert source.count(f'"--hidden-import={plugin_import}"') == 2, plugin_import


def test_packaged_cli_verifier_rejects_plugin_load_failures() -> None:
    from scripts.release import verify_packaged_cli

    proc = subprocess.CompletedProcess(
        ["DocWenCLI.exe", "doctor", "--json", "--quiet"],
        0,
        stdout='{"success": true}',
        stderr=(
            "Failed to load plugin docwen_plugin_markup: No module named "
            "'docwen_plugin_markup'. It will not be available."
        ),
    )

    with pytest.raises(RuntimeError, match="unavailable packaged plugins"):
        verify_packaged_cli._load_json_payload(proc, command_name="doctor")
