from __future__ import annotations

import json
import os
from pathlib import Path

import pytest

from docwen_core.mermaid_runtime import inspect_mermaid_runtime

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


def _installation(root: Path, cli_version: str, core_version: str) -> Path:
    cli = root / "node_modules/@mermaid-js/mermaid-cli"
    (cli / "src").mkdir(parents=True)
    (cli / "src/cli.js").write_text("")
    (cli / "package.json").write_text(json.dumps({"name": "@mermaid-js/mermaid-cli", "version": cli_version}))
    core = root / "node_modules/mermaid"
    core.mkdir()
    (core / "package.json").write_text(json.dumps({"name": "mermaid", "version": core_version}))
    shim = root / "mmdc.cmd"
    shim.write_text("")
    (root / ("node.exe" if os.name == "nt" else "node")).write_text("")
    return shim


@pytest.mark.parametrize(
    "cli,core,available",
    [
        ("11.17.0", "11.17.2", True),
        ("11.16.0", "11.16.0", True),
        ("11.16.0", "11.15.0", False),
        ("11.15.0", "11.16.0", False),
        ("12.0.0", "12.0.0", False),
        ("11.17.0", "invalid", False),
    ],
)
def test_discovery_checks_actual_cli_and_core_versions(tmp_path, cli, core, available) -> None:
    runtime = inspect_mermaid_runtime(str(_installation(tmp_path, cli, core)))
    assert runtime.available is available
    assert runtime.cli_version == cli
    assert runtime.mermaid_version == core
    assert runtime.reason == ("ready" if available else "unsupported_version")


def test_explicit_missing_path_never_silently_selects_another_installation(tmp_path, monkeypatch) -> None:
    shim = _installation(tmp_path, "11.17.0", "11.17.2")
    monkeypatch.setenv("DOCWEN_MERMAID_CLI", str(shim))
    assert inspect_mermaid_runtime().available
    assert inspect_mermaid_runtime(str(tmp_path / "missing")).reason == "cli_unavailable"
