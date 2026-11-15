"""Mermaid selects the existing process-family owner before business launch."""

from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from docwen_core import mermaid_render, office_process
from docwen_core.errors import CancellationRequested, ExternalProcessCleanupUnconfirmed

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("platform", ["linux", "win32"])
def test_renderer_uses_owned_launch_instead_of_attaching_a_running_launcher(tmp_path, monkeypatch, platform):
    monkeypatch.setattr(mermaid_render, "sys", SimpleNamespace(platform=platform))
    monkeypatch.setenv("NODE_OPTIONS", "--require=untrusted.js")
    direct = MagicMock(side_effect=AssertionError("renderer must not start outside its owner"))
    monkeypatch.setattr(mermaid_render.subprocess, "Popen", direct)
    owner = MagicMock(return_value=office_process.OfficeProcessResult(True, True))
    monkeypatch.setattr(office_process, "run_owned_office", owner)

    mermaid_render._run_renderer(["venv-python.exe", "worker.py"], tmp_path, None, 5)

    direct.assert_not_called()
    owner.assert_called_once()
    assert owner.call_args.args == (["venv-python.exe", "worker.py"],)
    assert owner.call_args.kwargs["timeout_s"] == 5
    launch = owner.call_args.kwargs["launch"]
    assert launch.cwd == str(tmp_path)
    assert launch.stdin_bytes == b"go\n"
    assert launch.log_path == str(tmp_path / "renderer.log")
    assert launch.log_limit == 1024 * 1024
    assert launch.preserve_confirmed_interruption
    environment = dict(launch.environment)
    assert "NODE_OPTIONS" not in environment
    assert {environment[key] for key in ("TMP", "TEMP", "TMPDIR")} == {str(tmp_path)}


@pytest.mark.parametrize("error", [CancellationRequested("cancelled"), KeyboardInterrupt()])
def test_windows_renderer_preserves_confirmed_interruption(tmp_path, monkeypatch, error):
    monkeypatch.setattr(mermaid_render, "sys", SimpleNamespace(platform="win32"))
    monkeypatch.setattr(office_process, "run_owned_office", MagicMock(side_effect=error))
    with pytest.raises(type(error)) as caught:
        mermaid_render._run_renderer(["renderer.exe"], tmp_path, None, 5)
    assert caught.value is error


def test_windows_renderer_retains_workspace_when_owner_startup_is_unconfirmed(tmp_path, monkeypatch):
    monkeypatch.setattr(mermaid_render, "sys", SimpleNamespace(platform="win32"))
    result = office_process.OfficeProcessResult(False, False, "owner startup not confirmed")
    monkeypatch.setattr(office_process, "run_owned_office", MagicMock(return_value=result))
    with pytest.raises(ExternalProcessCleanupUnconfirmed), mermaid_render._renderer_directory(tmp_path) as directory:
        (directory / "evidence").write_text("keep")
        mermaid_render._run_renderer(["renderer.exe"], directory, None, 5)
    assert (directory / "evidence").read_text() == "keep"
