"""Real short-TTL process families, plus cross-platform retention contracts."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock
from urllib.parse import unquote, urlparse

import pytest

from docwen_core import office_bridge, office_process
from docwen_core import office_native_process as native
from docwen_core.office_process import OfficeProcessResult
from docwen_runtime.workspace.manager import WorkspaceManager

from . import _office_native_lifecycle_support as support

pytestmark = [pytest.mark.integration, pytest.mark.office, pytest.mark.pr_gate, pytest.mark.release_gate]


@pytest.mark.parametrize("mode", ["normal", "large-output", "cancel", "timeout"])
def test_current_host_owner_waits_for_child_after_launcher_exit(tmp_path: Path, mode: str) -> None:
    evidence = tmp_path / "family"
    started = time.monotonic()
    result = office_process.run_owned_office(
        support.family_command(evidence, mode),
        is_cancelled=lambda: mode == "cancel" and evidence.with_suffix(".child").exists(),
        timeout_s=1.0 if mode == "timeout" else 5,
    )
    assert result.cleanup_confirmed, result
    assert result.succeeded is (mode in {"normal", "large-output"})
    assert result.owner_exit_code == 0 and result.launcher_exit_code == 0
    assert evidence.with_suffix(".child").exists()
    assert time.monotonic() - started < 7
    previous = evidence.with_suffix(".heartbeat").read_bytes()
    time.sleep(0.04)
    assert evidence.with_suffix(".heartbeat").read_bytes() == previous
    if result.succeeded:
        assert evidence.with_suffix(".finished").exists()
    else:
        # A child's natural TTL expiry must not conceal a missing native kill.
        assert not evidence.with_suffix(".finished").exists()


@pytest.mark.skipif(
    sys.platform != "linux", reason="controlled POSIX group probe uses Linux process queries as Darwin API substitutes"
)
@pytest.mark.parametrize("mode", ["normal", "large-output", "cancel", "timeout", "kill-refused"])
def test_native_group_closeout_with_real_processes_and_substituted_darwin_queries(tmp_path: Path, mode: str) -> None:
    evidence = tmp_path / "family"
    command = [sys.executable, str(Path(support.__file__)), "portable-group-probe", str(evidence), mode]
    with subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE) as process:
        # Observe the production result while a refused child can still run;
        # helper rescue/reap happens only after writing that result.
        deadline = time.monotonic() + 7
        while not evidence.with_suffix(".result").exists():
            if process.poll() is not None or time.monotonic() >= deadline:
                break
            time.sleep(0.01)
        if mode == "kill-refused" and evidence.with_suffix(".result").exists():
            heartbeat = evidence.with_suffix(".heartbeat").read_bytes()
            time.sleep(0.04)
            assert evidence.with_suffix(".heartbeat").read_bytes() != heartbeat
        stdout, stderr = process.communicate(timeout=4)
    assert process.returncode == 0, (stdout, stderr)
    result = json.loads(evidence.with_suffix(".result").read_text())
    assert result["cleanup_confirmed"] is (mode != "kill-refused"), result
    assert result["succeeded"] is (mode in {"normal", "large-output"})
    assert result["launcher_exit_code"] == (None if mode == "kill-refused" else 0)
    if mode == "kill-refused":
        assert "could not be confirmed" in result["reason"]


@pytest.mark.parametrize("platform", ["win32", "darwin", "linux"])
@pytest.mark.parametrize("cancelled", [False, True])
def test_unconfirmed_family_preserves_every_private_tree_after_task_cleanup(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, platform: str, cancelled: bool
) -> None:
    # The process result is injected, so this is a lifecycle/propagation check,
    # not Windows/macOS native acceptance. QA temp is outside the task subtree.
    temp_root = tmp_path / "governed-temp"
    temp_root.mkdir()
    monkeypatch.setattr(office_bridge.tempfile, "tempdir", str(temp_root))
    monkeypatch.setattr(office_bridge.sys, "platform", platform)
    source = tmp_path / "source.docx"
    source.write_bytes(b"source bytes")
    manager = WorkspaceManager(temp_root / "runtime-tasks")
    handle = manager.create("owned-test-task", str(source))
    requested = Path(handle.staging_dir) / "output.pdf"
    requested.write_bytes(b"previous output")
    observed: dict[str, Path] = {}

    class Cancellation:
        is_cancelled = False

    cancellation = Cancellation()

    def unconfirmed(command: list[str], **_kwargs: object) -> office_process.OfficeProcessResult:
        raw = unquote(urlparse(command[1].partition("=")[2]).path)
        if os.name == "nt" and raw.startswith("/") and len(raw) > 2 and raw[2] == ":":
            raw = raw[1:]
        observed["profile"] = Path(raw)
        observed["input"] = Path(command[-1])
        observed["output"] = Path(command[command.index("--outdir") + 1])
        (observed["output"] / "input.pdf").write_bytes(b"partial output")
        cancellation.is_cancelled = cancelled
        return office_process.OfficeProcessResult(False, False, "owned child exit unconfirmed")

    monkeypatch.setattr(office_bridge, "find_soffice_path", lambda: "owned-office")
    monkeypatch.setattr(office_bridge, "_run_libreoffice_process", unconfirmed)
    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(requested),
        source_format="docx",
        backend_priority=["libreoffice", "must-not-run"],
        com_candidates={},
        libreoffice_format="pdf",
        cancel=cancellation,
    )
    assert not result.success and result.cleanup_failed
    assert result.cancelled is cancelled
    assert result.error_code == "OFFICE_PROCESS_CLEANUP_UNCONFIRMED"
    assert result.attempted_backend_ids == ("libreoffice",)
    assert requested.read_bytes() == b"previous output"
    assert not observed["input"].is_relative_to(Path(handle.staging_dir).parent)
    manager.cleanup("owned-test-task")
    assert not Path(handle.staging_dir).exists()
    assert observed["profile"].is_dir() and observed["output"].is_dir()
    assert observed["input"].read_bytes() == b"source bytes"
    assert (observed["output"] / "input.pdf").read_bytes() == b"partial output"
    assert str(observed["profile"]) in result.cleanup_message
    assert str(observed["input"].parent) in result.cleanup_message
    shutil.rmtree(observed["profile"])
    shutil.rmtree(observed["input"].parent)


def test_native_owner_uses_file_streams_and_reports_before_supervisor_join(monkeypatch: pytest.MonkeyPatch) -> None:
    connection, family = MagicMock(), MagicMock()
    family.pid = 123
    streams: list[int] = []

    def start(_command: list[str], **kwargs: object) -> None:
        stdin, stdout = kwargs["stdin"], kwargs["stdout"]
        assert isinstance(stdin, int) and isinstance(stdout, int)
        streams.extend((stdin, stdout))
        assert os.read(stdin, 1) == b""
        assert os.write(stdout, b"x" * (4 * 1024 * 1024)) == 4 * 1024 * 1024

    family.start.side_effect = start
    family.exited.return_value = True
    family.reap.return_value = 0
    monkeypatch.setattr(native, "_native_family", lambda: family)
    monkeypatch.setattr(native.signal, "signal", MagicMock())
    native.native_owner_worker(connection, SimpleNamespace(is_set=lambda: False), ["owned-office"], 1)
    result = connection.send.call_args.args[0]
    assert isinstance(result, OfficeProcessResult)
    assert result.succeeded and result.cleanup_confirmed
    assert result.launcher_pid == 123
    for descriptor in streams:
        with pytest.raises(OSError):
            os.fstat(descriptor)
    family.close.assert_called_once()
    connection.close.assert_called_once()


def test_owner_close_failure_is_not_a_cleanup_receipt(monkeypatch: pytest.MonkeyPatch) -> None:
    connection, family = MagicMock(), MagicMock()
    family.pid = 123
    family.exited.return_value = True
    family.reap.return_value = 0
    family.close.side_effect = OSError("cannot close native handles")
    monkeypatch.setattr(native, "_native_family", lambda: family)
    monkeypatch.setattr(native.signal, "signal", MagicMock())
    native.native_owner_worker(connection, SimpleNamespace(is_set=lambda: False), ["owned-office"], 1)
    result = connection.send.call_args.args[0]
    assert not result.cleanup_confirmed and not result.succeeded
    assert "closeout" in result.reason
