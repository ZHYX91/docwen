"""Native Linux ownership for detached browser-shaped children, without a browser."""

from __future__ import annotations

import multiprocessing
import os
import signal
import subprocess
import sys
from pathlib import Path
from typing import Any, cast

import pytest

from docwen_core import office_process
from docwen_core.errors import CancellationRequested
from docwen_core.mermaid_render import _run_renderer

pytestmark = [
    pytest.mark.integration,
    pytest.mark.pr_gate,
    pytest.mark.skipif(sys.platform != "linux", reason="Linux subreaper/pidfd"),
]


def _detached_script(root: Path) -> Path:
    script = root / "renderer.py"
    script.write_text(
        "import subprocess, sys, os\n"
        "sys.stdin.readline()\n"
        "subprocess.Popen([sys.executable, '-c', "
        '"import signal,time,os; from pathlib import Path; '
        "signal.signal(signal.SIGTERM, signal.SIG_IGN); "
        "signal.signal(signal.SIGINT, signal.SIG_IGN); "
        "Path('child.pid').write_text(str(os.getpid())); time.sleep(60)\"], start_new_session=True)\n"
        "os._exit(0)\n"
    )
    return script


@pytest.mark.parametrize("kind", ["timeout", "cancel", "interrupt"])
def test_launcher_exit_does_not_release_detached_descendant(tmp_path, kind):
    script = _detached_script(tmp_path)
    error = KeyboardInterrupt() if kind == "interrupt" else CancellationRequested("cancelled")

    class Cancellation:
        def check(self):
            if kind != "timeout" and (tmp_path / "child.pid").exists():
                raise error

    from docwen_core.mermaid_render import MermaidRenderError

    expected = MermaidRenderError if kind == "timeout" else type(error)
    with pytest.raises(expected) as caught:
        _run_renderer([sys.executable, str(script)], tmp_path, Cancellation(), 1 if kind == "timeout" else 10)
    if kind != "timeout":
        assert caught.value is error
    pid = int((tmp_path / "child.pid").read_text())
    with pytest.raises(ProcessLookupError):
        os.kill(pid, 0)  # ECHILD proof includes reaping, not merely zombie state.


def _handshake_owner(connection, stop, cmd, timeout, launch):
    original = subprocess.Popen

    class BrokenInput:
        def write(self, _data):
            raise BrokenPipeError("injected handshake failure")

        def close(self):
            pass

    def start(*args, **kwargs):
        process = original(*args, **kwargs)
        assert process.stdin is not None
        process.stdin.close()
        cast(Any, process).stdin = BrokenInput()
        return process

    subprocess.Popen = start
    office_process._owner_worker(connection, stop, cmd, timeout, launch)


def test_handshake_failure_after_launch_still_reaches_reap_proof(tmp_path, monkeypatch):
    def start(cmd, timeout, launch):
        context = multiprocessing.get_context("spawn")
        parent, child = context.Pipe(duplex=False)
        stop = context.Event()
        owner = context.Process(target=_handshake_owner, args=(child, stop, cmd, timeout, launch))
        owner.start()
        child.close()
        return owner, parent, stop

    monkeypatch.setattr(office_process, "_start_owner", start)
    result = office_process.run_owned_office(
        [sys.executable, "-c", "import time; time.sleep(60)"],
        is_cancelled=lambda: False,
        timeout_s=10,
        launch=office_process.OwnedProcessLaunch(
            cwd=str(tmp_path),
            stdin_bytes=b"go\n",
            log_path=str(tmp_path / "renderer.log"),
            grace_signal=signal.SIGINT,
        ),
    )
    assert not result.succeeded
    assert result.cleanup_confirmed
    assert result.launcher_pid in result.reaped_pids
    assert "handshake failed" in result.reason
