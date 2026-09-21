from __future__ import annotations

import os
import sys
import time

import pytest

from docwen_core.cancellation import CancellationToken
from docwen_core.mermaid_render import MermaidRenderError, _run_renderer

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


@pytest.mark.parametrize("cancel", [False, True])
def test_timeout_and_cancellation_reap_real_child_processes(tmp_path, cancel) -> None:
    script = tmp_path / "worker.py"
    script.write_text(
        "import sys, subprocess, time\n"
        "from pathlib import Path\n"
        "sys.stdin.readline()\n"
        "child = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(60)'])\n"
        "Path('child.pid').write_text(str(child.pid))\n"
        "time.sleep(60)\n"
    )
    token = CancellationToken()

    class CancelAfterChildStarts:
        def check(self):
            if (tmp_path / "child.pid").exists():
                token.cancel()
            token.check()

    started = time.monotonic()
    with pytest.raises(Exception if cancel else MermaidRenderError, match=r"[Cc]ancel|exceeded"):
        _run_renderer([sys.executable, "-B", str(script)], tmp_path, CancelAfterChildStarts() if cancel else None, 2)
    assert time.monotonic() - started < 8
    pid = int((tmp_path / "child.pid").read_text())
    deadline = time.monotonic() + 5
    while time.monotonic() < deadline:
        if os.name == "nt":
            import win32api
            import win32con
            import win32event

            try:
                handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
            except Exception:
                return
            try:
                if win32event.WaitForSingleObject(handle, 0) == win32event.WAIT_OBJECT_0:
                    return
            finally:
                win32api.CloseHandle(handle)
        else:
            try:
                os.kill(pid, 0)
            except ProcessLookupError:
                return
            status = __import__("pathlib").Path(f"/proc/{pid}/stat")
            if status.exists() and status.read_text().split()[2] == "Z":
                return
        time.sleep(0.05)
    pytest.fail("Mermaid child survived cancellation/timeout")
