from __future__ import annotations

import os
import sys
import time

import pytest

from docwen_core.cancellation import CancellationToken
from docwen_core.errors import CancellationRequested
from docwen_core.mermaid_render import MermaidRenderError, _renderer_directory, _run_renderer

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


@pytest.mark.parametrize("winerror", [32, 33, 5, None])
def test_profile_cleanup_preserves_cancellation_and_reports_nonsharing_errors(tmp_path, monkeypatch, winerror) -> None:
    from docwen_core import mermaid_render

    original_cleanup = mermaid_render.shutil.rmtree
    attempts = []
    error = PermissionError("profile in use")
    monkeypatch.setattr(error, "winerror", winerror, raising=False)

    def cleanup(temporary):
        attempts.append(temporary)
        if len(attempts) == 1:
            raise error
        original_cleanup(temporary)

    monkeypatch.setattr(mermaid_render.shutil, "rmtree", cleanup)
    expected = CancellationRequested
    with pytest.raises(expected), _renderer_directory(tmp_path) as directory:
        (directory / "profile").write_text("owned browser data")
        raise CancellationRequested("cancelled")
    if winerror in (32, 33):
        assert len(attempts) == 2
        assert not directory.exists()
    else:
        assert len(attempts) == 1


def test_profile_cleanup_sharing_retry_has_a_deadline(tmp_path, monkeypatch) -> None:
    from docwen_core import mermaid_render

    error = PermissionError("profile remains in use")
    monkeypatch.setattr(error, "winerror", 32, raising=False)

    def cleanup(temporary):
        raise error

    clock = iter([0, 0, 6])
    monkeypatch.setattr(mermaid_render.shutil, "rmtree", cleanup)
    monkeypatch.setattr(mermaid_render.time, "monotonic", lambda: next(clock))
    monkeypatch.setattr(mermaid_render.time, "sleep", lambda _: None)
    with pytest.raises(PermissionError, match="remains in use"), _renderer_directory(tmp_path):
        pass


@pytest.mark.parametrize("cancel", [False, True])
@pytest.mark.skipif(sys.platform == "darwin", reason="macOS interrupted process-family exit remains unconfirmed")
def test_timeout_and_cancellation_reap_real_child_processes(tmp_path, cancel) -> None:
    token = CancellationToken()
    children = []
    ready_at = []

    class CancelAfterChildStarts:
        def check(self):
            if (job_dir / "child.ready").exists() and (job_dir / "child.pid").exists():
                if not children:
                    children.append(int((job_dir / "child.pid").read_text()))
                    ready_at.append(time.monotonic())
                if cancel:
                    token.cancel()
            token.check()

    started = time.monotonic()
    with (
        pytest.raises(CancellationRequested if cancel else MermaidRenderError, match=r"[Cc]ancel|exceeded|timed out"),
        _renderer_directory(tmp_path) as job_dir,
    ):
        script = job_dir / "worker.py"
        script.write_text(
            "import sys, subprocess, time\n"
            "from pathlib import Path\n"
            "sys.stdin.readline()\n"
            "child = subprocess.Popen([sys.executable, '-c', "
            "\"import time; from pathlib import Path; handle=open('child.lock', 'wb'); "
            "Path('child.ready').touch(); time.sleep(60)\"])\n"
            "Path('child.pid').write_text(str(child.pid))\n"
            "time.sleep(60)\n"
        )
        _run_renderer([sys.executable, "-B", str(script)], job_dir, CancelAfterChildStarts(), 15 if cancel else 5)
    assert children, "The test must exercise a running child, not only terminate its parent"
    assert time.monotonic() - (ready_at[0] if cancel else started) < 12
    assert not job_dir.exists(), "The production cleanup boundary must remove the locked browser profile"
    pid = children[0]
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


@pytest.mark.parametrize("error", [CancellationRequested("cancelled"), KeyboardInterrupt()])
def test_profile_cleanup_cannot_replace_interruption_with_ordinary_error(tmp_path, monkeypatch, error):
    from docwen_core import mermaid_render

    def fail(_path):
        raise OSError(39, "Directory not empty")

    monkeypatch.setattr(mermaid_render.shutil, "rmtree", fail)
    with pytest.raises(type(error)) as caught, _renderer_directory(tmp_path):
        raise error
    assert caught.value is error


def test_unconfirmed_process_directory_has_no_automatic_finalizer(tmp_path):
    import gc

    from docwen_core.errors import ExternalProcessCleanupUnconfirmed

    with pytest.raises(ExternalProcessCleanupUnconfirmed), _renderer_directory(tmp_path) as directory:
        (directory / "evidence").write_text("still owned")
        raise ExternalProcessCleanupUnconfirmed("unknown")
    gc.collect()
    assert (directory / "evidence").read_text() == "still owned"


@pytest.mark.skipif(sys.platform != "darwin", reason="macOS conservative cleanup contract")
def test_macos_nonzero_launcher_exit_preserves_unconfirmed_directory(tmp_path):
    from docwen_core.errors import ExternalProcessCleanupUnconfirmed

    with pytest.raises(ExternalProcessCleanupUnconfirmed), _renderer_directory(tmp_path) as directory:
        _run_renderer([sys.executable, "-c", "import sys; sys.stdin.readline(); sys.exit(1)"], directory, None, 5)
    assert directory.exists()
