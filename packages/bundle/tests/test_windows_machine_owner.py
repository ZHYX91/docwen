"""Real Windows Job lifetime tests for the installed-CLI supervision boundary."""

from __future__ import annotations

import io
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import pytest

pytestmark = [pytest.mark.integration, pytest.mark.skipif(sys.platform != "win32", reason="Windows Job API")]

# A Windows venv python.exe can itself be a waiting redirector that retains
# stdio. Use the actual interpreter for synthetic workers and supervisors so
# the test observes exactly the two processes in the installed CLI contract.
PYTHON = getattr(sys, "_base_executable", sys.executable)


def _wait_json(path: Path) -> dict:
    deadline = time.monotonic() + 15
    while time.monotonic() < deadline:
        if path.exists():
            try:
                return json.loads(path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                pass
        time.sleep(0.02)
    raise AssertionError(f"Worker did not become ready: {path.name}")


def _worker(root: Path) -> list[str]:
    # The grandchild deliberately owns no stdio pipe, so an inherited-pipe
    # close event alone would not prove that it has been cleaned up.
    code = (
        "import json,os,pathlib,subprocess,sys,time; "
        "p=subprocess.Popen([sys.executable,'-c','import time;time.sleep(120)'],"
        "stdin=subprocess.DEVNULL,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL); "
        f"pathlib.Path({str(root / 'ready.json')!r}).write_text(json.dumps({{'worker':os.getpid(),'child':p.pid}})); "
        "sys.stdin.buffer.read(1)"
    )
    return [PYTHON, "-c", code]


def _launch(command: list[str], parent_pid: int) -> subprocess.Popen:
    code = (
        "from docwen_bundle.windows_machine_owner import run_job_worker; "
        f"raise SystemExit(run_job_worker({command!r},parent_pid={parent_pid}))"
    )
    return subprocess.Popen([PYTHON, "-c", code], stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)


def _open_process(pid: int):
    import win32api

    return win32api.OpenProcess(0x100000, False, pid)


@pytest.mark.parametrize("finish", ["eof", "worker_exit", "kill_owner"])
def test_cli_owner_cleans_detached_stdio_descendants(tmp_path: Path, finish: str) -> None:
    import win32api
    import win32event

    control = _launch(_worker(tmp_path), os.getpid())
    handles = []
    sentinel = subprocess.Popen([sys.executable, "-c", "import time;time.sleep(120)"])
    try:
        ready = _wait_json(tmp_path / "ready.json")
        handles = [_open_process(ready[key]) for key in ("worker", "child")]
        assert control.stdin is not None
        if finish == "kill_owner":
            control.kill()
        elif finish == "worker_exit":
            control.stdin.write(b"x")
            control.stdin.flush()
        else:
            control.stdin.close()
        control.wait(timeout=15)
        assert control.stderr is not None
        assert b"DOCWEN_PROCESS_OWNER_READY:1" in control.stderr.read()
        if finish != "kill_owner":
            assert control.returncode == 0
        for handle in handles:
            assert win32event.WaitForSingleObject(handle, 3000) == win32event.WAIT_OBJECT_0
        assert sentinel.poll() is None
    finally:
        if control.poll() is None:
            control.kill()
        control.wait(timeout=5)
        for stream in (control.stdin, control.stdout, control.stderr):
            if stream is not None:
                stream.close()
        for handle in handles:
            win32api.CloseHandle(handle)
        sentinel.terminate()
        sentinel.wait(timeout=5)


def test_cli_owner_refuses_a_parent_that_already_exited(tmp_path: Path) -> None:
    parent = subprocess.Popen([sys.executable, "-c", "pass"])
    parent.wait(timeout=10)
    control = _launch(_worker(tmp_path), parent.pid)
    _, stderr = control.communicate(timeout=15)
    assert control.returncode in (126, 128)
    assert b"DOCWEN_OWNER_CREATE_ERROR:" in stderr
    assert not (tmp_path / "ready.json").exists()


def test_parent_exit_terminates_a_worker_that_ignores_stdin(tmp_path: Path) -> None:
    import win32api
    import win32event

    command = _worker(tmp_path)
    command[-1] = command[-1].replace("sys.stdin.buffer.read(1)", "time.sleep(120)")
    supervisor = (
        "import os; from docwen_bundle.windows_machine_owner import run_job_worker; "
        f"raise SystemExit(run_job_worker({command!r},parent_pid=int(os.environ['TEST_PARENT_PID'])))"
    )
    parent_code = (
        "import json,os,pathlib,subprocess,sys,time; "
        "env=dict(os.environ,TEST_PARENT_PID=str(os.getpid())); "
        f"p=subprocess.Popen([sys.executable,'-c',{supervisor!r}],env=env,stdin=subprocess.DEVNULL,"
        "stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL); "
        f"pathlib.Path({str(tmp_path / 'owner.json')!r}).write_text(json.dumps({{'pid':p.pid}})); "
        "time.sleep(120)"
    )
    parent = subprocess.Popen([sys.executable, "-c", parent_code])
    owner_handle = None
    handles = []
    try:
        owner = _wait_json(tmp_path / "owner.json")
        owner_handle = win32api.OpenProcess(0x100001, False, owner["pid"])
        ready = _wait_json(tmp_path / "ready.json")
        handles = [_open_process(ready[key]) for key in ("worker", "child")]
        parent.terminate()
        parent.wait(timeout=5)
        for handle in [owner_handle, *handles]:
            assert win32event.WaitForSingleObject(handle, 5000) == win32event.WAIT_OBJECT_0
    finally:
        if parent.poll() is None:
            parent.terminate()
        parent.wait(timeout=5)
        if owner_handle is not None:
            if win32event.WaitForSingleObject(owner_handle, 0) != win32event.WAIT_OBJECT_0:
                win32api.TerminateProcess(owner_handle, 1)
                win32event.WaitForSingleObject(owner_handle, 5000)
            win32api.CloseHandle(owner_handle)
        for handle in handles:
            win32api.CloseHandle(handle)


@pytest.mark.parametrize("exit_code", [125, 126, 128])
def test_worker_exit_cannot_impersonate_supervisor_failure(exit_code: int) -> None:
    control = _launch([sys.executable, "-c", f"raise SystemExit({exit_code})"], os.getpid())
    _, stderr = control.communicate(timeout=15)
    assert control.returncode == 127
    assert b"DOCWEN_PROCESS_OWNER_READY:1" in stderr
    assert b"DOCWEN_OWNER_CREATE_ERROR" not in stderr


def test_supervisor_does_not_keep_a_closed_worker_stdin_alive(tmp_path: Path) -> None:
    if sys.platform != "win32":
        pytest.skip("Windows pipe handles")
    import msvcrt

    import pywintypes
    import win32file

    command = [
        PYTHON,
        "-c",
        (
            "import json,os,pathlib,time; os.close(0); "
            f"pathlib.Path({str(tmp_path / 'ready.json')!r}).write_text(json.dumps({{'ready':True}})); "
            "time.sleep(120)"
        ),
    ]
    control = _launch(command, os.getpid())
    try:
        _wait_json(tmp_path / "ready.json")
        assert control.stdin is not None
        with pytest.raises(pywintypes.error) as failure:
            win32file.WriteFile(msvcrt.get_osfhandle(control.stdin.fileno()), b"request")
        assert failure.value.winerror in (109, 232)  # BROKEN_PIPE / NO_DATA
    finally:
        control.kill()
        control.wait(timeout=5)
        for stream in (control.stdin, control.stdout, control.stderr):
            if stream is not None:
                stream.close()


def test_missing_worker_reports_creation_failure_without_readiness(tmp_path: Path) -> None:
    control = _launch([str(tmp_path / "missing.exe")], os.getpid())
    _, stderr = control.communicate(timeout=15)
    assert control.returncode == 126
    assert stderr.splitlines() == [b"DOCWEN_OWNER_CREATE_ERROR:2"]


@pytest.mark.parametrize("release_lock", [True, False])
def test_real_executable_sharing_conflict_has_bounded_retry(tmp_path: Path, release_lock: bool) -> None:
    import win32file

    target = tmp_path / "worker.exe"
    shutil.copyfile(Path(os.environ["SYSTEMROOT"]) / "System32" / "cmd.exe", target)
    lock = win32file.CreateFile(str(target), 0x80000000, 0, None, 3, 0, None)
    control = None
    try:
        started = time.monotonic()
        control = _launch([str(target), "/d", "/c", "exit", "0"], os.getpid())
        time.sleep(0.5)
        assert control.poll() is None
        if release_lock:
            lock.Close()
            lock = None
        _, stderr = control.communicate(timeout=12)
        elapsed = time.monotonic() - started
        if release_lock:
            assert control.returncode == 0
            assert b"DOCWEN_PROCESS_OWNER_READY:1" in stderr
            assert b"DOCWEN_OWNER_CREATE_ERROR" not in stderr
        else:
            assert control.returncode == 128
            assert stderr.splitlines() == [b"DOCWEN_OWNER_CREATE_ERROR:32"]
            assert 5 <= elapsed < 12
    finally:
        if lock is not None:
            lock.Close()
        if control is not None and control.poll() is None:
            control.kill()
            control.communicate(timeout=5)


def test_managed_cli_preserves_machine_frames_and_normal_eof(tmp_path: Path) -> None:
    from docwen_cli.machine.framing import FrameWriter, read_frame

    payload = io.BytesIO()
    writer = FrameWriter(payload)
    writer.write(
        {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocol": {"name": "docwen.machine", "major": 2, "minor": 0},
                "client": {"name": "managed-owner-test", "version": "1.0.0"},
                "features": {"progress": True, "cancellation": True},
            },
        }
    )
    writer.write({"jsonrpc": "2.0", "id": 2, "method": "capability/list", "params": {}})
    environment = dict(
        os.environ,
        DOCWEN_DATA_DIR=str(tmp_path / "profile"),
        DOCWEN_CONFIG_DIR=str(tmp_path / "profile/configs"),
        DOCWEN_LOG_DIR=str(tmp_path / "profile/logs"),
        DOCWEN_LOG_TO_TEMP="",
    )
    control = subprocess.Popen(
        [
            sys.executable,
            "-m",
            "docwen_bundle.cli_entry",
            "serve",
            "--stdio",
            "--managed",
            "--parent-pid",
            str(os.getpid()),
        ],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=environment,
    )
    try:
        stdout, stderr = control.communicate(payload.getvalue(), timeout=45)
        assert control.returncode == 0, stderr
        assert b"DOCWEN_PROCESS_OWNER_READY:1" in stderr
        frames = io.BytesIO(stdout)
        initialized = read_frame(frames)
        discovered = read_frame(frames)
        assert initialized is not None and initialized["id"] == 1
        assert initialized["result"]["server"]["name"] == "DocWen"
        assert discovered is not None and discovered["id"] == 2
        assert "convert.markdown.to_docx" in {c["capability_id"] for c in discovered["result"]["capabilities"]}
        assert read_frame(frames) is None
    finally:
        if control.poll() is None:
            control.kill()
        control.communicate(timeout=5)
