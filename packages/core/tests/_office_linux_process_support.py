"""Native launcher, process fixtures and bounded ownership fault workers."""

from __future__ import annotations

import json
import multiprocessing
import os
import shutil
import signal
import time
from pathlib import Path

import pytest
from tests.support.subprocess_runner import append_subprocess_run_record, run_subprocess

from docwen_core import office_process

_WNOHANG = getattr(os, "WNOHANG", 0)

_SIGKILL = getattr(signal, "SIGKILL", 0)

_NATIVE_SOURCE = r"""
#define _POSIX_C_SOURCE 200809L
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <signal.h>
#include <unistd.h>
#include <sys/types.h>
#include <time.h>
static void pause_tick(void) { struct timespec delay = {0, 50000000}; nanosleep(&delay, NULL); }
static void beat(const char *base, int index) {
    char name[8192]; snprintf(name, sizeof name, "%s.heartbeat", base);
    FILE *f = fopen(name, "w"); if (f) { fprintf(f, "%d\n", index); fclose(f); }
}
int main(int argc, char **argv) {
    if (argc == 3 && strcmp(argv[1], "--sentinel") == 0) {
        for (int i = 0; i < 400; i++) { beat(argv[2], i); pause_tick(); } return 0;
    }
    const char *base = getenv("DOCWEN_NATIVE_OFFICE_EVIDENCE");
    const char *mode = getenv("DOCWEN_NATIVE_OFFICE_MODE");
    const char *profile = ""; const char *out = "";
    if (!base || !mode || argc < 2) return 2;
    for (int i = 1; i < argc; i++) {
        if (strncmp(argv[i], "-env:UserInstallation=", 22) == 0) profile = argv[i] + 22;
        if (strcmp(argv[i], "--outdir") == 0 && i + 1 < argc) out = argv[i + 1];
    }
    pid_t child = fork(); if (child < 0) return 3; if (child > 0) return 0;
    if (setsid() < 0) return 4;
    if (strcmp(mode, "stubborn") == 0) signal(SIGTERM, SIG_IGN);
    char name[8192]; snprintf(name, sizeof name, "%s.identity", base);
    FILE *identity = fopen(name, "w"); if (!identity) return 5;
    fprintf(identity, "%ld\n%s\n%s\n%s\n", (long)getpid(), profile, out, argv[argc - 1]);
    char proc_path[256]; ssize_t proc_length = readlink("/proc/self", proc_path, sizeof proc_path - 1);
    if (proc_length <= 0) return 8; proc_path[proc_length] = 0; fprintf(identity, "%s\n", proc_path);
    FILE *status = fopen("/proc/self/status", "r"); if (!status) return 9;
    char status_line[1024]; int found = 0;
    while (fgets(status_line, sizeof status_line, status)) {
        if (strncmp(status_line, "NSpid:", 6) == 0) { fputs(status_line, identity); found = 1; break; }
    }
    fclose(status); fclose(identity); if (!found) return 10;
    for (int i = 0; i < 300; i++) {
        beat(base, i); pause_tick();
        if (strcmp(mode, "normal") == 0 && i >= 5) {
            snprintf(name, sizeof name, "%s.release", base);
            if (access(name, F_OK) != 0) continue;
            const char *source = argv[argc - 1]; const char *leaf = strrchr(source, '/'); leaf = leaf ? leaf + 1 : source;
            snprintf(name, sizeof name, "%s/%s", out, leaf);
            FILE *input = fopen(source, "rb"); FILE *output = fopen(name, "wb");
            if (!input || !output) return 6;
            char buffer[4096]; size_t count;
            while ((count = fread(buffer, 1, sizeof buffer, input)) > 0) fwrite(buffer, 1, count, output);
            fclose(input); fclose(output); return 0;
        }
    }
    return 7;
}
"""


@pytest.fixture
def native_office(tmp_path, monkeypatch):
    compiler = shutil.which("cc") or shutil.which("gcc")
    if compiler is None:
        pytest.skip("Native C compiler is required for the Linux lifecycle fixture")
    if not hasattr(os, "pidfd_open") or not hasattr(signal, "pidfd_send_signal"):
        pytest.skip("Linux pidfd support is required")
    source = tmp_path / "office-fixture.c"
    binary = tmp_path / "office-fixture"
    source.write_text(_NATIVE_SOURCE, encoding="utf-8")
    run_subprocess([compiler, "-std=c11", "-O2", str(source), "-o", str(binary)], check=True, timeout=30)
    actual_run = office_process.run_owned_office

    def recorded_owner(cmd, *, is_cancelled, timeout_s):
        started = time.perf_counter()
        result = None
        try:
            result = actual_run(cmd, is_cancelled=is_cancelled, timeout_s=timeout_s)
            return result
        finally:
            append_subprocess_run_record(
                {
                    "argv": cmd,
                    "command": "Owned native Office launcher: " + " ".join(cmd),
                    "returncode": result.launcher_exit_code if result is not None else None,
                    "owner_exitcode": result.owner_exit_code if result is not None else None,
                    "owner_pid": result.owner_pid if result is not None else None,
                    "reaped_pids": list(result.reaped_pids) if result is not None else [],
                    "cleanup_confirmed": result.cleanup_confirmed if result is not None else False,
                    "timed_out": result is not None and "timed out" in result.reason,
                    "timeout_seconds": timeout_s,
                    "duration_seconds": time.perf_counter() - started,
                    "text": False,
                    "capture_output": False,
                }
            )

    monkeypatch.setattr(office_process, "run_owned_office", recorded_owner)
    return binary


def _start_ticks(pid: int) -> int | None:
    try:
        text = Path(f"/proc/{pid}/stat").read_text()
    except FileNotFoundError:
        return None
    return int(text[text.rfind(")") + 2 :].split()[19])


def _active_evidence(base: Path, deadline: float = 5.0) -> tuple[int, int, list[str]]:
    limit = time.monotonic() + deadline
    while time.monotonic() < limit:
        try:
            lines = base.with_suffix(".identity").read_text().splitlines()
            if len(lines) == 6 and base.with_suffix(".heartbeat").read_text().strip():
                pid = int(lines[0])
                proc_pid = int(lines[4])
                namespace_ids = [int(value) for value in lines[5].split()[1:]]
                assert lines[5].startswith("NSpid:")
                assert namespace_ids[-1] == pid and namespace_ids[0] == proc_pid
                ticks = _start_ticks(proc_pid)
                if ticks is not None:
                    return pid, ticks, lines
        except (FileNotFoundError, ValueError):
            pass
        time.sleep(0.01)
    raise AssertionError("The native Office child was never observed active")


def _drop_receipt_worker(connection, stop, cmd, timeout_s):
    class DropFinalReceipt:
        def send(self, message):
            if not isinstance(message, office_process.OfficeProcessResult):
                connection.send(message)

        def close(self):
            connection.close()

    office_process._owner_worker(DropFinalReceipt(), stop, cmd, timeout_s)


def _abnormal_receipt_worker(connection, stop, cmd, timeout_s):
    class ExitAfterReceipt:
        def send(self, message):
            connection.send(message)
            if isinstance(message, office_process.OfficeProcessResult):
                os._exit(17)

        def close(self):
            connection.close()

    office_process._owner_worker(ExitAfterReceipt(), stop, cmd, timeout_s)


def _start_owner_with_abnormal_exit(cmd, timeout_s):
    context = multiprocessing.get_context("spawn")
    parent, child = context.Pipe(duplex=False)
    stop = context.Event()
    owner = context.Process(target=_abnormal_receipt_worker, args=(child, stop, cmd, timeout_s), daemon=True)
    owner.start()
    child.close()
    return owner, parent, stop


def _start_owner_losing_receipt(cmd, timeout_s):
    context = multiprocessing.get_context("spawn")
    parent, child = context.Pipe(duplex=False)
    stop = context.Event()
    owner = context.Process(target=_drop_receipt_worker, args=(child, stop, cmd, timeout_s), daemon=True)
    owner.start()
    child.close()
    return owner, parent, stop


def _discovery_fixture_worker(connection, stop, cmd, timeout_s):
    """Inject only proc discovery availability; all process/wait/pidfd work is real."""
    evidence = Path(os.environ["DOCWEN_NATIVE_OFFICE_EVIDENCE"])
    mode = os.environ["DOCWEN_NATIVE_OFFICE_DISCOVERY"]
    task_root = evidence.with_suffix(".task-view")
    # This is a test-provided children interface populated from the independent
    # native C identity. It does not claim that this host exposes /proc children.
    owner_proc_pid = int(os.readlink("/proc/self"))
    task = task_root / str(owner_proc_pid)
    task.mkdir(parents=True)
    if mode == "available":
        (task / "children").write_text("", encoding="ascii")
    office_process._TASK_ROOT = task_root
    actual_task_reader = office_process._task_child_proc_pids
    actual_status_reader = office_process._status_direct_children

    def record(method, values):
        with evidence.with_suffix(".discovery.jsonl").open("a", encoding="utf-8") as stream:
            stream.write(
                json.dumps({"method": method, "values": sorted(values) if values is not None else None}) + "\n"
            )

    def task_reader(proc_pid, deadline):
        if mode == "available":
            try:
                native = evidence.with_suffix(".identity").read_text().splitlines()
                if len(native) == 6:
                    (task / "children").write_text(native[4] + "\n", encoding="ascii")
            except FileNotFoundError:
                pass
        values = actual_task_reader(proc_pid, deadline)
        record("task", values)
        return values

    def status_reader(proc_pid, depth, deadline):
        values = actual_status_reader(proc_pid, depth, deadline)
        record("status", values)
        return values

    office_process._task_child_proc_pids = task_reader
    office_process._status_direct_children = status_reader
    office_process._owner_worker(connection, stop, cmd, timeout_s)


def _start_owner_with_discovery_fixture(cmd, timeout_s):
    context = multiprocessing.get_context("spawn")
    parent, child = context.Pipe(duplex=False)
    stop = context.Event()
    owner = context.Process(target=_discovery_fixture_worker, args=(child, stop, cmd, timeout_s), daemon=True)
    owner.start()
    child.close()
    return owner, parent, stop


def _deny_kill_then_rescue_worker(connection, stop, cmd, timeout_s):
    """Real owner/native child; deny only production KILL, then test-owned rescue."""
    evidence = Path(os.environ["DOCWEN_NATIVE_OFFICE_EVIDENCE"])
    actual_open = getattr(os, "pidfd_open", None)
    actual_send = getattr(signal, "pidfd_send_signal", None)
    waitid = getattr(os, "waitid", None)
    p_pid = getattr(os, "P_PID", None)
    wexited = getattr(os, "WEXITED", None)
    wnowait = getattr(os, "WNOWAIT", None)
    assert actual_open is not None and actual_send is not None and waitid is not None
    assert p_pid is not None and wexited is not None and wnowait is not None
    opened: dict[int, int] = {}
    held: list[tuple[int, int]] = []
    denied = 0

    def open_pidfd(pid, flags=0):
        descriptor = actual_open(pid, flags)
        opened[descriptor] = pid
        return descriptor

    def send_pidfd(descriptor, sig, *args):
        nonlocal denied
        if sig == _SIGKILL:
            denied += 1
            native = evidence.with_suffix(".identity").read_text().splitlines()
            target = opened[descriptor]
            assert len(native) == 6 and int(native[0]) == target
            if not held:
                # Keep the original kernel identity alive across production's
                # fd close; rescue does not rediscover a numeric PID.
                held.append((target, os.dup(descriptor)))
            raise PermissionError("Test fault: production KILL is refused")
        return actual_send(descriptor, sig, *args)

    patch = pytest.MonkeyPatch()
    patch.setattr(os, "pidfd_open", open_pidfd)
    patch.setattr(signal, "pidfd_send_signal", send_pidfd)
    office_process._owner_worker(connection, stop, cmd, timeout_s)
    assert held and denied
    native_pid, rescue_fd = held[0]
    flags = wexited | _WNOHANG | wnowait
    live_observation = waitid(p_pid, native_pid, flags)
    assert live_observation is None or live_observation.si_pid == 0
    evidence.with_suffix(".kill-denied.json").write_text(
        json.dumps(
            {
                "nativePid": native_pid,
                "productionKillAttempts": denied,
                "nativeKernelWaitChildLiveAfterProductionReturn": True,
                "productionKillFault": "PermissionError",
                "rescueReleased": False,
            }
        ),
        encoding="utf-8",
    )
    deadline = time.monotonic() + 12.0
    while not evidence.with_suffix(".rescue-release").exists() and time.monotonic() < deadline:
        time.sleep(0.02)
    assert evidence.with_suffix(".rescue-release").exists(), "Parent did not authorize fixture rescue"
    # The parent has independently asserted continued activity and retained
    # resources. This rescue is test closeout, never a production stop receipt.
    live_observation = waitid(p_pid, native_pid, flags)
    assert live_observation is None or live_observation.si_pid == 0
    actual_send(rescue_fd, _SIGKILL)
    reaped = []
    complete = False
    deadline = time.monotonic() + 2.0
    while time.monotonic() < deadline:
        try:
            child, status = os.waitpid(-1, _WNOHANG)
        except ChildProcessError:
            complete = True
            break
        if child:
            reaped.append({"pid": child, "exitCode": os.waitstatus_to_exitcode(status)})
        else:
            time.sleep(0.01)
    os.close(rescue_fd)
    evidence.with_suffix(".rescue.json").write_text(
        json.dumps(
            {
                "nativePid": native_pid,
                "originalPidfdUsed": True,
                "kernelWaitChildLiveBeforeRescue": True,
                "reaped": reaped,
                "ECHILD": complete,
            }
        ),
        encoding="utf-8",
    )
    assert complete and any(row["pid"] == native_pid and row["exitCode"] == -_SIGKILL for row in reaped)
