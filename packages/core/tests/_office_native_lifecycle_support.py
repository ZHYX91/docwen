"""Bounded, test-owned subprocesses; no installed Office is started or killed."""

from __future__ import annotations

import json
import os
import signal
import subprocess
import sys
import time
from dataclasses import asdict
from pathlib import Path


def family_command(evidence: Path, mode: str) -> list[str]:
    return [sys.executable, str(Path(__file__).resolve()), "launcher", str(evidence), mode]


def _child(evidence: Path, mode: str) -> None:
    if mode in {"cancel", "timeout", "kill-refused"}:
        signal.signal(signal.SIGTERM, signal.SIG_IGN)
    evidence.with_suffix(".heartbeat").write_text(str(time.monotonic_ns()) + "\n", encoding="utf-8")
    evidence.with_suffix(".child").write_text(str(os.getpid()), encoding="utf-8")
    if mode == "large-output":
        sent = {"stdout": 0, "stderr": 0}
        for _ in range(8):
            for descriptor, value, name in ((1, b"o", "stdout"), (2, b"e", "stderr")):
                pending = memoryview(value * (512 * 1024))
                while pending:
                    written = os.write(descriptor, pending)
                    assert 0 < written <= len(pending)
                    sent[name] += written
                    pending = pending[written:]
        evidence.with_suffix(".output-bytes").write_text(json.dumps(sent), encoding="utf-8")
    deadline = time.monotonic() + (2.0 if mode in {"cancel", "timeout", "kill-refused"} else 0.25)
    while time.monotonic() < deadline:
        with evidence.with_suffix(".heartbeat").open("a", encoding="utf-8") as heartbeat:
            heartbeat.write(str(time.monotonic_ns()) + "\n")
        time.sleep(0.01)
    evidence.with_suffix(".finished").write_text("natural exit", encoding="utf-8")


def _portable_group_probe(evidence: Path, mode: str) -> None:
    """Exercise the real group owner on Linux, substituting Darwin query APIs.

    The test subreaper also collects its test-owned orphan zombies after the
    production group proof. It does not upgrade the production cleanup receipt.
    """
    if sys.platform != "linux":
        raise RuntimeError("The inventory-substitution probe requires Linux")
    from docwen_core import macos_process as mac
    from docwen_core import office_native_process as native
    from docwen_core.office_process import OfficeProcessResult, _enable_owner

    class ProcfsInventory:
        def identity(self, pid: int) -> mac.MacProcessIdentity | None:
            try:
                stat = Path(f"/proc/{pid}/stat").read_text().rsplit(")", 1)[1].split()
            except FileNotFoundError:
                return None
            return mac.MacProcessIdentity(pid, int(stat[2]), 5 if stat[0] == "Z" else 2, (int(stat[19]), 0))

        def group(self, group_id: int) -> tuple[mac.MacProcessIdentity, ...]:
            members = []
            for entry in Path("/proc").iterdir():
                if entry.name.isdecimal():
                    identity = self.identity(int(entry.name))
                    if identity is not None and identity.group == group_id:
                        members.append(identity)
            return tuple(members)

    class POSIXChildWaiter:
        def exited(self, pid: int) -> bool:
            result = os.waitid(os.P_PID, pid, os.WEXITED | os.WNOHANG | os.WNOWAIT)
            return result is not None and result.si_pid == pid

    class Stop:
        def is_set(self) -> bool:
            return mode in {"cancel", "kill-refused"} and evidence.with_suffix(".child").exists()

    class Connection:
        def send(self, message: object) -> None:
            if isinstance(message, OfficeProcessResult):
                evidence.with_suffix(".result").write_text(json.dumps(asdict(message)), encoding="utf-8")

        def close(self) -> None:
            pass

    _enable_owner()
    mac.MacProcessInventory = ProcfsInventory
    mac.MacChildWaiter = POSIXChildWaiter
    native._native_family = mac.MacOSProcessGroup
    if mode == "kill-refused":
        actual_signal = mac.os.killpg

        def deny_kill(group: int, requested: int) -> None:
            if requested == signal.SIGKILL:
                raise PermissionError("test-owned native KILL refusal")
            actual_signal(group, requested)

        mac.os.killpg = deny_kill
        native._TERM_GRACE_S, native._REAP_TIMEOUT_S = 0.05, 0.2
    native.native_owner_worker(Connection(), Stop(), family_command(evidence, mode), 0.5 if mode == "timeout" else 5)
    # Give a refused child only its original finite TTL; never interpret this
    # test rescue as a production receipt or delete its still-active files.
    deadline = time.monotonic() + 3
    while time.monotonic() < deadline:
        try:
            child, _status = os.waitpid(-1, os.WNOHANG)
        except ChildProcessError:
            return
        if child == 0:
            time.sleep(0.02)
    raise RuntimeError("test-owned child exceeded its short TTL")


if __name__ == "__main__":
    action, raw_evidence, selected_mode = sys.argv[1:]
    selected = Path(raw_evidence)
    if action == "child":
        _child(selected, selected_mode)
    elif action == "launcher":
        subprocess.Popen([sys.executable, __file__, "child", str(selected), selected_mode])
        # Exit before the child, deliberately leaving inherited standard handles.
        os._exit(0)
    elif action == "portable-group-probe":
        _portable_group_probe(selected, selected_mode)
