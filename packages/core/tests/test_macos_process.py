"""Darwin ABI/identity logic with injected native queries, never host claims."""

from __future__ import annotations

import ctypes
import errno
import os
from types import SimpleNamespace
from typing import Any
from unittest.mock import MagicMock

import pytest

from docwen_core import macos_process as mac

pytestmark = pytest.mark.unit


def _identity(pid: int, *, group: int = 123, state: int = 2, start: int = 10) -> mac.MacProcessIdentity:
    return mac.MacProcessIdentity(pid, group, state, (start, 0))


@pytest.fixture
def family(monkeypatch: pytest.MonkeyPatch) -> tuple[mac.MacOSProcessGroup, MagicMock]:
    inventory = MagicMock()
    inventory.identity.return_value = _identity(123)
    monkeypatch.setattr(mac, "MacProcessInventory", lambda: inventory)
    monkeypatch.setattr(mac.os, "setsid", MagicMock(), raising=False)
    monkeypatch.setattr(mac.os, "killpg", MagicMock(), raising=False)
    monkeypatch.setattr(mac.signal, "SIGKILL", getattr(mac.signal, "SIGKILL", 9), raising=False)
    monkeypatch.setattr(mac.os, "waitid", MagicMock(return_value=SimpleNamespace(si_pid=123)), raising=False)
    for name, value in (("P_PID", 1), ("WEXITED", 4), ("WNOHANG", 1), ("WNOWAIT", 0x01000000)):
        monkeypatch.setattr(mac.os, name, getattr(os, name, value), raising=False)
    process = MagicMock(pid=123)
    process.poll.side_effect = AssertionError("must not reap the group identity through poll")
    process.communicate.side_effect = AssertionError("must not wait on inherited pipes")
    monkeypatch.setattr(mac.subprocess, "Popen", MagicMock(return_value=process))
    family = mac.MacOSProcessGroup()
    family.start(["office"], stdin=10, stdout=11, cwd=None, environment=None)
    return family, inventory


def test_parent_exit_is_not_group_exit_and_launcher_is_reaped_last(
    family: tuple[mac.MacOSProcessGroup, MagicMock], monkeypatch: pytest.MonkeyPatch
) -> None:
    owner, inventory = family
    identities = {123: _identity(123, state=5), 124: _identity(124)}
    inventory.group.side_effect = lambda _group: tuple(identities.values())
    inventory.identity.side_effect = identities.get
    assert not owner.exited()
    signals = MagicMock()
    monkeypatch.setattr(mac.os, "killpg", signals, raising=False)
    owner.terminate(force=True)
    signals.assert_called_once_with(123, 9)
    identities[124] = _identity(124, state=5)
    assert owner.exited()
    waiter = MagicMock(return_value=(123, 0))
    monkeypatch.setattr(mac.os, "waitpid", waiter)
    assert owner.reap() == 0
    waiter.assert_called_once_with(123, 1)
    owner.terminate(force=True)
    assert signals.call_count == 1  # No numeric group signals after releasing its PID.


def test_observed_member_leaving_group_invalidates_the_receipt(family: tuple[mac.MacOSProcessGroup, MagicMock]) -> None:
    owner, inventory = family
    owner.known[124] = _identity(124)
    inventory.group.return_value = (_identity(123, state=5),)
    inventory.identity.side_effect = lambda pid: _identity(pid, state=5) if pid == 123 else _identity(pid, group=999)
    with pytest.raises(OSError, match="escaped"):
        owner.exited()


def test_empty_snapshot_cannot_erase_a_previously_observed_live_member(
    family: tuple[mac.MacOSProcessGroup, MagicMock],
) -> None:
    owner, inventory = family
    owner.known[124] = _identity(124)
    inventory.group.return_value = ()
    inventory.identity.side_effect = lambda pid: _identity(pid, state=5 if pid == 123 else 2)
    assert not owner.exited()


def test_lost_wait_child_authority_never_sends_a_group_signal(
    family: tuple[mac.MacOSProcessGroup, MagicMock], monkeypatch: pytest.MonkeyPatch
) -> None:
    owner, _inventory = family
    monkeypatch.setattr(mac.os, "waitid", MagicMock(side_effect=ChildProcessError("identity gone")))
    signals = MagicMock()
    monkeypatch.setattr(mac.os, "killpg", signals, raising=False)
    with pytest.raises(ChildProcessError):
        owner.terminate(force=True)
    signals.assert_not_called()


def test_reused_nonlauncher_pid_is_not_claimed_as_a_current_owned_child(
    family: tuple[mac.MacOSProcessGroup, MagicMock], monkeypatch: pytest.MonkeyPatch
) -> None:
    owner, inventory = family
    owner.known[124] = _identity(124)
    inventory.group.return_value = (_identity(123, state=5),)
    inventory.identity.side_effect = lambda pid: (
        _identity(pid, state=5) if pid == 123 else _identity(pid, group=999, start=20)
    )
    assert owner.exited()
    assert 124 not in owner.known


def _inventory(api: MagicMock) -> mac.MacProcessInventory:
    inventory = object.__new__(mac.MacProcessInventory)
    inventory.api = api
    return inventory


def test_darwin_identity_reads_the_full_public_abi() -> None:
    api = MagicMock()

    def fill(pid: int, flavor: int, argument: int, pointer: Any, size: int) -> int:
        assert flavor == 3 and argument == 0 and size == 136
        info = ctypes.cast(pointer, ctypes.POINTER(mac._ProcBSDInfo)).contents
        info.pid, info.pgid, info.status = pid, 123, 2
        info.start_seconds, info.start_microseconds = 100, 250
        return size

    api.proc_pidinfo.side_effect = fill
    assert _inventory(api).identity(124) == mac.MacProcessIdentity(124, 123, 2, (100, 250))


@pytest.mark.parametrize("error,size", [(0, 0), (errno.EPERM, 0), (0, 12)])
def test_incomplete_identity_is_not_an_exited_process(error: int, size: int) -> None:
    api = MagicMock()
    api.proc_pidinfo.side_effect = lambda *_args: (ctypes.set_errno(error), size)[1]
    with pytest.raises(OSError, match="identity"):
        _inventory(api).identity(124)


def test_esrch_is_the_only_missing_identity_receipt() -> None:
    api = MagicMock()
    api.proc_pidinfo.side_effect = lambda *_args: (ctypes.set_errno(errno.ESRCH), 0)[1]
    assert _inventory(api).identity(124) is None


@pytest.mark.parametrize("size,error", [(-1, errno.EPERM), (0, errno.EPERM), (3, 0), (4 * 32769, 0)])
def test_group_inventory_errors_and_truncation_are_not_empty_groups(size: int, error: int) -> None:
    api = MagicMock()
    api.proc_listpids.side_effect = lambda *_args: (ctypes.set_errno(error), size)[1]
    with pytest.raises(OSError):
        _inventory(api).group(123)
