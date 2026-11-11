"""Darwin ABI/identity logic with injected native queries, never host claims."""

from __future__ import annotations

import ctypes
import errno
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
    monkeypatch.setattr(mac.os, "WNOHANG", 1, raising=False)
    waiter = MagicMock()
    waiter.exited.return_value = True
    monkeypatch.setattr(mac, "MacChildWaiter", lambda: waiter)
    process = MagicMock(pid=123)
    process.poll.side_effect = AssertionError("must not reap the group identity through poll")
    process.communicate.side_effect = AssertionError("must not wait on inherited pipes")
    monkeypatch.setattr(mac.subprocess, "Popen", MagicMock(return_value=process))
    family = mac.MacOSProcessGroup()
    family.start(["office"], stdin=10, stdout=11, cwd=None, environment=None)
    return family, inventory


def test_group_owner_does_not_require_python_waitid(
    family: tuple[mac.MacOSProcessGroup, MagicMock], monkeypatch: pytest.MonkeyPatch
) -> None:
    # Python 3.12 on Darwin omits this wrapper and its associated constants.
    for name in ("waitid", "P_PID", "WEXITED", "WNOWAIT"):
        monkeypatch.delattr(mac.os, name, raising=False)
    owner = mac.MacOSProcessGroup()
    owner.start(["office"], stdin=10, stdout=11, cwd=None, environment=None)
    inventory = family[1]
    inventory.identity.return_value = _identity(123, state=5)
    inventory.group.return_value = (_identity(123, state=5),)
    assert owner.exited()
    owner.terminate(force=True)


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
    monkeypatch.setattr(owner.child_waiter, "exited", MagicMock(side_effect=ChildProcessError("identity gone")))
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


@pytest.mark.parametrize("state", [2, 5])
def test_darwin_identity_reads_the_full_public_abi_including_unreaped_zombies(state: int) -> None:
    api = MagicMock()

    def fill(pid: int, flavor: int, argument: int, pointer: Any, size: int) -> int:
        assert flavor == 3 and argument == 1 and size == 136
        info = ctypes.cast(pointer, ctypes.POINTER(mac._ProcBSDInfo)).contents
        info.pid, info.pgid, info.status = pid, 123, state
        info.start_seconds, info.start_microseconds = 100, 250
        return size

    api.proc_pidinfo.side_effect = fill
    assert _inventory(api).identity(124) == mac.MacProcessIdentity(124, 123, state, (100, 250))


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


@pytest.fixture
def wait_api(monkeypatch: pytest.MonkeyPatch) -> tuple[mac.MacChildWaiter, MagicMock]:
    api = MagicMock()
    loader = MagicMock(return_value=api)
    monkeypatch.setattr(mac.sys, "platform", "darwin")
    monkeypatch.setattr(mac.ctypes, "CDLL", loader)
    waiter = mac.MacChildWaiter()
    loader.assert_called_once_with("/usr/lib/libSystem.B.dylib", use_errno=True)
    assert api.waitid.argtypes == [ctypes.c_int, ctypes.c_uint32, ctypes.POINTER(mac._SigInfo), ctypes.c_int]
    assert api.waitid.restype is ctypes.c_int
    return waiter, api


@pytest.mark.parametrize("exited", [False, True])
def test_darwin_wait_child_uses_native_layout_and_nonreaping_nonblocking_flags(
    wait_api: tuple[mac.MacChildWaiter, MagicMock], exited: bool
) -> None:
    waiter, api = wait_api
    assert ctypes.sizeof(mac._SigInfo) == 104
    assert (mac._SigInfo.pid.offset, mac._SigInfo.value.offset, mac._SigInfo.reserved.offset) == (12, 32, 48)

    def wait(id_type: int, pid: int, pointer: Any, flags: int) -> int:
        assert id_type == 1 and pid == 123 and flags == 0x25
        assert ctypes.string_at(pointer, 104) == bytes(104)
        if exited:
            info = ctypes.cast(pointer, ctypes.POINTER(mac._SigInfo)).contents
            info.pid, info.signo, info.code = pid, 20, 1
        return 0

    api.waitid.side_effect = wait
    assert waiter.exited(123) is exited


@pytest.mark.parametrize("error", [errno.ECHILD, errno.EPERM, errno.EINVAL])
def test_darwin_wait_child_query_error_is_never_an_exit_receipt(
    wait_api: tuple[mac.MacChildWaiter, MagicMock], error: int
) -> None:
    waiter, api = wait_api
    api.waitid.side_effect = lambda *_args: (ctypes.set_errno(error), -1)[1]
    with pytest.raises(OSError) as caught:
        waiter.exited(123)
    assert caught.value.errno == error
    assert api.waitid.call_count == 1


@pytest.mark.parametrize("pid,signo,code", [(124, 20, 1), (123, 0, 1), (123, 20, 4)])
def test_darwin_wait_child_rejects_mismatched_or_nonterminal_receipts(
    wait_api: tuple[mac.MacChildWaiter, MagicMock], pid: int, signo: int, code: int
) -> None:
    waiter, api = wait_api

    def wait(_id_type: int, _pid: int, pointer: Any, _flags: int) -> int:
        info = ctypes.cast(pointer, ctypes.POINTER(mac._SigInfo)).contents
        info.pid, info.signo, info.code = pid, signo, code
        return 0

    api.waitid.side_effect = wait
    with pytest.raises(OSError, match=r"invalid.*receipt"):
        waiter.exited(123)


def test_darwin_wait_child_retries_eintr_only_inside_one_finite_budget(
    wait_api: tuple[mac.MacChildWaiter, MagicMock], monkeypatch: pytest.MonkeyPatch
) -> None:
    waiter, api = wait_api
    api.waitid.side_effect = lambda *_args: (ctypes.set_errno(errno.EINTR), -1)[1]
    monkeypatch.setattr(mac.time, "monotonic", MagicMock(side_effect=[0.0, 0.1, 0.3]))
    with pytest.raises(InterruptedError):
        waiter.exited(123)
    assert api.waitid.call_count == 2
