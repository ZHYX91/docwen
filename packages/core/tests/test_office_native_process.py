"""The native-family closeout budget and cleanup permission are independent."""

from __future__ import annotations

import io
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from docwen_core import office_native_process as native

pytestmark = pytest.mark.unit


@pytest.fixture
def clock(monkeypatch: pytest.MonkeyPatch) -> list[float]:
    now = [0.0]
    monkeypatch.setattr(native.time, "monotonic", lambda: now[0])
    monkeypatch.setattr(native.time, "sleep", lambda seconds: now.__setitem__(0, now[0] + seconds))
    return now


def test_launcher_exit_cannot_complete_a_still_active_family(clock: list[float]) -> None:
    family = MagicMock()
    family.exited.side_effect = lambda: clock[0] >= 0.12
    family.reap.return_value = 0
    confirmed, status, reason = native._drain_family(
        family, is_stopped=lambda: False, deadline=1, log=io.BytesIO(), log_limit=0
    )
    assert confirmed and status == 0 and not reason
    assert clock[0] >= 0.12
    family.terminate.assert_not_called()
    family.reap.assert_called_once()


@pytest.mark.parametrize("mode", ["cancel", "timeout", "launch-failed"])
def test_interrupted_family_is_drained_before_cleanup_permission(mode: str, clock: list[float]) -> None:
    family = MagicMock()
    forced: list[bool] = []
    family.terminate.side_effect = lambda *, force: forced.append(force)
    family.exited.side_effect = lambda: True in forced
    family.reap.return_value = -9
    confirmed, status, reason = native._drain_family(
        family,
        is_stopped=lambda: mode == "cancel",
        deadline=0.05 if mode == "timeout" else 20,
        log=io.BytesIO(),
        log_limit=0,
        initial_error="launch failed after creation" if mode == "launch-failed" else "",
    )
    assert confirmed and status == -9 and reason
    assert forced[0] is False and forced[-1] is True
    assert clock[0] < 1


@pytest.mark.parametrize("fault", ["terminate", "query", "never-exits"])
def test_no_termination_or_query_fault_can_reset_the_closeout_deadline(fault: str, clock: list[float]) -> None:
    family = MagicMock()
    family.exited.return_value = False
    if fault == "terminate":
        family.terminate.side_effect = PermissionError("native termination refused")
    if fault == "query":
        family.exited.side_effect = OSError("identity unavailable")
    confirmed, status, reason = native._drain_family(
        family, is_stopped=lambda: True, deadline=30, log=io.BytesIO(), log_limit=0
    )
    assert not confirmed and status is None and "could not be confirmed" in reason
    assert native._TERM_GRACE_S + native._REAP_TIMEOUT_S <= clock[0] < 3.5
    family.reap.assert_not_called()


def test_a_later_empty_group_does_not_erase_an_escaped_identity(clock: list[float]) -> None:
    family = MagicMock()
    family.exited.side_effect = [OSError("observed child escaped"), True]
    family.reap.return_value = 0
    confirmed, status, reason = native._drain_family(
        family, is_stopped=lambda: False, deadline=1, log=io.BytesIO(), log_limit=0
    )
    assert not confirmed and status == 0 and "ownership failed" in reason


def test_pre_cancelled_owner_never_creates_a_business_process(monkeypatch: pytest.MonkeyPatch) -> None:
    connection = MagicMock()
    factory = MagicMock(side_effect=AssertionError("must not launch"))
    monkeypatch.setattr(native, "_native_family", factory)
    monkeypatch.setattr(native.signal, "signal", MagicMock())
    native.native_owner_worker(connection, SimpleNamespace(is_set=lambda: True), ["owned-office"], 1)
    result = connection.send.call_args.args[0]
    assert result.cleanup_confirmed and not result.succeeded
    factory.assert_not_called()
