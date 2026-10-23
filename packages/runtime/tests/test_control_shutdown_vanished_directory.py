"""Shutdown only tolerates a vanished owned Unix directory after cleanup."""

from __future__ import annotations

import sys
from typing import Any
from uuid import uuid4

import pytest

pytestmark = pytest.mark.contract


@pytest.mark.skipif(sys.platform == "win32", reason="Unix managed control endpoint race")
@pytest.mark.parametrize(
    ("root_cause", "suppress"),
    [
        (FileNotFoundError("owned directory removed by shutdown"), True),
        (PermissionError("endpoint ownership changed"), False),
        (OSError("endpoint identity changed"), False),
    ],
)
def test_stop_handles_only_vanished_directory_wake_race_after_clean_shutdown(
    monkeypatch: pytest.MonkeyPatch,
    root_cause: OSError,
    suppress: bool,
) -> None:
    from docwen_runtime.control import ControlClient, ControlEndpointError, ControlServer

    # Force fallback managed-directory ownership even when CI provides an
    # XDG private runtime directory.
    monkeypatch.delenv("XDG_RUNTIME_DIR", raising=False)
    server = ControlServer(lambda _action, _payload: {}, app_name=f"dw-owned-dir-race-{uuid4().hex}")
    server.start(timeout=2)
    assert server._unix_lease is not None
    assert server._unix_lease.plan.directory_name is not None
    original_request = ControlClient.request

    def wake_then_fail(
        client: ControlClient,
        action: str,
        payload: dict[str, Any] | None = None,
        *,
        timeout: float,
    ) -> dict[str, Any]:
        # Really wake the serving thread; the injected failure models the
        # endpoint-plan validation error racing removal by that same thread.
        original_request(client, action, payload, timeout=timeout)
        try:
            raise root_cause
        except OSError as exc:
            raise ControlEndpointError() from exc

    monkeypatch.setattr(ControlClient, "request", wake_then_fail)
    try:
        if suppress:
            server.stop(timeout=2)
        else:
            with pytest.raises(ControlEndpointError) as error:
                server.stop(timeout=2)
            assert isinstance(error.value.__cause__, type(root_cause))
        assert server._thread is None
        assert server._listener is None
        assert server._unix_lease is None
        server.stop(timeout=2)
    finally:
        if server._thread is not None:
            monkeypatch.setattr(ControlClient, "request", original_request)
            server.stop(timeout=2)


def test_stop_does_not_ignore_arbitrary_wake_errors_after_clean_join(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import ControlClient, ControlServer

    server = ControlServer(lambda _action, _payload: {}, app_name=f"dw-arbitrary-wake-{uuid4().hex}")
    server.start(timeout=2)
    original_request = ControlClient.request

    def wake_then_fail(
        client: ControlClient,
        action: str,
        payload: dict[str, Any] | None = None,
        *,
        timeout: float,
    ) -> dict[str, Any]:
        original_request(client, action, payload, timeout=timeout)
        raise RuntimeError("other wake failure")

    monkeypatch.setattr(ControlClient, "request", wake_then_fail)
    try:
        with pytest.raises(RuntimeError, match="other wake failure"):
            server.stop(timeout=2)
        assert server._listener is None
        assert server._thread is None
    finally:
        if server._thread is not None:
            monkeypatch.setattr(ControlClient, "request", original_request)
            server.stop(timeout=2)
