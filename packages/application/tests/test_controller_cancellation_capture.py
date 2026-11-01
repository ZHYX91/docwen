"""Exact cancellation capture before aliases and safe same-ID reuse."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_application.controller import ApplicationController, ControllerError
from docwen_core.models.request import ConversionRequest, FileRef

pytestmark = pytest.mark.contract


def _request(tmp_path: Path) -> ConversionRequest:
    refs = []
    for index in range(2):
        path = tmp_path / f"input-{index}.csv"
        path.write_bytes(f"Name,Amount\nrow-{index},7\n".encode())
        refs.append(FileRef(path=str(path), format="csv", category="spreadsheet"))
    return ConversionRequest(request_id="captured-owner", input_refs=refs, target_format="xlsx")


def test_failed_capture_publishes_no_alias_and_retry_retains_legacy_semantics(tmp_path):
    controller = ApplicationController()
    request = _request(tmp_path)
    captured = []

    def capture(scope):
        assert controller._cancellation_scopes == {}
        captured.append(scope)
        raise RuntimeError("controlled capture failure")

    with pytest.raises(RuntimeError, match="controlled capture failure"):
        controller.prepare_execution_cancellation(request, batch=True, on_reserved=capture)
    assert controller._cancellation_scopes == {}
    [failed] = captured
    retained = controller.prepare_execution_cancellation(request, batch=True)
    assert retained is not failed
    assert controller.prepare_execution_cancellation(request, batch=True) is retained
    controller.release_execution_cancellation(request.request_id, failed)
    assert set(controller._cancellation_scopes) == {"captured-owner", "captured-owner-0", "captured-owner-1"}
    assert all(scope is retained for scope in controller._cancellation_scopes.values())
    controller.release_execution_cancellation(request.request_id, retained)
    assert controller._cancellation_scopes == {}


def test_interrupted_partial_alias_publication_removes_only_captured_scope(tmp_path):
    controller = ApplicationController()
    request = _request(tmp_path)
    captured = []
    seen_aliases = []

    class InterruptAfterSecondAlias(dict):
        armed = True

        def __setitem__(self, key, value):
            super().__setitem__(key, value)
            if self.armed and len(self) == 2:
                self.armed = False
                seen_aliases.extend(self)
                raise KeyboardInterrupt

    controller._cancellation_scopes = InterruptAfterSecondAlias()

    def capture(scope):
        assert controller._cancellation_scopes == {}
        captured.append(scope)

    with pytest.raises(KeyboardInterrupt):
        controller.prepare_execution_cancellation(request, batch=True, on_reserved=capture)
    assert seen_aliases == ["captured-owner", "captured-owner-0"]
    assert controller._cancellation_scopes == {}
    assert len(captured) == 1
    retry = controller.prepare_execution_cancellation(request, batch=True, on_reserved=captured.append)
    assert retry is captured[-1] and retry is not captured[0]
    controller.release_execution_cancellation(request.request_id, retry)
    assert controller._cancellation_scopes == {}


def test_exclusive_capture_refuses_existing_owner_without_changing_retention(tmp_path):
    controller = ApplicationController()
    request = _request(tmp_path)
    original = controller.prepare_execution_cancellation(request, batch=True)
    captured = []
    with pytest.raises(ControllerError, match="already owned"):
        controller.prepare_execution_cancellation(request, batch=True, on_reserved=captured.append)
    assert captured == []
    assert original.retain_until_release and not original.claimed and not original.token.is_cancelled
    assert all(scope is original for scope in controller._cancellation_scopes.values())
    controller.release_execution_cancellation(request.request_id, original)
    assert controller._cancellation_scopes == {}


def test_stale_exact_cancel_and_release_cannot_affect_a_new_same_id_owner(tmp_path):
    controller = ApplicationController()
    request = _request(tmp_path)
    handles = []
    previous = controller.prepare_execution_cancellation(request, on_reserved=handles.append)
    controller.release_execution_cancellation(request.request_id, previous)
    current = controller.prepare_execution_cancellation(request, on_reserved=handles.append)
    assert previous is handles[0] and current is handles[1] and previous is not current
    controller.cancel(request.request_id, reservation=previous)
    controller.release_execution_cancellation(request.request_id, previous)
    assert not current.token.is_cancelled
    assert controller._cancellation_scopes[request.request_id] is current
    controller.cancel(request.request_id, reservation=current)
    assert current.token.is_cancelled
    controller.release_execution_cancellation(request.request_id, current)
    assert controller._cancellation_scopes == {}
