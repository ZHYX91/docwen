"""Tests for CancellationToken and CancellationTokenView."""

from __future__ import annotations

import threading

import pytest

from docwen_core.cancellation import CancellationToken
from docwen_core.errors import CancellationRequested

pytestmark = pytest.mark.unit


class TestCancellationToken:
    def test_initial_state_not_cancelled(self) -> None:
        token = CancellationToken()
        assert token.is_cancelled is False

    def test_cancel_sets_flag(self) -> None:
        token = CancellationToken()
        token.cancel()
        assert token.is_cancelled is True

    def test_cancel_idempotent(self) -> None:
        token = CancellationToken()
        token.cancel()
        token.cancel()
        token.cancel()
        assert token.is_cancelled is True

    def test_check_raises_after_cancel(self) -> None:
        token = CancellationToken()
        token.cancel()
        with pytest.raises(CancellationRequested):
            token.check()

    def test_check_does_not_raise_before_cancel(self) -> None:
        token = CancellationToken()
        token.check()  # should not raise

    def test_add_callback_invoked_on_cancel(self) -> None:
        token = CancellationToken()
        called: list[str] = []

        def cb() -> None:
            called.append("fired")

        token.add_callback(cb)
        token.cancel()
        assert called == ["fired"]

    def test_add_callback_when_already_cancelled(self) -> None:
        token = CancellationToken()
        token.cancel()
        called: list[str] = []

        def cb() -> None:
            called.append("fired")

        token.add_callback(cb)
        assert called == ["fired"]  # called immediately

    def test_callback_exception_does_not_break_cancel(self) -> None:
        token = CancellationToken()

        def bad_cb() -> None:
            raise RuntimeError("oops")

        called: list[str] = []

        def good_cb() -> None:
            called.append("good")

        token.add_callback(bad_cb)
        token.add_callback(good_cb)
        token.cancel()
        assert called == ["good"]


class TestCancellationTokenView:
    def test_view_reflects_token_state(self) -> None:
        token = CancellationToken()
        view = token.view()
        assert view.is_cancelled is False
        token.cancel()
        assert view.is_cancelled is True

    def test_view_check_raises(self) -> None:
        token = CancellationToken()
        view = token.view()
        token.cancel()
        with pytest.raises(CancellationRequested):
            view.check()

    def test_view_has_no_cancel_method(self) -> None:
        token = CancellationToken()
        view = token.view()
        assert not hasattr(view, "cancel")


class TestCancellationTokenThreadSafety:
    def test_concurrent_cancel_and_check(self) -> None:
        """Concurrent cancel and is_cancelled checks should not crash.

        The checker reads ``is_cancelled`` in a tight loop; the canceller
        sets the flag from another thread.  The token's internal lock must
        prevent data races.
        """
        token = CancellationToken()
        errors: list[Exception] = []

        def canceller() -> None:
            token.cancel()

        def checker() -> None:
            try:
                for _ in range(100000):
                    _ = token.is_cancelled
            except Exception as e:
                errors.append(e)  # pragma: no cover

        t1 = threading.Thread(target=canceller)
        t2 = threading.Thread(target=checker)
        t2.start()
        import time

        time.sleep(0.005)
        t1.start()
        t1.join()
        t2.join()
        # Should not crash with a data race
        assert len(errors) == 0
