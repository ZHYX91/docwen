"""Contract tests: verify test fakes match production protocol interfaces.

These tests break intentionally when a production protocol changes but the
corresponding fake wasn't updated. This prevents interface drift between
test doubles and real implementations.
"""

from __future__ import annotations

import pytest
from tests.support.cancellation import FakeCancellationTokenView
from tests.support.config import FakeConfigView
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

pytestmark = pytest.mark.contract


def _public_methods(cls: type) -> set[str]:
    """Return the set of non-dunder callable names declared on *cls*."""
    return {m for m in dir(cls) if not m.startswith("_") and callable(getattr(cls, m, None))}


class TestFakePluginLogger:
    def test_matches_plugin_logger_interface(self) -> None:
        from docwen_core.protocols.execution_context import PluginLogger

        fake = _public_methods(FakePluginLogger)
        real = _public_methods(PluginLogger)
        missing = real - fake
        assert missing == set(), f"FakePluginLogger missing: {missing}"


class TestFakeProgressSink:
    def test_matches_progress_sink_interface(self) -> None:
        from docwen_core.protocols.execution_context import ProgressSink

        fake = _public_methods(FakeProgressSink)
        real = _public_methods(ProgressSink)
        missing = real - fake
        assert missing == set(), f"FakeProgressSink missing: {missing}"


class TestFakeConfigView:
    def test_matches_read_only_config_view_interface(self) -> None:
        from docwen_core.protocols.execution_context import ReadOnlyConfigView

        fake = _public_methods(FakeConfigView)
        real = _public_methods(ReadOnlyConfigView)
        missing = real - fake
        assert missing == set(), f"FakeConfigView missing: {missing}"


class TestFakeCancellationTokenView:
    def test_matches_cancellation_token_view_interface(self) -> None:
        from docwen_core.protocols.execution_context import CancellationTokenView

        fake = _public_methods(FakeCancellationTokenView)
        real = _public_methods(CancellationTokenView)
        missing = real - fake
        assert missing == set(), f"FakeCancellationTokenView missing: {missing}"


class TestFakeWorkspaceHandle:
    def test_matches_workspace_handle_interface(self) -> None:
        from docwen_core.protocols.execution_context import WorkspaceHandle

        fake = _public_methods(FakeWorkspaceHandle)
        real = _public_methods(WorkspaceHandle)
        missing = real - fake
        assert missing == set(), f"FakeWorkspaceHandle missing: {missing}"
