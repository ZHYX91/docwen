"""Tests for bundle/gui_bootstrap.py.

Tests cover argument parsing, control-disable logic, and the bootstrap decision
flow. Real file locks are avoided by setting control-disable flags or using
unique app names with cleanup.
"""

from __future__ import annotations

import os
import time
from collections.abc import Generator
from pathlib import Path
from unittest.mock import patch

import pytest

from docwen_bundle.gui_bootstrap import (
    ENV_DISABLE_CONTROL,
    ENV_TEST_AUTOCLOSE_MS,
    GuiBootstrapDecision,
    bootstrap_gui,
    parse_startup_files,
    should_disable_control,
)

pytestmark = pytest.mark.integration

# ── Fixtures ────────────────────────────────────────────────────────────


@pytest.fixture(autouse=True)
def _clean_env_and_control() -> Generator[None, None, None]:
    """Remove IPC-related env vars and reset IPC toggle before each test.

    Both the bootstrap-layer env vars AND the runtime-layer
    ``disable_ipc()`` global must be reset for proper test isolation.
    """
    for key in (ENV_DISABLE_CONTROL, ENV_TEST_AUTOCLOSE_MS, "DEBUGPY_LAUNCHER_PORT"):
        os.environ.pop(key, None)
    from docwen_runtime.ipc import enable_ipc

    enable_ipc()
    yield
    for key in (ENV_DISABLE_CONTROL, ENV_TEST_AUTOCLOSE_MS, "DEBUGPY_LAUNCHER_PORT"):
        os.environ.pop(key, None)
    enable_ipc()


# ── parse_startup_files tests ───────────────────────────────────────────


class TestParseStartupFiles:
    """Tests for CLI argument parsing."""

    def test_no_args_returns_empty(self) -> None:
        """Just the script name — no files."""
        result = parse_startup_files(["script.py"])
        assert result == []

    def test_empty_argv_returns_empty(self) -> None:
        """Empty argv should return empty list."""
        result = parse_startup_files([])
        assert result == []

    def test_flags_are_skipped(self) -> None:
        """Arguments starting with - are treated as flags."""
        result = parse_startup_files(["script.py", "--help", "-v", "-style", "fusion"])
        assert result == []

    def test_nonexistent_file_is_ignored(self) -> None:
        """Non-existent paths are skipped (logged at debug)."""
        result = parse_startup_files(["script.py", "/nonexistent/path_12345.xyz"])
        assert result == []

    def test_existing_file_is_returned_as_absolute(self) -> None:
        """An existing file should be resolved to an absolute path."""
        # Use the test file itself as an existing file
        test_file = Path(__file__)
        result = parse_startup_files(["script.py", str(test_file)])
        assert len(result) == 1
        assert result[0] == str(test_file.resolve())

    def test_multiple_files(self) -> None:
        """Multiple existing files should all be returned."""
        test_file = Path(__file__)
        result = parse_startup_files(["script.py", str(test_file), str(test_file)])
        assert len(result) == 2

    def test_uses_sys_argv_by_default(self) -> None:
        """Should default to sys.argv when no argv provided."""
        # We can't easily control sys.argv in pytest, but we can verify
        # the function signature works.
        result = parse_startup_files()  # uses sys.argv
        assert isinstance(result, list)


# ── should_disable_control tests ────────────────────────────────────────


class TestShouldDisableControl:
    """Tests for the IPC disable decision logic."""

    def test_default_allows_ipc(self) -> None:
        """By default (no env vars, no debugger), IPC should be enabled."""
        assert not should_disable_control()

    def test_env_var_disables_ipc(self) -> None:
        """DOCWEN_GUI_DISABLE_CONTROL=1 should disable control."""
        os.environ[ENV_DISABLE_CONTROL] = "1"
        assert should_disable_control()

    def test_env_var_truthy_values(self) -> None:
        """Various truthy values should all disable IPC."""
        for val in ("1", "true", "True", "yes", "YES", "on", "ON", "y", "Y"):
            os.environ[ENV_DISABLE_CONTROL] = val
            assert should_disable_control(), f"Value {val!r} should disable control"

    def test_env_var_falsy_values(self) -> None:
        """Falsy values should not disable IPC."""
        for val in ("0", "false", "no", "off", "", "maybe"):
            os.environ[ENV_DISABLE_CONTROL] = val
            assert not should_disable_control(), f"Value {val!r} should not disable control"

    def test_test_autoclose_disables_ipc(self) -> None:
        """Positive DOCWEN_GUI_TEST_AUTOCLOSE_MS should disable IPC."""
        os.environ[ENV_TEST_AUTOCLOSE_MS] = "5000"
        assert should_disable_control()

    def test_test_autoclose_zero_does_not_disable(self) -> None:
        """DOCWEN_GUI_TEST_AUTOCLOSE_MS=0 should not disable IPC."""
        os.environ[ENV_TEST_AUTOCLOSE_MS] = "0"
        assert not should_disable_control()

    def test_debugpy_port_disables_ipc(self) -> None:
        """DEBUGPY_LAUNCHER_PORT should disable IPC."""
        os.environ["DEBUGPY_LAUNCHER_PORT"] = "5678"
        assert should_disable_control()

    @patch("sys.gettrace", return_value=lambda frame: frame)
    def test_debugger_attached_disables_ipc(self, _mock_gettrace) -> None:
        """Active debugger (sys.gettrace() is not None) should disable IPC."""
        # Note: pytest itself has a trace function, so this may already be true.
        # We test the logic directly via the helper.
        assert should_disable_control()


# ── GuiBootstrapDecision tests ─────────────────────────────────────────


class TestGuiBootstrapDecision:
    """Tests for the GUI bootstrap decision."""

    def test_defaults(self) -> None:
        """Default decision should start GUI, not exit."""
        d = GuiBootstrapDecision()
        assert d.should_start_gui is True
        assert d.should_exit is False
        assert d.exit_code == 0
        assert d.files_to_add == []
        assert d.instance_lock is None

    def test_exit_decision(self) -> None:
        """Exit decision with non-zero code."""
        d = GuiBootstrapDecision(should_start_gui=False, should_exit=True, exit_code=42)
        assert d.should_start_gui is False
        assert d.should_exit is True
        assert d.exit_code == 42

    def test_files_passed_to_decision(self) -> None:
        """Files from argv should be in the decision."""
        d = GuiBootstrapDecision(files_to_add=["/a/b.pdf", "/c/d.docx"])
        assert len(d.files_to_add) == 2


# ── bootstrap_gui integration tests ─────────────────────────────────────


class TestBootstrapGui:
    """Integration tests for the bootstrap_gui function."""

    def test_forward_attempt_never_exceeds_outer_deadline(self, monkeypatch: pytest.MonkeyPatch) -> None:
        from docwen_bundle import gui_bootstrap

        calls: list[tuple[str, dict[str, object], float]] = []

        class _Client:
            def request(self, action: str, payload: dict[str, object], *, timeout: float) -> None:
                calls.append((action, payload, timeout))

        clock = iter((100.0, 100.8, 100.9))
        monkeypatch.setattr(gui_bootstrap.time, "monotonic", lambda: next(clock))

        gui_bootstrap._send_when_control_ready(
            _Client(),
            "activate",
            {},
            deadline=101.0,
        )

        assert len(calls) == 1
        action, payload, timeout = calls[0]
        assert action == "activate"
        assert payload == {"_deadline_monotonic": 101.0}
        assert timeout == pytest.approx(0.1)

    def test_bootstrap_with_ipc_disabled(self) -> None:
        """When IPC is disabled, bootstrap should return start_gui=True."""
        os.environ[ENV_DISABLE_CONTROL] = "1"
        decision = bootstrap_gui("test_bootstrap_disabled")
        assert decision.should_start_gui is True
        assert decision.should_exit is False
        assert decision.instance_lock is None

    def test_bootstrap_primary_instance(self) -> None:
        """First bootstrap should acquire lock and return start_gui=True."""
        decision = bootstrap_gui("test_bootstrap_primary")
        try:
            assert decision.should_start_gui is True
            assert decision.should_exit is False
            assert decision.instance_lock is not None
        finally:
            if decision.instance_lock:
                decision.instance_lock.release()

    def test_bootstrap_secondary_instance_forwards(self) -> None:
        """Second bootstrap activates the primary through runtime/control."""
        from docwen_runtime.control import ControlServer

        app_name = "test_bootstrap_secondary"
        received: list[tuple[str, dict[str, object]]] = []
        decision1 = bootstrap_gui(app_name, argv=["script.py"])
        server = ControlServer(
            lambda action, payload: received.append((action, payload)) or {"accepted": True},
            app_name=app_name,
        )
        server.start()
        try:
            assert decision1.should_start_gui is True
            decision2 = bootstrap_gui(app_name, argv=["script.py"])
            assert decision2.should_start_gui is False
            assert decision2.should_exit is True
            assert decision2.exit_code == 0
            assert decision2.instance_lock is None
            assert len(received) == 1
            action, payload = received[0]
            assert action == "activate"
            deadline = payload.get("_deadline_monotonic")
            assert isinstance(deadline, float)
            assert 0 < deadline - time.monotonic() <= 1.0
        finally:
            server.stop()
            if decision1.instance_lock:
                decision1.instance_lock.release()

    def test_bootstrap_secondary_with_files(self) -> None:
        """Secondary instance forwards canonical files through runtime/control."""
        from docwen_runtime.control import ControlServer

        test_file = str(Path(__file__).resolve())
        app_name = "test_bootstrap_files"
        received: list[tuple[str, dict[str, object]]] = []
        decision1 = bootstrap_gui(app_name, argv=["script.py"])
        server = ControlServer(
            lambda action, payload: received.append((action, payload)) or {"accepted": True},
            app_name=app_name,
        )
        server.start()
        try:
            assert decision1.should_start_gui is True
            decision2 = bootstrap_gui(
                app_name,
                argv=["script.py", test_file],
            )
            assert decision2.should_start_gui is False
            assert decision2.should_exit is True
            assert decision2.exit_code == 0
            assert len(received) == 1
            action, payload = received[0]
            assert action == "open"
            assert payload.get("file") == test_file
            deadline = payload.get("_deadline_monotonic")
            assert isinstance(deadline, float)
            assert 0 < deadline - time.monotonic() <= 1.0
        finally:
            server.stop()
            if decision1.instance_lock:
                decision1.instance_lock.release()
