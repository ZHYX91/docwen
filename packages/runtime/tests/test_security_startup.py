"""Tests for runtime security startup: strict security resolution,
protection chain, failure labels, and exit codes.
"""

from __future__ import annotations

import logging
import os
import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def _run_security_probe(
    source: str, *, env_overrides: dict[str, str] | None = None
) -> subprocess.CompletedProcess[str]:
    runtime_src = Path(__file__).resolve().parents[1] / "src"
    probe = f"import sys\nsys.path.insert(0, {str(runtime_src)!r})\n{source}"
    env = os.environ.copy()
    if env_overrides:
        env.update(env_overrides)
    return subprocess.run(
        [sys.executable, "-c", probe],
        capture_output=True,
        env=env,
        text=True,
        timeout=10,
        check=False,
    )


# ── resolve_strict_security ───────────────────────────────────────────


class TestResolveStrictSecurity:
    """F-B2-003: DOCWEN_STRICT_SECURITY env var resolution."""

    def test_explicit_true_overrides_env(self) -> None:
        from docwen_runtime.security import resolve_strict_security

        # explicit arg wins over any env value
        assert resolve_strict_security(True) is True
        assert resolve_strict_security(False) is False

    def test_env_defaults_to_strict_when_unset(self, monkeypatch) -> None:
        from docwen_runtime.security import resolve_strict_security

        monkeypatch.delenv("DOCWEN_STRICT_SECURITY", raising=False)
        assert resolve_strict_security() is True

    def test_env_explicit_1(self, monkeypatch) -> None:
        from docwen_runtime.security import resolve_strict_security

        monkeypatch.setenv("DOCWEN_STRICT_SECURITY", "1")
        assert resolve_strict_security() is True

    def test_env_explicit_0(self, monkeypatch) -> None:
        from docwen_runtime.security import resolve_strict_security

        monkeypatch.setenv("DOCWEN_STRICT_SECURITY", "0")
        assert resolve_strict_security() is False

    def test_env_false_string(self, monkeypatch) -> None:
        from docwen_runtime.security import resolve_strict_security

        for val in ("false", "False", "FALSE", "no", "No", "NO", "off", "Off", "OFF"):
            monkeypatch.setenv("DOCWEN_STRICT_SECURITY", val)
            assert resolve_strict_security() is False, f"value {val!r} should be false"

    def test_env_unknown_string_is_strict(self, monkeypatch) -> None:
        from docwen_runtime.security import resolve_strict_security

        monkeypatch.setenv("DOCWEN_STRICT_SECURITY", "yes")
        assert resolve_strict_security() is True

    def test_env_whitespace_handled(self, monkeypatch) -> None:
        from docwen_runtime.security import resolve_strict_security

        monkeypatch.setenv("DOCWEN_STRICT_SECURITY", " 0 ")
        assert resolve_strict_security() is False


# ── run_security_protections - non-strict mode ────────────────────────


class TestRunSecurityProtectionsNonStrict:
    """F-B2-004: protection chain in non-strict (degraded) mode."""

    def test_returns_none_when_protections_pass(self) -> None:
        from docwen_runtime.security import run_security_protections

        result = run_security_protections(strict_security=False)
        # No debugger attached in test → protections pass
        assert result is None

    def test_returns_degraded_message_on_import_error(self, monkeypatch) -> None:
        """Simulate a missing protection module by patching _run_all_protections."""
        from docwen_runtime.security import run_security_protections

        # Patch _run_all_protections to raise ImportError
        monkeypatch.setattr(
            "docwen_runtime.security._run_all_protections",
            lambda: (_ for _ in ()).throw(ImportError("no module named foo")),
        )
        result = run_security_protections(strict_security=False)
        # Non-strict: ImportError is NOT silently treated as pass — degraded msg is returned
        assert result is not None
        assert "降级运行" in result
        assert "no module named foo" in result

    def test_returns_degraded_message_on_generic_error_non_strict(self, monkeypatch) -> None:
        from docwen_runtime.security import run_security_protections

        monkeypatch.setattr(
            "docwen_runtime.security._run_all_protections",
            lambda: (_ for _ in ()).throw(RuntimeError("simulated failure")),
        )
        result = run_security_protections(strict_security=False)
        assert result is not None
        assert "降级运行" in result
        assert "simulated failure" in result


class TestDebuggerProtectionSubprocess:
    def test_python_trace_detection_terminates_process(self) -> None:
        result = _run_security_probe(
            "from docwen_runtime.security import _check_debugger\n"
            "sys.gettrace = lambda: object()\n"
            "_check_debugger()\n"
            "raise SystemExit(99)\n"
        )

        assert result.returncode == 1

    def test_coverage_override_alone_does_not_disable_trace_detection(self) -> None:
        result = _run_security_probe(
            "from docwen_runtime.security import _check_debugger\n"
            "sys.gettrace = lambda: object()\n"
            "_check_debugger()\n"
            "raise SystemExit(99)\n",
            env_overrides={"DOCWEN_TEST_ALLOW_COVERAGE_TRACE": "1"},
        )

        assert result.returncode == 1

    def test_explicit_pytest_coverage_trace_is_accepted_for_source_tests(self) -> None:
        result = _run_security_probe(
            "import types\n"
            "from docwen_runtime.security import _check_debugger\n"
            "CoverageTracer = type('CoverageTracer', (), {'__module__': 'coverage'})\n"
            "sys.modules['pytest'] = types.ModuleType('pytest')\n"
            "sys.modules['coverage'] = types.ModuleType('coverage')\n"
            "sys.gettrace = lambda: CoverageTracer()\n"
            "_check_debugger()\n"
            "raise SystemExit(99)\n",
            env_overrides={
                "DOCWEN_TEST_ALLOW_COVERAGE_TRACE": "1",
                "PYTEST_CURRENT_TEST": "security coverage probe (call)",
            },
        )

        assert result.returncode == 99

    def test_frozen_process_rejects_explicit_pytest_coverage_trace(self) -> None:
        result = _run_security_probe(
            "import types\n"
            "from docwen_runtime.security import _check_debugger\n"
            "CoverageTracer = type('CoverageTracer', (), {'__module__': 'coverage'})\n"
            "sys.modules['pytest'] = types.ModuleType('pytest')\n"
            "sys.modules['coverage'] = types.ModuleType('coverage')\n"
            "sys.frozen = True\n"
            "sys.gettrace = lambda: CoverageTracer()\n"
            "_check_debugger()\n"
            "raise SystemExit(99)\n",
            env_overrides={
                "DOCWEN_TEST_ALLOW_COVERAGE_TRACE": "1",
                "PYTEST_CURRENT_TEST": "security frozen coverage probe (call)",
            },
        )

        assert result.returncode == 1

    def test_windows_debugger_detection_terminates_full_protection_chain(self) -> None:
        result = _run_security_probe(
            "import ctypes, types\n"
            "from docwen_runtime.security import _run_all_protections\n"
            "sys.gettrace = lambda: None\n"
            "sys.platform = 'win32'\n"
            "ctypes.windll = types.SimpleNamespace(\n"
            "    kernel32=types.SimpleNamespace(IsDebuggerPresent=lambda: 1)\n"
            ")\n"
            "_run_all_protections()\n"
            "raise SystemExit(99)\n"
        )

        assert result.returncode == 1


# ── run_security_protections - strict mode ────────────────────────────


class TestRunSecurityProtectionsStrict:
    """F-B2-004: protection chain in strict mode blocks startup."""

    def test_returns_none_when_protections_pass_strict(self) -> None:
        from docwen_runtime.security import run_security_protections

        result = run_security_protections(strict_security=True)
        assert result is None

    def test_raises_security_check_failed_on_import_error_strict(self, monkeypatch) -> None:
        from docwen_runtime.errors import SecurityCheckFailedError
        from docwen_runtime.security import run_security_protections

        monkeypatch.setattr(
            "docwen_runtime.security._run_all_protections",
            lambda: (_ for _ in ()).throw(ImportError("no module named security_utils")),
        )
        with pytest.raises(SecurityCheckFailedError) as exc_info:
            run_security_protections(strict_security=True)
        assert exc_info.value.code == "security_check_failed"
        assert "no module named security_utils" in (exc_info.value.details or "")

    def test_raises_security_check_failed_on_error_strict(self, monkeypatch) -> None:
        from docwen_runtime.errors import SecurityCheckFailedError
        from docwen_runtime.security import run_security_protections

        monkeypatch.setattr(
            "docwen_runtime.security._run_all_protections",
            lambda: (_ for _ in ()).throw(RuntimeError("strict mode simulated failure")),
        )
        with pytest.raises(SecurityCheckFailedError) as exc_info:
            run_security_protections(strict_security=True)
        assert exc_info.value.code == "security_check_failed"
        assert "strict mode simulated failure" in (exc_info.value.details or "")

    def test_respects_env_var_strict(self, monkeypatch) -> None:
        from docwen_runtime.errors import SecurityCheckFailedError
        from docwen_runtime.security import run_security_protections

        monkeypatch.setenv("DOCWEN_STRICT_SECURITY", "1")
        monkeypatch.setattr(
            "docwen_runtime.security._run_all_protections",
            lambda: (_ for _ in ()).throw(RuntimeError("env strict failure")),
        )
        with pytest.raises(SecurityCheckFailedError) as exc_info:
            run_security_protections()  # env defaults to strict
        assert exc_info.value.code == "security_check_failed"


# ── SecurityCheckFailedError ──────────────────────────────────────────


class TestSecurityCheckFailedError:
    def test_has_code_attribute(self) -> None:
        from docwen_runtime.errors import SecurityCheckFailedError

        exc = SecurityCheckFailedError(details="test detail")
        assert exc.code == "security_check_failed"

    def test_stores_details(self) -> None:
        from docwen_runtime.errors import SecurityCheckFailedError

        exc = SecurityCheckFailedError(details="some detail")
        assert exc.details == "some detail"

    def test_str_uses_details(self) -> None:
        from docwen_runtime.errors import SecurityCheckFailedError

        exc = SecurityCheckFailedError(details="test")
        assert "test" in str(exc)

    def test_str_defaults_when_no_details(self) -> None:
        from docwen_runtime.errors import SecurityCheckFailedError

        exc = SecurityCheckFailedError()
        assert "核心安全检查失败" in str(exc)

    def test_sets_cause(self) -> None:
        from docwen_runtime.errors import SecurityCheckFailedError

        cause = RuntimeError("root cause")
        exc = SecurityCheckFailedError(details="boom", cause=cause)
        assert exc.__cause__ is cause

    def test_is_docwen_error(self) -> None:
        from docwen_core.errors import DocWenError
        from docwen_runtime.errors import SecurityCheckFailedError

        assert issubclass(SecurityCheckFailedError, DocWenError)


# ── FailureCategory and FAILURE_LABELS ────────────────────────────────


class TestFailureLabels:
    """F-B1-004: Centralised Chinese failure category labels."""

    def test_all_categories_have_labels(self) -> None:
        from docwen_runtime.errors import FAILURE_LABELS, FailureCategory

        for cat in FailureCategory:
            assert cat in FAILURE_LABELS, f"{cat} missing from FAILURE_LABELS"
            assert isinstance(FAILURE_LABELS[cat], str)
            assert len(FAILURE_LABELS[cat]) > 0

    def test_all_categories_have_codes(self) -> None:
        from docwen_runtime.errors import FAILURE_CODES, FailureCategory

        for cat in FailureCategory:
            assert cat in FAILURE_CODES, f"{cat} missing from FAILURE_CODES"
            assert isinstance(FAILURE_CODES[cat], int)

    def test_security_check_label(self) -> None:
        from docwen_runtime.errors import FAILURE_LABELS, FailureCategory

        assert FAILURE_LABELS[FailureCategory.SECURITY_CHECK] == "安全检查失败"

    def test_initialization_label(self) -> None:
        from docwen_runtime.errors import FAILURE_LABELS, FailureCategory

        assert FAILURE_LABELS[FailureCategory.INITIALIZATION] == "初始化失败"

    def test_dependency_missing_label(self) -> None:
        from docwen_runtime.errors import FAILURE_LABELS, FailureCategory

        assert FAILURE_LABELS[FailureCategory.DEPENDENCY_MISSING] == "依赖缺失"

    def test_user_interrupt_label(self) -> None:
        from docwen_runtime.errors import FAILURE_LABELS, FailureCategory

        assert FAILURE_LABELS[FailureCategory.USER_INTERRUPT] == "用户中断"

    def test_logging_init_label(self) -> None:
        from docwen_runtime.errors import FAILURE_LABELS, FailureCategory

        assert FAILURE_LABELS[FailureCategory.LOGGING_INIT] == "日志初始化失败"

    def test_user_interrupt_exit_code(self) -> None:
        from docwen_runtime.errors import FAILURE_CODES, FailureCategory

        # SIGINT-like code
        assert FAILURE_CODES[FailureCategory.USER_INTERRUPT] == 130

    def test_labels_not_empty_strings(self) -> None:
        from docwen_runtime.errors import FAILURE_LABELS

        for label in FAILURE_LABELS.values():
            assert label.strip(), f"empty label found: {label!r}"

    def test_exactly_five_categories(self) -> None:
        from docwen_runtime.errors import FailureCategory

        names = list(FailureCategory)
        assert len(names) == 5, f"expected 5 categories, got {len(names)}: {names}"


# ── Dependency egress guard ───────────────────────────────────────────


class TestNetworkIsolationStillWorks:
    """Verify the audit guard blocks use without replacing socket types."""

    def test_dependency_egress_guard_blocks_ip_connect(self) -> None:
        import socket

        from docwen_runtime.security import NetworkAccessBlockedError, dependency_egress_guard

        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            with dependency_egress_guard(), pytest.raises(NetworkAccessBlockedError):
                sock.connect(("127.0.0.1", 9))
        finally:
            sock.close()

    def test_dependency_egress_guard_never_replaces_socket(self) -> None:
        import socket

        from docwen_runtime.security import dependency_egress_guard

        orig = socket.socket
        with dependency_egress_guard():
            pass
        assert socket.socket is orig


# ── initialize_network_isolation ──────────────────────────────────────


class TestInitializeNetworkIsolation:
    def test_succeeds_when_module_available(self, caplog) -> None:
        from docwen_runtime.security import initialize_network_isolation

        with caplog.at_level(logging.INFO, logger="docwen_runtime.security"):
            initialize_network_isolation()
        assert "第三方依赖出站保护已安装" in caplog.text
