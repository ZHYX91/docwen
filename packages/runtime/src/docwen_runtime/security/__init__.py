"""Runtime security controls.

Startup security: strict mode resolution, protection chain execution,
and network isolation setup.  These behaviours live here (runtime layer)
rather than being scattered across CLI or plugin code.
"""

from __future__ import annotations

import logging
import os

from docwen_runtime.security.network import (
    NetworkAccessBlockedError,
    NetworkGuardInstallationError,
    NetworkGuardStatus,
    activate_process_lifetime_dependency_egress_guard,
    dependency_egress_guard,
    dependency_egress_guard_status,
    install_dependency_egress_guard,
)

__all__ = [
    "NetworkAccessBlockedError",
    "NetworkGuardInstallationError",
    "NetworkGuardStatus",
    "activate_process_lifetime_dependency_egress_guard",
    "dependency_egress_guard",
    "dependency_egress_guard_status",
    "initialize_network_isolation",
    "install_dependency_egress_guard",
    "resolve_strict_security",
    "run_security_protections",
]


# ── Strict security resolution ────────────────────────────────────────


def resolve_strict_security(strict_security: bool | None = None) -> bool:
    """Resolve the strict security mode.

    Priority:
    1. Explicit *strict_security* argument.
    2. ``DOCWEN_STRICT_SECURITY`` environment variable (defaults to ``"1"``).

    Returns:
        ``True`` when strict security mode is enabled.
    """
    if strict_security is not None:
        return strict_security
    env_value = os.environ.get("DOCWEN_STRICT_SECURITY", "1").strip().lower()
    return env_value not in {"0", "false", "no", "off"}


# ── Protection chain ──────────────────────────────────────────────────


def _check_debugger() -> None:
    """Detect attached debuggers and exit immediately if one is found.

    Mirrors the old ``docwen.security.protection_utils.check_debugger()``
    behaviour: uses ``sys.gettrace()`` plus Windows ``IsDebuggerPresent``.
    """
    import os as _os
    import sys as _sys

    def _is_authorized_pytest_coverage_trace(trace: object) -> bool:
        """Accept only the repository's explicit pytest coverage process.

        Coverage uses Python's tracing API, which is otherwise intentionally
        indistinguishable from a debugger here.  The exception remains
        unavailable to frozen builds and requires the CI/local coverage job,
        an actively running pytest test, and coverage.py's own tracer.
        """
        trace_module = type(trace).__module__
        return (
            not getattr(_sys, "frozen", False)
            and _os.environ.get("DOCWEN_TEST_ALLOW_COVERAGE_TRACE") == "1"
            and bool(_os.environ.get("PYTEST_CURRENT_TEST"))
            and "pytest" in _sys.modules
            and "coverage" in _sys.modules
            and (trace_module == "coverage" or trace_module.startswith("coverage."))
        )

    try:
        trace = _sys.gettrace()
        if trace is not None and not _is_authorized_pytest_coverage_trace(trace):
            _os._exit(1)

        if _sys.platform == "win32":
            import ctypes as _ctypes

            if _ctypes.windll.kernel32.IsDebuggerPresent():
                _os._exit(1)
    except Exception:
        _os._exit(1)


def _run_all_protections() -> None:
    """Run all protection checks (mirrors old ``run_all_protections``)."""
    _check_debugger()


def run_security_protections(
    *,
    logger: logging.Logger | None = None,
    strict_security: bool | None = None,
) -> str | None:
    """Run the startup protection chain.

    In strict mode, a failure raises :exc:`SecurityCheckFailedError` and
    blocks startup.  In non-strict (degraded) mode the function returns a
    diagnostic message so callers can log or display it, but startup
    continues.

    Args:
        logger: Logger for diagnostics.  Uses a module-level logger if
            *None*.
        strict_security: Override for :func:`resolve_strict_security`.
            When *None* the env var ``DOCWEN_STRICT_SECURITY`` is checked.

    Returns:
        *None* when all protections pass, or a diagnostic string when
        running in degraded (non-strict) mode.

    Raises:
        SecurityCheckFailedError: In strict mode when a protection fails.
    """
    if logger is None:
        logger = logging.getLogger(__name__)

    from docwen_runtime.errors import SecurityCheckFailedError

    strict = resolve_strict_security(strict_security)

    try:
        _run_all_protections()
        logger.info("核心安全防护检查完成。")
        return None
    except SystemExit:
        # _check_debugger calls os._exit; if we somehow reach here,
        # re-raise so the process terminates.
        raise
    except ImportError as exc:
        logger.critical("无法导入安全模块: %s", exc, exc_info=True)
        if strict:
            logger.critical("严格安全模式已启用，程序将终止。")
            _msg = str(exc)
            raise SecurityCheckFailedError(details=_msg, cause=exc) from exc
        return f"无法导入安全模块，已降级运行: {exc}"
    except Exception as exc:
        logger.critical("核心安全检查失败: %s", exc, exc_info=True)
        if strict:
            logger.critical("严格安全模式已启用，程序将终止。")
            _msg = str(exc)
            raise SecurityCheckFailedError(details=_msg, cause=exc) from exc
        return f"核心安全检查失败，已降级运行: {exc}"


# ── Network isolation ─────────────────────────────────────────────────


def initialize_network_isolation(logger: logging.Logger | None = None) -> None:
    """Install and verify the process-wide dependency egress audit hook.

    The hook becomes enforcing only while :func:`dependency_egress_guard` is
    active.  Supported CLI and GUI entry points own that lifecycle.
    """
    if logger is None:
        logger = logging.getLogger(__name__)
    try:
        status = install_dependency_egress_guard()
        if not status.installed:
            raise NetworkGuardInstallationError()
        logger.info("第三方依赖出站保护已安装。")
    except NetworkGuardInstallationError:
        raise
    except Exception as exc:
        logger.critical("第三方依赖出站保护安装失败: %s", exc, exc_info=True)
        raise NetworkGuardInstallationError(cause=exc) from exc
