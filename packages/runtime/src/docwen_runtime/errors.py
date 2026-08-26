"""Runtime error types and startup failure semantics.

Shared by CLI and GUI for startup error handling with Chinese labels.
This module provides the centralised failure category mapping that was
lost during the rewrite (F-B1-004).
"""

from __future__ import annotations

from enum import StrEnum

from docwen_core.errors import DocWenError

# ── Security check error ──────────────────────────────────────────────

_ERROR_CODE_SECURITY_CHECK_FAILED = "security_check_failed"


class SecurityCheckFailedError(DocWenError):
    """Raised when strict security mode blocks startup due to a protection failure.

    Attributes:
        code: Error code string for exit-code mapping.
        details: Diagnostic detail about which check failed.
    """

    def __init__(
        self,
        *,
        details: str | None = None,
        cause: BaseException | None = None,
    ) -> None:
        super().__init__(details or "核心安全检查失败")
        self.code: str = _ERROR_CODE_SECURITY_CHECK_FAILED
        self.details: str | None = details
        if cause is not None:
            self.__cause__ = cause


# ── Failure categories and Chinese labels ─────────────────────────────


class FailureCategory(StrEnum):
    """Well-known startup failure categories."""

    SECURITY_CHECK = "security_check"
    INITIALIZATION = "initialization"
    DEPENDENCY_MISSING = "dependency_missing"
    USER_INTERRUPT = "user_interrupt"
    LOGGING_INIT = "logging_init"


FAILURE_LABELS: dict[FailureCategory, str] = {
    FailureCategory.SECURITY_CHECK: "安全检查失败",
    FailureCategory.INITIALIZATION: "初始化失败",
    FailureCategory.DEPENDENCY_MISSING: "依赖缺失",
    FailureCategory.USER_INTERRUPT: "用户中断",
    FailureCategory.LOGGING_INIT: "日志初始化失败",
}

FAILURE_CODES: dict[FailureCategory, int] = {
    FailureCategory.SECURITY_CHECK: 1,
    FailureCategory.INITIALIZATION: 1,
    FailureCategory.DEPENDENCY_MISSING: 1,
    FailureCategory.USER_INTERRUPT: 130,
    FailureCategory.LOGGING_INIT: 1,
}
