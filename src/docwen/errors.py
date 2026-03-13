"""
DocWen 领域异常定义。
用于在 CLI/GU I/服务层之间传递一致的错误码与用户可读信息。
"""

from __future__ import annotations

from enum import IntEnum

from docwen.services.error_codes import (
    ERROR_CODE_DEPENDENCY_MISSING,
    ERROR_CODE_INVALID_INPUT,
    ERROR_CODE_SECURITY_CHECK_FAILED,
    ERROR_CODE_STRATEGY_NOT_FOUND,
)


class ExitCode(IntEnum):
    OK = 0
    UNKNOWN_ERROR = 1
    INVALID_INPUT = 2
    STRATEGY_NOT_FOUND = 3
    DEPENDENCY_MISSING = 4
    SECURITY_CHECK_FAILED = 5


def exit_code_from_error_code(error_code: str | None) -> ExitCode:
    from docwen.services.error_registry import get_error_definition

    definition = get_error_definition(error_code)
    try:
        return ExitCode(int(definition.exit_code))
    except Exception:
        return ExitCode.UNKNOWN_ERROR


class DocWenError(Exception):
    code: str
    user_message: str
    details: str | None

    def __init__(
        self,
        code: str,
        user_message: str,
        *,
        details: str | None = None,
        cause: BaseException | None = None,
    ) -> None:
        super().__init__(user_message)
        self.code = code
        self.user_message = user_message
        self.details = details
        if cause is not None:
            self.__cause__ = cause

    def __str__(self) -> str:
        if self.details:
            return f"{self.user_message} ({self.details})"
        return self.user_message


class InvalidInputError(DocWenError):
    def __init__(self, user_message: str, *, details: str | None = None, cause: BaseException | None = None) -> None:
        super().__init__(ERROR_CODE_INVALID_INPUT, user_message, details=details, cause=cause)


class DependencyMissingError(DocWenError):
    def __init__(self, user_message: str, *, details: str | None = None, cause: BaseException | None = None) -> None:
        super().__init__(ERROR_CODE_DEPENDENCY_MISSING, user_message, details=details, cause=cause)


class StrategyNotFoundError(DocWenError):
    def __init__(
        self,
        *,
        action_type: str | None = None,
        source_format: str | None = None,
        target_format: str | None = None,
    ) -> None:
        parts: list[str] = []
        if action_type:
            parts.append(f"action='{action_type}'")
        if source_format and target_format:
            parts.append(f"conversion='{source_format}->{target_format}'")
        details = " ".join(parts) if parts else None
        user_message = "没有找到对应的处理策略"
        super().__init__(ERROR_CODE_STRATEGY_NOT_FOUND, user_message, details=details)


class SecurityCheckFailedError(DocWenError):
    def __init__(self, *, details: str | None = None, cause: BaseException | None = None) -> None:
        super().__init__(ERROR_CODE_SECURITY_CHECK_FAILED, "核心安全检查失败", details=details, cause=cause)
