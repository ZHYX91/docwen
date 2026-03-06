"""
ConversionResult 的结构化错误归一与用户提示生成

约定：
- 策略层优先通过 ConversionResult.fail(error_code/details) 传递结构化错误
- 若 error 是 DocWenError，则以其 code/user_message/details 补齐缺失字段
"""

from __future__ import annotations

from docwen.errors import DocWenError
from docwen.services.result import ConversionResult
from docwen.services.error_codes import ERROR_CODE_UNKNOWN_ERROR


def normalize_result_error_fields(result: ConversionResult) -> ConversionResult:
    if result.success:
        return result

    if isinstance(result.error, DocWenError):
        error_code = result.error_code or result.error.code
        details = result.details or result.error.details
        message = result.message or result.error.user_message
        if error_code != result.error_code or details != result.details or message != result.message:
            return result._replace(message=message, error_code=error_code, details=details)

    if not result.error_code:
        details = result.details
        if not details and result.error:
            details = str(result.error) or None
        return result._replace(error_code=ERROR_CODE_UNKNOWN_ERROR, details=details)

    return result


def format_result_message(result: ConversionResult, default_message: str) -> str:
    normalized = normalize_result_error_fields(result)
    message = normalized.message or default_message
    details = normalized.details
    if details:
        return f"{message} ({details})"
    return message


def format_exception_message(exc: BaseException, default_message: str) -> str:
    if isinstance(exc, DocWenError):
        if exc.details:
            return f"{exc.user_message} ({exc.details})"
        return exc.user_message
    return str(exc) or default_message
