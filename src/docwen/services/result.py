"""
服务层结果对象模块
定义了所有服务层操作返回的标准数据结构。
"""

from typing import Any, NamedTuple


class ConversionResult(NamedTuple):
    """
    一个标准化的转换结果容器。

    Attributes:
        success (bool): 操作是否成功。
        output_path (Optional[str]): 成功时生成的输出文件路径。
        message (Optional[str]): 附加的用户友好消息（成功或失败原因）。
        error (Optional[Exception]): 如果发生异常，则包含异常对象。
    """

    success: bool
    output_path: str | None = None
    message: str | None = None
    error: Exception | None = None
    error_code: str | None = None
    details: str | None = None
    metadata: dict[str, Any] | None = None

    @classmethod
    def ok(
        cls,
        *,
        output_path: str | None = None,
        message: str | None = None,
        error_code: str | None = None,
        details: str | None = None,
        metadata: dict[str, Any] | None = None,
    ) -> "ConversionResult":
        return cls(
            success=True,
            output_path=output_path,
            message=message,
            error_code=error_code,
            details=details,
            metadata=metadata,
        )

    @classmethod
    def fail(
        cls,
        message: str | None = None,
        *,
        error: object | None = None,
        error_code: str,
        details: str | None = None,
        metadata: dict[str, Any] | None = None,
    ) -> "ConversionResult":
        if not error_code:
            raise ValueError("error_code is required for a failed ConversionResult")
        if error is not None and not isinstance(error, Exception):
            raise TypeError("error must be an Exception or None")
        return cls(
            success=False, message=message, error=error, error_code=error_code, details=details, metadata=metadata
        )
