from __future__ import annotations

from typing import Any

class MessageDialog:
    def __init__(
        self,
        parent: Any = ...,
        title: str = ...,
        message: str = ...,
        alert: bool = ...,
        localize: bool = ...,
        **kwargs: Any,
    ) -> None: ...
    def show(self) -> Any: ...
