from __future__ import annotations

from typing import Any, Protocol, cast

import ttkbootstrap as tb


class HasStyle(Protocol):
    style: tb.Style


def attach_style(root: Any, style: tb.Style) -> HasStyle:
    root.style = style
    return cast(HasStyle, root)


def as_tb_window(root: Any) -> tb.Window:
    return cast(tb.Window, root)
