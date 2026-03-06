from __future__ import annotations

from tkinter.ttk import Treeview
from typing import Any

from . import Frame

class Tableview(Frame):
    view: Treeview

    def __init__(
        self,
        master: Any = ...,
        *,
        coldata: Any = ...,
        rowdata: Any = ...,
        paginated: bool = ...,
        searchable: bool = ...,
        bootstyle: str | None = ...,
        height: int | None = ...,
        pagesize: int | None = ...,
        **kwargs: Any,
    ) -> None: ...
    def delete_rows(self) -> None: ...
    def insert_row(self, index: str, values: list[Any]) -> None: ...
    def load_table_data(self) -> None: ...
    def get_rows(self, *, selected: bool = ...) -> list[Any]: ...
