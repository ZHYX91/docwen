from __future__ import annotations

import tkinter as tk
from typing import Any

DND_FILES: str

class DnDWrapper:
    def drop_target_register(self, *dnd_types: str) -> Any: ...
    def drop_target_unregister(self) -> Any: ...
    def dnd_bind(self, sequence: str, func: Any, add: Any = ...) -> Any: ...

class TkinterDnD:
    @staticmethod
    def Tk(*args: Any, **kwargs: Any) -> tk.Tk: ...
