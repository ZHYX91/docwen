from __future__ import annotations

from tkinter import Canvas as _TkCanvas
from tkinter import Text as _TkText
from tkinter import Tk
from tkinter import Toplevel as _TkToplevel
from tkinter import Variable as _TkVariable
from tkinter.ttk import Button as _TtkButton
from tkinter.ttk import Checkbutton as _TtkCheckbutton
from tkinter.ttk import Combobox as _TtkCombobox
from tkinter.ttk import Entry as _TtkEntry
from tkinter.ttk import Frame as _TtkFrame
from tkinter.ttk import Label as _TtkLabel
from tkinter.ttk import Labelframe as _TtkLabelframe
from tkinter.ttk import Notebook as _TtkNotebook
from tkinter.ttk import Radiobutton as _TtkRadiobutton
from tkinter.ttk import Scale as _TtkScale
from tkinter.ttk import Scrollbar as _TtkScrollbar
from tkinter.ttk import Separator as _TtkSeparator
from tkinter.ttk import Spinbox as _TtkSpinbox
from typing import Any

class Colors:
    bg: str
    fg: str
    info: str
    success: str
    warning: str

class Style:
    colors: Colors

    def __init__(self, theme: str | None = ...) -> None: ...
    def configure(self, style: str, cnf: Any = ..., **kw: Any) -> Any: ...
    def theme_use(self, theme: str | None = ...) -> str: ...
    @classmethod
    def get_instance(cls) -> Style: ...

class Window(Tk):
    style: Style
    def __init__(self, *args: Any, **kwargs: Any) -> None: ...

class Label(_TtkLabel):
    image: Any
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Frame(_TtkFrame):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Button(_TtkButton):
    image: Any
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Canvas(_TkCanvas):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...

class Labelframe(_TtkLabelframe):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Checkbutton(_TtkCheckbutton):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Combobox(_TtkCombobox):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Entry(_TtkEntry):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Text(_TkText):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Scrollbar(_TtkScrollbar):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Notebook(_TtkNotebook):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Radiobutton(_TtkRadiobutton):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Spinbox(_TtkSpinbox):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Separator(_TtkSeparator):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Scale(_TtkScale):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...
    def configure(self, cnf: Any = ..., **kw: Any) -> Any: ...
    config = configure

class Toplevel(_TkToplevel):
    def __init__(self, master: Any = ..., **kwargs: Any) -> None: ...

class StringVar(_TkVariable):
    def __init__(self, master: Any = ..., value: str | None = ..., name: str | None = ...) -> None: ...

class IntVar(_TkVariable):
    def __init__(self, master: Any = ..., value: int | None = ..., name: str | None = ...) -> None: ...

class BooleanVar(_TkVariable):
    def __init__(self, master: Any = ..., value: bool | None = ..., name: str | None = ...) -> None: ...
