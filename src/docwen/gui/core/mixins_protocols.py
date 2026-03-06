from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from typing import Any, Protocol

import ttkbootstrap as tb


class ActionPanelHost(Protocol):
    file_type: str | None
    file_path: str | None

    button_container: tb.Frame
    button_colors: dict[str, str]
    button_style_1: dict[str, object]
    button_style_2: dict[str, object]

    status_var: Any
    config_manager: Any
    on_action: Callable[..., Any] | None

    def __getattr__(self, name: str) -> Any: ...
    def __setattr__(self, name: str, value: Any) -> None: ...

    def _clear_buttons(self) -> None: ...

    def _clear_options(self) -> None: ...

    def clear_buttons(self) -> None: ...

    def clear_options(self) -> None: ...


class DragDropHost(Protocol):
    master: Any
    canvas: Any
    drag_enabled: bool
    is_dragging: bool

    def add_files(
        self, file_paths: list[str], auto_select_first: bool = ...
    ) -> tuple[list[str], list[tuple[str, str]]]: ...

    def on_drag_enter(self, event: Any) -> Any: ...
    def on_drag_leave(self, event: Any) -> Any: ...
    def on_drop(self, event: Any) -> Any: ...
    def get_tabbed_batch_file_list(self) -> TabbedFileSelectorHost | None: ...
    def parse_dropped_files(self, file_data: str) -> list[str]: ...
    def process_batch_files(self, files: list[str]) -> list[str]: ...


class ConversionPanelHost(Protocol):
    config_manager: Any
    on_action: Callable[..., Any] | None

    conversion_container: tb.Frame
    saveas_container: tb.Frame
    expand_container: tb.Frame
    extra_frame: Any
    extra_container: Any

    default_font: str
    default_size: int
    small_font: str
    small_size: int

    button_colors: dict[str, str]
    button_style_2col: dict[str, object]
    button_style_1col: dict[str, object]
    format_buttons: dict[str, tb.Button]

    checkbox_vars: dict[str, tk.BooleanVar]
    merge_mode_var: tk.IntVar | None
    reference_table_var: tk.StringVar | None
    current_file_path: str | None
    validate_button: tb.Button | None

    def __getattr__(self, name: str) -> Any: ...
    def __setattr__(self, name: str, value: Any) -> None: ...

    def on_format_clicked(self, fmt: str) -> None: ...
    def update_button_states(self) -> None: ...


class DecorationHost(Protocol):
    decoration_canvas: tk.Canvas

    def __getattr__(self, name: str) -> Any: ...
    def __setattr__(self, name: str, value: Any) -> None: ...

    def after(self, ms: int, func: Callable[[], Any]) -> str: ...

    def winfo_toplevel(self) -> Any: ...


class TabbedFileSelectorHost(Protocol):
    def add_files(self, file_paths: list[str]) -> tuple[list[str], list[tuple[str, str]]]: ...
