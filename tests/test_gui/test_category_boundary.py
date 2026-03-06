"""GUI 逻辑单元测试。"""

from __future__ import annotations

import threading

import pytest

from docwen.gui.core.logic import MainWindowLogic
from docwen.services.result import ConversionResult

pytestmark = [pytest.mark.unit, pytest.mark.windows_only]


class _Root:
    def after(self, _ms, callback=None, *args):
        if callback is None:
            return None
        return callback(*args)


class _MainWindow:
    def __init__(self) -> None:
        self.root = _Root()
        self.selected_template = ("x", "x")
        self.batch_panel_visible = False
        self.tabbed_file_manager = None

    def set_transient_status(self, *_a, **_k):
        return None


def test_gui_strategy_lookup_uses_actual_format_not_text(monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
    f = tmp_path / "a.txt"
    f.write_text("x", encoding="utf-8")

    monkeypatch.setattr(
        "docwen.utils.file_type_utils.detect_actual_file_format",
        lambda *_a, **_k: "txt",
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.utils.file_type_utils.get_actual_file_category",
        lambda *_a, **_k: "text",
        raising=True,
    )

    def _fake_get_strategy(*, action_type=None, source_format=None, target_format=None):
        assert source_format == "txt"
        assert source_format != "text"

        class _Strategy:
            def execute(self, file_path: str, options=None, progress_callback=None):
                return ConversionResult.ok(message="ok", output_path=file_path)

        return _Strategy

    monkeypatch.setattr("docwen.services.strategies.get_strategy", _fake_get_strategy, raising=True)

    logic = MainWindowLogic(_MainWindow())
    cancel_event = threading.Event()
    logic._run_in_background(
        operation_id="x",
        action_type="convert",
        file_path=str(f),
        options={"target_format": "docx", "template_name": "x"},
        cancel_event=cancel_event,
    )
