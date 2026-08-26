from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace
from typing import Any, cast

import pytest

pytestmark = pytest.mark.gui


def test_ocr_probe_stops_when_output_config_batch_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from PySide6.QtCore import QTimer

    from docwen_gui.app import _schedule_test_ocr_report

    source = tmp_path / "source.png"
    source.write_bytes(b"not-needed-after-config-failure")
    report = tmp_path / "report.json"
    output_dir = tmp_path / "outputs"
    monkeypatch.setenv("DOCWEN_GUI_TEST_OCR_REPORT", str(report))
    monkeypatch.setenv("DOCWEN_GUI_TEST_OCR_INPUT", str(source))
    monkeypatch.setenv("DOCWEN_GUI_TEST_OCR_OUTPUT_DIR", str(output_dir))

    batches: list[dict[str, object]] = []

    class _ConfigPort:
        def set_many(self, values: dict[str, object]) -> bool:
            batches.append(dict(values))
            return False

    class _ViewModel:
        controller = SimpleNamespace(config_port=_ConfigPort())

        def add_files(self, _paths: list[str]) -> None:
            raise AssertionError("OCR conversion must not start after config failure")

    close_calls: list[bool] = []
    window = SimpleNamespace(view_model=_ViewModel(), close=lambda: close_calls.append(True))
    app = SimpleNamespace(processEvents=lambda: None)
    monkeypatch.setattr(QTimer, "singleShot", lambda _delay, callback: callback())

    _schedule_test_ocr_report(cast(Any, app), cast(Any, window))

    assert batches == [
        {
            "output.directory.mode": "custom",
            "output.directory.custom_path": str(output_dir),
            "output.directory.create_date_subfolder": False,
        }
    ]
    assert close_calls == [True]
    payload = json.loads(report.read_text(encoding="utf-8"))
    assert payload["success"] is False
    assert payload["status"] == "failed"
    assert payload["error"] == "output_config_persist_failed"
