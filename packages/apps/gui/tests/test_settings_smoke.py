from __future__ import annotations

import json
from types import SimpleNamespace

import pytest
from PySide6.QtWidgets import QWidget

pytestmark = pytest.mark.gui


def test_settings_smoke_constructs_every_page_without_placeholders(qapp, tmp_path, monkeypatch) -> None:
    from docwen_gui.settings_smoke import _schedule_test_settings_report
    from docwen_gui.widgets.settings.dialog import TAB_KEYS

    report_path = tmp_path / "settings-smoke.json"
    monkeypatch.setenv("DOCWEN_GUI_TEST_SETTINGS_REPORT", str(report_path))
    window = QWidget()
    window.view_model = SimpleNamespace(controller=None)  # type: ignore[attr-defined]

    _schedule_test_settings_report(qapp, window)  # type: ignore[arg-type]
    qapp.processEvents()

    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["success"] is True
    assert report["expectedTabs"] == TAB_KEYS
    assert report["loadedTabs"] == TAB_KEYS
    assert report["failedTabs"] == []
    assert report["missingTabs"] == []
    assert report["unexpectedTabs"] == []
    assert report["error"] is None
    window.close()


def test_settings_smoke_reports_a_page_factory_failure(qapp, tmp_path, monkeypatch) -> None:
    from docwen_gui.settings_smoke import _schedule_test_settings_report
    from docwen_gui.widgets.settings import dialog as dialog_module

    original = dialog_module._TAB_SPECS["document"]  # pyright: ignore[reportPrivateUsage]

    def fail_document(_view_model):
        raise ImportError("document settings dependency unavailable")

    monkeypatch.setitem(
        dialog_module._TAB_SPECS,  # pyright: ignore[reportPrivateUsage]
        "document",
        original._replace(factory=fail_document),
    )
    report_path = tmp_path / "settings-smoke-failed.json"
    monkeypatch.setenv("DOCWEN_GUI_TEST_SETTINGS_REPORT", str(report_path))
    window = QWidget()
    window.view_model = SimpleNamespace(controller=None)  # type: ignore[attr-defined]

    _schedule_test_settings_report(qapp, window)  # type: ignore[arg-type]
    qapp.processEvents()

    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["success"] is False
    assert report["failedTabs"] == ["document"]
    assert "document" not in report["loadedTabs"]
    assert report["pageObjectNames"]["document"] == "settingsTabLoadErrorPage"
    window.close()
