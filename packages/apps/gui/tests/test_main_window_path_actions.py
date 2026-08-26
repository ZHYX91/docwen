"""Main-window routing contracts for open versus precise reveal actions."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any, cast

import pytest

from docwen_gui import path_actions
from docwen_gui.main_window import MainWindow
from docwen_gui.path_actions import PathActionResult

pytestmark = pytest.mark.unit


class _InfoArea:
    def __init__(self) -> None:
        self.messages: list[tuple[str, str]] = []

    def add_message(self, message: str, tone: str) -> None:
        self.messages.append((message, tone))


def test_main_window_open_parent_routes_file_to_precise_reveal(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    target = tmp_path / "report.docx"
    target.write_text("done", encoding="utf-8")
    calls: list[str] = []
    monkeypatch.setattr(
        path_actions,
        "reveal_path",
        lambda value: calls.append(str(value)) or PathActionResult(success=True, precise=True),
    )
    fake = SimpleNamespace(_info_area_vm=_InfoArea())

    opened = MainWindow._open_path(cast(Any, fake), str(target), open_parent=True)

    assert opened is True
    assert calls == [str(target)]
    assert fake._info_area_vm.messages == []


def test_main_window_plain_open_uses_desktop_open(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    target = tmp_path / "report.docx"
    target.write_text("done", encoding="utf-8")
    calls: list[str] = []
    monkeypatch.setattr(
        path_actions,
        "open_path",
        lambda value: calls.append(str(value)) or PathActionResult(success=True),
    )
    fake = SimpleNamespace(_info_area_vm=_InfoArea())

    opened = MainWindow._open_path(cast(Any, fake), str(target), open_parent=False)

    assert opened is True
    assert calls == [str(target)]


def test_navigation_request_uses_reveal_semantics() -> None:
    calls: list[tuple[str, bool]] = []
    fake = SimpleNamespace(
        _open_path=lambda target, *, open_parent=False: calls.append((target, open_parent)) or True,
    )

    MainWindow._handle_navigation_request(cast(Any, fake), "D:/output/report.docx")

    assert calls == [("D:/output/report.docx", True)]


def test_main_window_reveal_missing_path_surfaces_warning(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    missing = tmp_path / "missing.docx"
    monkeypatch.setattr(
        path_actions,
        "reveal_path",
        lambda value: PathActionResult(success=False, error=str(value), error_code="missing_path"),
    )
    fake = SimpleNamespace(_info_area_vm=_InfoArea())

    opened = MainWindow._open_path(cast(Any, fake), str(missing), open_parent=True)

    assert opened is False
    assert len(fake._info_area_vm.messages) == 1
    message, tone = fake._info_area_vm.messages[0]
    assert str(missing) in message
    assert tone == "warning"
