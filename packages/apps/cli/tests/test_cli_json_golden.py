"""Small stable golden surfaces for protocol 3 discovery and errors."""

from __future__ import annotations

import json

import pytest

pytestmark = pytest.mark.contract


def test_info_golden(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    assert main(["info", "--json"]) == 0
    payload = json.loads(capsys.readouterr().out)

    assert payload["protocol_version"] == 3
    assert payload["product_version"] == "0.9.0"
    assert payload["command"] == "info"
    assert payload["success"] is True
    assert payload["data"]["protocol"] == {"major": 3, "envelope": "docwen.cli.v3"}


def test_invalid_arguments_golden(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    assert main(["convert", "source.docx", "--to", "md", "--json"]) == 2
    payload = json.loads(capsys.readouterr().out)

    assert payload["protocol_version"] == 3
    assert payload["command"] == "convert"
    assert payload["success"] is False
    assert payload["error"]["category"] == "invalid_input"
    assert payload["error"]["code"] == "invalid_arguments"
    assert payload["error"]["hint"].startswith("usage: docwen convert")


def test_gui_unavailable_golden(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    assert main(["gui", "status", "--json"]) == 6
    payload = json.loads(capsys.readouterr().out)

    assert payload["command"] == "gui status"
    assert payload["error"] == {
        "category": "unavailable",
        "code": "capability_unavailable",
        "message": "GUI control is unavailable in this CLI assembly.",
        "details": {},
        "hint": None,
    }
