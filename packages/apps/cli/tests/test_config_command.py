"""Protocol 3 config reset is explicit, scoped, and non-interactive."""

from __future__ import annotations

import json
from unittest.mock import MagicMock

import pytest

pytestmark = pytest.mark.contract


def test_config_reset_requires_group_and_yes() -> None:
    from docwen_cli.main import main

    assert main(["config", "reset", "general"]) == 2
    assert main(["config", "reset", "--yes"]) == 2


def test_config_reset_group_calls_one_group(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    controller = MagicMock()
    controller.config_port.reset_group.return_value = True

    exit_code = main(["config", "reset", "general", "--yes", "--json"], controller=controller)

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 0
    controller.config_port.reset_group.assert_called_once_with("general")
    controller.config_port.reset_all.assert_not_called()
    assert payload["command"] == "config reset"
    assert payload["protocol_version"] == 3


def test_config_reset_all_calls_reset_all() -> None:
    from docwen_cli.main import main

    controller = MagicMock()
    controller.config_port.reset_all.return_value = True

    assert main(["config", "reset", "all", "--yes"], controller=controller) == 0

    controller.config_port.reset_all.assert_called_once_with()
    controller.config_port.reset_group.assert_not_called()


def test_config_reset_never_prompts(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_cli.main import main

    controller = MagicMock()
    controller.config_port.reset_group.return_value = True
    monkeypatch.setattr("builtins.input", lambda *_: (_ for _ in ()).throw(AssertionError("must not prompt")))

    assert main(["config", "reset", "output", "--yes"], controller=controller) == 0


def test_config_reset_failure_is_typed(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    controller = MagicMock()
    controller.config_port.reset_group.return_value = False

    exit_code = main(["config", "reset", "logging", "--yes", "--json"], controller=controller)

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 1
    assert payload["error"]["category"] == "internal"
    assert payload["error"]["code"] == "reset_failed"


@pytest.mark.parametrize(
    "group",
    [
        "general",
        "output",
        "logging",
        "export",
        "other",
        "formatting",
        "document",
        "text",
        "spreadsheet",
        "image",
        "layout",
        "link",
        "proofread",
    ],
)
def test_every_public_group_is_accepted(group: str) -> None:
    from docwen_cli.main import _build_parser

    args = _build_parser().parse_args(["config", "reset", group, "--yes"])
    assert args.group == group
