"""Protocol 3 parser and command error behavior."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract


def test_unknown_command_is_usage_error(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    assert main(["unknown-command"]) == 2
    captured = capsys.readouterr()
    assert captured.out == ""
    assert "invalid choice" in captured.err


def test_unknown_command_json_is_one_document(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    assert main(["unknown-command", "--json"]) == 2
    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert captured.err == ""
    assert payload["protocol_version"] == 3
    assert payload["error"]["code"] == "invalid_arguments"


@pytest.mark.parametrize("argv", [["--help"], ["convert", "--help"], ["merge", "pdf", "--help"]])
def test_help_is_success(argv: list[str]) -> None:
    from docwen_cli.main import main

    assert main(argv) == 0


def test_old_command_help_is_not_a_compatibility_alias() -> None:
    from docwen_cli.main import main

    assert main(["run", "--help"]) == 2


def test_missing_required_output_is_usage_error() -> None:
    from docwen_cli.main import main

    assert main(["convert", "a.docx", "--to", "md"]) == 2


def test_quiet_and_verbose_are_mutually_exclusive() -> None:
    from docwen_cli.main import main

    assert main(["info", "--quiet", "--verbose"]) == 2


def test_inspect_missing_file_is_typed(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    missing = tmp_path / "missing.docx"
    assert main(["inspect", str(missing), "--json"]) == 2
    payload = json.loads(capsys.readouterr().out)
    assert payload["error"]["category"] == "invalid_input"
    assert payload["error"]["code"] == "file_not_found"


def test_resources_missing_id_is_not_found(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    class EmptyOptimizationController:
        def describe_runtime_capabilities(self) -> dict[str, object]:
            runtime = {"state": "available", "platform": "windows"}
            return {
                "resource": "formats",
                "contract": {"id": "docwen.runtime-capabilities", "version": 1},
                "runtime": runtime,
                "security": {"dependency_egress_guard": {}},
                "gates": [],
                "sources": [],
                "counts": {
                    "sources": 0,
                    "routes": 0,
                    "available_routes": 0,
                    "unavailable_routes": 0,
                    "actions": 0,
                },
                "optimizations": {
                    "resource": "optimizations",
                    "contract": {"id": "docwen.optimizations", "version": 1},
                    "runtime": runtime,
                    "resources": [],
                    "counts": {
                        "resources": 0,
                        "available_resources": 0,
                        "unavailable_resources": 0,
                        "bindings": 0,
                        "available_bindings": 0,
                        "unavailable_bindings": 0,
                    },
                },
            }

    assert (
        main(
            ["resources", "show", "optimizations", "missing", "--json"],
            controller=EmptyOptimizationController(),
        )
        == 3
    )
    payload = json.loads(capsys.readouterr().out)
    assert payload["error"]["code"] == "resource_not_found"
