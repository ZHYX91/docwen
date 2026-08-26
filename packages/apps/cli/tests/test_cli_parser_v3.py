"""Parser and adapter contracts for the breaking CLI 0.9 command tree."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from packages.apps.cli.tests.capability_fixtures import bundled_available_runtime_projection

pytestmark = pytest.mark.contract


def _top_level_commands() -> set[str]:
    from docwen_cli.main import _build_parser

    parser = _build_parser()
    action = next(item for item in parser._actions if isinstance(item, argparse._SubParsersAction))
    return set(action.choices)


def test_top_level_command_tree_is_frozen() -> None:
    assert _top_level_commands() == {
        "info",
        "doctor",
        "inspect",
        "resources",
        "schema",
        "serve",
        "convert",
        "validate",
        "number",
        "merge",
        "split",
        "batch",
        "gui",
        "config",
    }


@pytest.mark.parametrize("old_command", ["run", "list", "templates", "numbering-schemes", "merge-pdfs"])
def test_old_top_level_commands_are_rejected(old_command: str) -> None:
    from docwen_cli.main import main

    assert main([old_command]) == 2


def test_batch_controls_are_not_global() -> None:
    from docwen_cli.main import main

    assert main(["info", "--jobs", "2"]) == 2
    assert main(["doctor", "--continue-on-error"]) == 2
    assert main(["inspect", "missing.docx", "--yes"]) == 2


def test_convert_projects_exact_output_policy(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main
    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.result import ConversionResult

    source = tmp_path / "source.md"
    source.write_text("# Title", encoding="utf-8")
    output = tmp_path / "chosen.docx"

    controller = MagicMock()
    controller.has_runtime = True
    controller.describe_runtime_capabilities.return_value = bundled_available_runtime_projection()
    controller.execute_single.return_value = ConversionResult(
        task_id="convert-v3",
        success=True,
        artifacts=[
            ArtifactManifest(
                artifact_id="primary",
                kind="primary",
                staging_path=str(output),
                suggested_name=output.name,
                is_primary=True,
            )
        ],
    )

    exit_code = main(
        ["convert", str(source), "--to", "docx", "--output", str(output), "--json"],
        controller=controller,
    )

    request = controller.execute_single.call_args.args[0]
    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 0
    assert request.action_name == ""
    assert request.target_format == "docx"
    assert request.output_policy.output_path == str(output)
    assert request.output_policy.output_dir is None
    assert request.output_policy.overwrite_mode == "error"
    assert payload["command"] == "convert"
    assert payload["protocol_version"] == 3


def test_convert_projects_markdown_parent_output_policy(tmp_path: Path) -> None:
    from docwen_cli.commands.execution_v3 import _prepare_args
    from docwen_cli.main import _build_parser

    output_parent = tmp_path / "published"
    args = _build_parser().parse_args(
        ["convert", str(tmp_path / "a.docx"), "--to", "md", "--output", str(output_parent)]
    )

    _prepare_args(args)

    assert args.output_path is None
    assert args.output_dir == str(output_parent)


def test_convert_overwrite_is_explicit(tmp_path: Path) -> None:
    from docwen_cli.commands.execution_v3 import _prepare_args
    from docwen_cli.main import _build_parser

    args = _build_parser().parse_args(
        ["convert", str(tmp_path / "a.md"), "--to", "docx", "--output", str(tmp_path / "a.docx"), "--overwrite"]
    )
    _prepare_args(args)

    assert args.output_path.endswith("a.docx")
    assert args.overwrite is True


@pytest.mark.parametrize(
    ("operation", "expected_clean", "expected_add"),
    [
        ("add", "keep", "gongwen_standard"),
        ("remove", "remove", "none"),
    ],
)
def test_number_markdown_projects_explicit_operation(
    tmp_path: Path,
    operation: str,
    expected_clean: str,
    expected_add: str,
) -> None:
    from docwen_cli.commands.execution_v3 import _prepare_args
    from docwen_cli.main import _build_parser

    argv = [
        "number",
        "markdown",
        str(tmp_path / "a.md"),
        "--operation",
        operation,
        "--in-place",
    ]
    if operation == "add":
        argv.extend(["--scheme", "gongwen_standard"])
    args = _build_parser().parse_args(argv)
    _prepare_args(args)

    assert args.clean_numbering == expected_clean
    assert args.add_numbering == expected_add
    assert args.output_path.endswith("a.md")
    assert args.overwrite is True


def test_number_markdown_add_uses_stable_default_scheme(tmp_path: Path) -> None:
    from docwen_cli.commands.execution_v3 import _prepare_args
    from docwen_cli.main import _build_parser

    args = _build_parser().parse_args(
        [
            "number",
            "markdown",
            str(tmp_path / "a.md"),
            "--operation",
            "add",
            "--in-place",
        ]
    )
    _prepare_args(args)

    assert args.clean_numbering == "keep"
    assert args.add_numbering == "hierarchical_standard"


def test_number_markdown_rejects_scheme_for_remove(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main

    source = tmp_path / "a.md"
    source.write_text("# Title\n", encoding="utf-8")
    exit_code = main(
        [
            "number",
            "markdown",
            str(source),
            "--operation",
            "remove",
            "--scheme",
            "gongwen_standard",
            "--in-place",
            "--json",
        ]
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 2
    assert payload["error"]["code"] == "invalid_input"


def test_number_markdown_requires_explicit_operation(tmp_path: Path) -> None:
    from docwen_cli.main import main

    assert main(["number", "markdown", str(tmp_path / "a.md"), "--in-place"]) == 2


def test_batch_owns_concurrency_controls() -> None:
    from docwen_cli.commands.execution_v3 import _prepare_args
    from docwen_cli.main import _build_parser

    args = _build_parser().parse_args(
        [
            "batch",
            "convert",
            "a.docx",
            "b.docx",
            "--to",
            "md",
            "--output-dir",
            "out",
            "--jobs",
            "4",
            "--continue-on-error",
        ]
    )
    _prepare_args(args)

    assert args.command_path == "batch convert"
    assert args.batch is True
    assert args.jobs == 4
    assert args.continue_on_error is True


@pytest.mark.parametrize("timeout", ["0", "1801", "not-a-number"])
def test_timeout_is_bounded(timeout: str) -> None:
    from docwen_cli.main import main

    assert main(["gui", "status", "--timeout", timeout]) == 2


def test_gui_control_fails_typed_until_transport_exists(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    exit_code = main(["gui", "status", "--json"])

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 6
    assert payload["command"] == "gui status"
    assert payload["error"]["category"] == "unavailable"
    assert payload["error"]["code"] == "capability_unavailable"


def test_schema_is_derived_from_new_parser(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    assert main(["schema", "merge", "pdf", "--json"]) == 0

    payload = json.loads(capsys.readouterr().out)
    assert payload["data"]["command"] == "merge pdf"
    names = {item["name"] for item in payload["data"]["arguments"]}
    assert {"files", "output", "overwrite", "timeout"}.issubset(names)
    assert "action" not in names


def test_optimization_argument_is_a_runtime_resource_id_not_a_static_choice(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import _build_parser, main

    args = _build_parser().parse_args(
        [
            "convert",
            str(tmp_path / "a.docx"),
            "--to",
            "md",
            "--output",
            str(tmp_path / "a.md"),
            "--optimization",
            "manifest-defined-resource",
        ]
    )
    assert args.optimization == "manifest-defined-resource"

    assert main(["schema", "convert", "--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    optimization = next(item for item in payload["data"]["arguments"] if item["name"] == "optimization")
    assert optimization["choices"] is None
