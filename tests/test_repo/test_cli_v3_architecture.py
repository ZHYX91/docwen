"""Architecture guards for the protocol 3 CLI boundary."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
CLI_ROOT = PROJECT_ROOT / "packages" / "apps" / "cli" / "src" / "docwen_cli"


def _source(relative_path: str) -> str:
    return (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")


def test_convert_module_is_execution_orchestration_not_a_cli_monolith() -> None:
    source = _source("packages/apps/cli/src/docwen_cli/commands/convert.py")
    request_support = _source("packages/apps/cli/src/docwen_cli/commands/execution_request.py")

    assert len(source.splitlines()) <= 1000
    assert "class ExecutionDeadline" not in source
    assert "def build_execution_options" not in source
    assert "def validate_execution_options" not in source
    assert "from docwen_cli.execution_deadline import" in source
    assert "from docwen_cli.commands.execution_options import" in source
    assert "from docwen_cli.commands import execution_request" in source
    assert "def file_ref_for_runtime" not in source
    assert "def output_policy" not in source
    assert "def file_ref_for_runtime" in request_support
    assert "def output_policy" in request_support
    for removed_alias in (
        "_configured_ocr_language",
        "_file_ref_for_runtime",
        "_output_policy",
        "_public_command",
        "_redacted_options",
        "_resolve_request_route",
        "_route_scoped_options",
        "_to_markdown_locale_options",
    ):
        assert removed_alias not in source


def test_protocol_3_adapter_has_no_removed_cli_fields_or_destination_fallback() -> None:
    adapter = _source("packages/apps/cli/src/docwen_cli/commands/execution_v3.py")
    options = _source("packages/apps/cli/src/docwen_cli/commands/execution_options.py")
    convert = _source("packages/apps/cli/src/docwen_cli/commands/convert.py")

    for removed in ('"base_table"', '"compress"', '"size_limit"', '"quality_mode"'):
        assert removed not in adapter
        assert removed not in options
    assert "Transitional support" not in convert
    assert "legacy_output" not in convert


def test_bundle_depends_on_gui_control_port_not_command_implementation() -> None:
    command = _source("packages/apps/cli/src/docwen_cli/commands/gui_control.py")
    adapter = _source("packages/bundle/src/docwen_bundle/gui_control_adapter.py")

    assert "class GuiControlError" not in command
    assert "from docwen_cli.gui_control_port import GuiControlError, GuiControlPort" in adapter
    assert "from docwen_cli.commands.gui_control import" not in adapter


def test_control_transport_does_not_claim_fixed_protocol_key_is_authentication() -> None:
    package = _source("packages/runtime/src/docwen_runtime/control/__init__.py")
    transport = _source("packages/runtime/src/docwen_runtime/control/transport.py")

    assert "Authenticated local" not in package
    assert "_AUTHKEY" not in transport
    assert "_PROTOCOL_KEY" in transport
    assert "is not a user secret or an authentication boundary" in transport
