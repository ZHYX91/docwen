"""Tests for CLI main() security startup behaviour.

Covers: strict security startup, non-strict degraded mode, failure labels,
exit codes, and user-facing error paths.
"""

from __future__ import annotations

import json

import pytest
from packages.apps.cli.tests.capability_fixtures import bundled_available_runtime_projection

pytestmark = pytest.mark.unit


# ── Helpers ────────────────────────────────────────────────────────────


def _json_flag(argv: list[str]) -> list[str]:
    return [*argv, "--json"]


# ── Strict security blocks startup ────────────────────────────────────


class TestStrictSecurityBlocksStartup:
    """F-B2-003, F-B2-004: strict mode blocks CLI startup on failure."""

    def test_strict_security_env_returns_exit_code_5(self, monkeypatch, capsys) -> None:
        from docwen_cli.exit_codes import ExitCode
        from docwen_cli.main import main

        monkeypatch.setenv("DOCWEN_STRICT_SECURITY", "1")
        monkeypatch.setattr(
            "docwen_runtime.security._run_all_protections",
            lambda: (_ for _ in ()).throw(RuntimeError("test strict failure")),
        )

        rc = main(["run", "test.md", "--to", "md"])
        assert rc == int(ExitCode.SECURITY_CHECK_FAILED), f"got {rc}"
        captured = capsys.readouterr()
        assert "安全检查失败" in captured.err

    def test_strict_security_blocks_startup_json_mode(self, monkeypatch, capsys) -> None:
        """In strict mode, security failure blocks before JSON presenter init."""
        from docwen_cli.exit_codes import ExitCode
        from docwen_cli.main import main

        monkeypatch.setenv("DOCWEN_STRICT_SECURITY", "1")
        monkeypatch.setattr(
            "docwen_runtime.security._run_all_protections",
            lambda: (_ for _ in ()).throw(RuntimeError("json mode failure")),
        )

        rc = main(_json_flag(["run", "test.md", "--to", "md"]))
        assert rc == int(ExitCode.SECURITY_CHECK_FAILED)
        captured = capsys.readouterr()
        assert "安全检查失败" in captured.err


# ── Non-strict (degraded) mode allows startup ─────────────────────────


class TestNonStrictSecurityAllowsStartup:
    """Non-strict mode logs warning but does not block startup."""

    def test_non_strict_logs_warning_and_continues(self, monkeypatch, capsys, caplog) -> None:
        import logging

        from docwen_cli.main import main

        monkeypatch.setenv("DOCWEN_STRICT_SECURITY", "0")
        monkeypatch.setattr(
            "docwen_runtime.security._run_all_protections",
            lambda: (_ for _ in ()).throw(RuntimeError("degraded test failure")),
        )

        with caplog.at_level(logging.WARNING, logger="docwen_cli"):
            rc = main(["info"])
        # Lightweight info remains available without runtime composition.
        assert rc == 0, f"expected 0, got {rc}"
        assert "降级运行" in caplog.text or "degraded test failure" in caplog.text

    def test_non_strict_env_passes_normal_startup(self, monkeypatch, capsys) -> None:
        from docwen_cli.main import main

        monkeypatch.setenv("DOCWEN_STRICT_SECURITY", "0")
        # _run_all_protections passes normally (no debugger in CI)
        rc = main(["info"])
        assert rc == 0


class TestInspectRuntimeBootstrap:
    """Inspect must initialize Runtime before advertising supported actions."""

    def test_inspect_owns_controller_for_capability_discovery(self, monkeypatch, tmp_path, capsys) -> None:
        from unittest.mock import MagicMock

        from docwen_cli.main import main

        source = tmp_path / "proofread.md"
        source.write_text("# 标题\n\n这是一段需要校对的文本。\n", encoding="utf-8")
        controller = MagicMock()
        controller.describe_runtime_capabilities.return_value = bundled_available_runtime_projection()
        create_controller = MagicMock(return_value=controller)
        monkeypatch.setattr("docwen_cli.main._create_controller", create_controller)

        assert main(["inspect", str(source), "--lang", "zh_CN", "--json"]) == 0

        payload = json.loads(capsys.readouterr().out)
        assert payload["data"]["supported_actions_discovery"]["state"] == "available"
        assert "validate" in payload["data"]["supported_actions"]
        create_controller.assert_called_once()
        controller.stop.assert_called_once()


# ── Security startup is first — before lang parsing ───────────────────


class TestSecurityBeforeLang:
    """Security startup runs BEFORE --lang parsing to avoid i18n dependency."""

    def test_security_failure_before_lang_init(self, monkeypatch, capsys) -> None:
        from docwen_cli.exit_codes import ExitCode
        from docwen_cli.main import main

        monkeypatch.setenv("DOCWEN_STRICT_SECURITY", "1")
        monkeypatch.setattr(
            "docwen_runtime.security._run_all_protections",
            lambda: (_ for _ in ()).throw(RuntimeError("pre-lang failure")),
        )

        # --lang should not be parsed before security check
        rc = main(["--lang", "zh_CN", "run", "test.md", "--to", "md"])
        assert rc == int(ExitCode.SECURITY_CHECK_FAILED)


# ── Failure labels used in output ──────────────────────────────────────


class TestFailureLabelsInOutput:
    """F-B1-004: CLI error output uses centralized FAILURE_LABELS."""

    def test_bootstrap_error_uses_initialization_label(self, monkeypatch, capsys) -> None:
        from docwen_cli.main import main

        # Force controller creation to fail AFTER argparse succeeds.
        # Provide enough args for argparse to accept the command.
        monkeypatch.setattr(
            "docwen_cli.main._create_controller",
            lambda args, **kwargs: (_ for _ in ()).throw(RuntimeError("bootstrap forced failure")),
        )

        main(["convert", "test.md", "--to", "md", "--output", "test-out.md"])
        captured = capsys.readouterr()
        assert "初始化失败" in captured.err

    def test_keyboard_interrupt_uses_user_interrupt_label(self, monkeypatch, capsys) -> None:
        from docwen_cli.exit_codes import ExitCode
        from docwen_cli.main import main

        # Force a KeyboardInterrupt during execution
        monkeypatch.setattr(
            "docwen_cli.main._init_command_table",
            lambda: None,
        )
        # Monkeypatch the command table to route to a function that raises
        from docwen_cli.main import _COMMAND_TABLE

        _COMMAND_TABLE["doctor"] = (
            lambda args, controller: (_ for _ in ()).throw(KeyboardInterrupt()),
            False,
        )

        rc = main(["doctor"])
        captured = capsys.readouterr()
        assert "用户中断" in captured.err
        assert rc == int(ExitCode.CANCELLED)

        # Clean up
        _COMMAND_TABLE.clear()

    def test_keyboard_interrupt_json_mode_uses_label(self, monkeypatch, capsys) -> None:
        from docwen_cli.main import _COMMAND_TABLE, main

        _COMMAND_TABLE["doctor"] = (
            lambda args, controller: (_ for _ in ()).throw(KeyboardInterrupt()),
            False,
        )

        main(_json_flag(["doctor"]))
        captured = capsys.readouterr()
        assert "用户中断" in captured.out

        _COMMAND_TABLE.clear()


# ── Default environment (no env var set) is strict ────────────────────


class TestDefaultStrict:
    """When DOCWEN_STRICT_SECURITY is unset, behaviour is strict (1)."""

    def test_default_is_strict(self, monkeypatch, capsys) -> None:
        from docwen_cli.exit_codes import ExitCode
        from docwen_cli.main import main

        monkeypatch.delenv("DOCWEN_STRICT_SECURITY", raising=False)
        monkeypatch.setattr(
            "docwen_runtime.security._run_all_protections",
            lambda: (_ for _ in ()).throw(RuntimeError("default strict failure")),
        )

        rc = main(["run", "test.md", "--to", "md"])
        assert rc == int(ExitCode.SECURITY_CHECK_FAILED)

    def test_no_security_issues_normal_startup(self, monkeypatch) -> None:
        from docwen_cli.main import main

        monkeypatch.delenv("DOCWEN_STRICT_SECURITY", raising=False)
        # No debugger attached → protections pass
        rc = main(["info"])
        assert rc == 0


# ── SecurityCheckFailedError caught generically ────────────────────────


class TestSecurityCheckErrorIntegration:
    def test_security_exception_has_exit_code_5(self, monkeypatch, capsys) -> None:
        from docwen_cli.exit_codes import ExitCode
        from docwen_cli.main import main

        monkeypatch.setenv("DOCWEN_STRICT_SECURITY", "1")
        monkeypatch.setattr(
            "docwen_runtime.security._run_all_protections",
            lambda: (_ for _ in ()).throw(RuntimeError("check")),
        )

        rc = main(["run", "test.md", "--to", "md"])
        assert rc == 5
        assert rc == int(ExitCode.SECURITY_CHECK_FAILED)
