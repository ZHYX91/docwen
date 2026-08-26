"""CLI main entry point.

Usage:
    docwen <command> [args]
    docwen <command> [args]  (via console_scripts)

Entry flow:
1. Reconfigure stdio for UTF-8 (Windows codepage 65001).
2. Pre-parse ``--lang`` to set the locale.
3. Build the argument parser with all registered commands.
4. Resolve the command and dispatch to its executor.
5. Convert exceptions to appropriate exit codes.
"""

from __future__ import annotations

import argparse
import contextlib
import os
import sys
from typing import Any

from docwen_cli.exit_codes import ExitCode, exit_code_from_error_code
from docwen_cli.i18n import cli_t, init_cli_locale
from docwen_cli.parser import (
    CliArgumentParser,
    CliUsageError,
    get_available_locale_codes,
    get_common_parser,
    get_root_common_defaults_parser,
)
from docwen_cli.path_policy import first_namespace_path_issue
from docwen_core.errors import DocWenError
from docwen_runtime.errors import FAILURE_LABELS, FailureCategory, SecurityCheckFailedError
from docwen_runtime.security import NetworkAccessBlockedError, resolve_strict_security, run_security_protections

# ── Windows UTF-8 console encoding ──────────────────────────────────

_CODEPAGE_SET = False


def ensure_console_utf8() -> None:
    """Reconfigure Windows console for UTF-8 output.

    Sets console codepage to 65001 (UTF-8) on Windows, then reconfigures
    ``sys.stdout``/``sys.stderr`` text wrappers to use UTF-8 encoding.
    """
    global _CODEPAGE_SET

    if sys.platform == "win32" and not _CODEPAGE_SET:
        _CODEPAGE_SET = True
        try:
            import ctypes

            kernel32 = ctypes.windll.kernel32
            kernel32.SetConsoleOutputCP(65001)
            kernel32.SetConsoleCP(65001)
        except Exception:
            pass  # non-fatal; UTF-8 wrapper below is the fallback

    try:
        import io

        stdout_reconf = getattr(sys.stdout, "reconfigure", None)
        stderr_reconf = getattr(sys.stderr, "reconfigure", None)
        if callable(stdout_reconf):
            with contextlib.suppress(OSError, ValueError):
                stdout_reconf(encoding="utf-8", errors="replace")
        if callable(stderr_reconf):
            with contextlib.suppress(OSError, ValueError):
                stderr_reconf(encoding="utf-8", errors="replace")

        if isinstance(sys.stdout, io.TextIOWrapper) and getattr(sys.stdout, "encoding", None) != "utf-8":
            with contextlib.suppress(Exception):
                sys.stdout = io.TextIOWrapper(
                    sys.stdout.buffer,
                    encoding="utf-8",
                    errors="replace",
                    line_buffering=True,
                )
        if isinstance(sys.stderr, io.TextIOWrapper) and getattr(sys.stderr, "encoding", None) != "utf-8":
            with contextlib.suppress(Exception):
                sys.stderr = io.TextIOWrapper(
                    sys.stderr.buffer,
                    encoding="utf-8",
                    errors="replace",
                    line_buffering=True,
                )
    except Exception:
        if os.environ.get("DOCWEN_FAIL_FAST") == "1":
            raise


# ── Pre-parse --lang ────────────────────────────────────────────────


def pre_parse_lang(argv: list[str] | None = None) -> str | None:
    """Pre-parse ``--lang`` so help text renders in the correct language.

    Must be called BEFORE the main argument parser is created.
    """
    args = argv if argv is not None else sys.argv[1:]
    available = frozenset(get_available_locale_codes())
    resolved: str | None = None
    for i, arg in enumerate(args):
        if arg == "--":
            break
        if arg == "--lang" and i + 1 < len(args):
            lang = args[i + 1]
            if lang in available:
                resolved = lang
        elif arg.startswith("--lang="):
            lang = arg.split("=", 1)[1]
            if lang in available:
                resolved = lang
    return resolved


def _argv_requests_json(argv: list[str] | None) -> bool:
    """Return whether machine output was explicitly requested."""

    args = argv if argv is not None else sys.argv[1:]
    return "--json" in args


def _command_hint(argv: list[str] | None) -> str:
    """Return a bounded command hint for a parser-error envelope."""

    args = argv if argv is not None else sys.argv[1:]
    values: list[str] = []
    skip_next = False
    for item in args:
        if skip_next:
            skip_next = False
            continue
        if item == "--lang":
            skip_next = True
            continue
        if item.startswith("-"):
            continue
        values.append(item)
        if len(values) == 2:
            break
    if not values:
        return ""
    containers = {"resources", "number", "merge", "split", "batch", "gui", "config"}
    if values[0] in containers and len(values) > 1:
        return " ".join(values[:2])
    return values[0]


def _public_command(args: argparse.Namespace, fallback: str) -> str:
    """Return the parsed leaf command used by every protocol envelope."""

    command_path = getattr(args, "command_path", None)
    if isinstance(command_path, str) and command_path:
        return command_path
    child = getattr(args, f"{fallback}_command", None)
    if isinstance(child, str) and child:
        return f"{fallback} {child}"
    return fallback


# ── Argument parser ─────────────────────────────────────────────────


def _build_parser() -> argparse.ArgumentParser:
    """Create the full CLI argument parser with all commands registered."""
    parser = CliArgumentParser(
        prog="docwen",
        parents=[get_common_parser(), get_root_common_defaults_parser()],
        description=cli_t("cli.description"),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            f"{cli_t('cli.example')}:\n"
            "  %(prog)s info --json\n"
            "  %(prog)s convert report.docx --to md --output published-documents\n"
            "  %(prog)s validate report.docx --json\n"
            "  %(prog)s merge pdf a.pdf b.pdf --output merged.pdf\n"
            "  %(prog)s gui status --json"
        ),
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # Import command registrars lazily so --version remains a light path.
    from docwen_cli.commands.config import register_config_parser
    from docwen_cli.commands.doctor import register_doctor_parser
    from docwen_cli.commands.execution_parsers import register_execution_parsers
    from docwen_cli.commands.gui_control import register_gui_control_parser
    from docwen_cli.commands.info import register_info_parser
    from docwen_cli.commands.inspect import register_inspect_parser
    from docwen_cli.commands.resources import register_resources_parser
    from docwen_cli.commands.schema_v3 import register_schema_v3_parser
    from docwen_cli.commands.serve import register_serve_parser

    register_info_parser(subparsers)
    register_inspect_parser(subparsers)
    register_doctor_parser(subparsers)
    register_resources_parser(subparsers)
    register_schema_v3_parser(subparsers)
    register_serve_parser(subparsers)
    register_execution_parsers(subparsers)
    register_gui_control_parser(subparsers)
    register_config_parser(subparsers)

    return parser


# ── Command dispatch ────────────────────────────────────────────────


# Map command name → (executor_fn, needs_controller)
# needs_controller=True means the command requires an ApplicationController
# with a working RuntimePort.
_COMMAND_TABLE: dict[str, tuple[Any, bool]] = {}


def _init_command_table() -> None:
    """Populate the command dispatch table (lazy)."""
    if _COMMAND_TABLE:
        return

    from docwen_cli.commands.config import execute_config
    from docwen_cli.commands.doctor import execute_doctor
    from docwen_cli.commands.execution_v3 import execute_execution
    from docwen_cli.commands.gui_control import execute_gui_control
    from docwen_cli.commands.info import execute_info
    from docwen_cli.commands.inspect import execute_inspect
    from docwen_cli.commands.resources import execute_resources
    from docwen_cli.commands.schema_v3 import execute_schema_v3
    from docwen_cli.commands.serve import execute_serve

    # Lightweight discovery — never initializes runtime or GUI.
    _COMMAND_TABLE["info"] = (execute_info, False)

    # Read-only discovery commands. Inspect and resources project the loaded
    # runtime composition, so both require an initialized controller.
    _COMMAND_TABLE["inspect"] = (execute_inspect, True)
    _COMMAND_TABLE["doctor"] = (execute_doctor, True)
    _COMMAND_TABLE["resources"] = (execute_resources, True)
    _COMMAND_TABLE["schema"] = (execute_schema_v3, False)
    _COMMAND_TABLE["serve"] = (execute_serve, False)

    # Domain execution commands share the application workflow, not parser aliases.
    for name in ("convert", "validate", "number", "merge", "split", "batch"):
        _COMMAND_TABLE[name] = (execute_execution, True)

    _COMMAND_TABLE["gui"] = (execute_gui_control, False)
    _COMMAND_TABLE["config"] = (execute_config, True)


# ── Controller bootstrap ────────────────────────────────────────────


def _build_args_config(args: argparse.Namespace) -> Any:
    """Build a minimal ``ConfigPort`` from parsed CLI arguments.

    Wraps CLI args in a config-like object so the controller always
    has access to basic config values (lang, json mode, etc.).
    Bundle layer overrides this in production with full typed config.
    """

    class _CliArgConfig:
        def get(self, key: str, default: Any = None) -> Any:
            # Map dotted keys to CLI args
            mapping: dict[str, str] = {
                "general.lang": "lang",
                "cli.json": "json",
                "cli.quiet": "quiet",
                "cli.verbose": "verbose",
                "cli.timing": "timing",
                "cli.batch": "batch",
                "cli.jobs": "jobs",
                "cli.continue_on_error": "continue_on_error",
            }
            attr = mapping.get(key)
            if attr and hasattr(args, attr):
                return getattr(args, attr)
            return default

        def snapshot(self) -> dict[str, Any]:
            return {}

        def set(self, key: str, value: Any) -> bool:
            return False

        def reset_file(self, rel_path: str) -> bool:
            return False

        def reset_group(self, group: str) -> bool:
            return False

        def reset_all(self) -> bool:
            return False

        def reload(self) -> None:
            return None

    return _CliArgConfig()


def _create_controller(
    args: argparse.Namespace,
    runtime_port: Any = None,
    config_port: Any = None,
    runtime_port_factory: Any = None,
    config_port_factory: Any = None,
) -> Any:
    """Create and start an ApplicationController.

    In production, *runtime_port* and *config_port* are injected by the
    bundle layer.
    """
    from docwen_application.controller import ApplicationController

    if config_port is None:
        config_port = config_port_factory() if config_port_factory is not None else _build_args_config(args)
    if runtime_port is None and runtime_port_factory is not None:
        runtime_port = runtime_port_factory()

    controller = ApplicationController(
        runtime_port=runtime_port,
        config_port=config_port,
    )
    controller.start()
    return controller


# ── Main entry ──────────────────────────────────────────────────────


def main(
    argv: list[str] | None = None,
    *,
    controller: Any | None = None,
    runtime_port: Any | None = None,
    config_port: Any | None = None,
    runtime_port_factory: Any | None = None,
    config_port_factory: Any | None = None,
    gui_control_port_factory: Any | None = None,
    machine_server_factory: Any | None = None,
) -> int:
    """CLI main entry point.

    Args:
        argv: Command-line arguments (defaults to ``sys.argv[1:]``).

    Returns:
        Exit code: 0 on success, non-zero on error.
    """
    # 0. Windows UTF-8 console encoding (matches old early_reconfigure_stdio)
    ensure_console_utf8()

    # 0.5. Security startup protections (F-B2-003, F-B2-004)
    import logging

    strict_security = resolve_strict_security()
    try:
        degraded_msg = run_security_protections(
            logger=logging.getLogger("docwen_cli"),
            strict_security=strict_security,
        )
        if degraded_msg is not None:
            # Non-strict mode: log diagnostic but do not block startup.
            logging.getLogger("docwen_cli").warning(degraded_msg)
    except SecurityCheckFailedError:
        # Strict mode: block startup with correct exit code and label.
        print(f"错误: {FAILURE_LABELS[FailureCategory.SECURITY_CHECK]}", file=sys.stderr)
        return int(ExitCode.SECURITY_CHECK_FAILED)

    # 1. Pre-parse --lang
    lang = pre_parse_lang(argv)
    init_cli_locale(lang)

    # Build parser and parse
    parser = _build_parser()
    try:
        args = parser.parse_args(argv)
    except CliUsageError as exc:
        if _argv_requests_json(argv):
            from docwen_cli.presenters.json_presenter import JsonPresenter

            JsonPresenter().present_error(
                _command_hint(argv),
                str(exc),
                error_code="invalid_arguments",
                hint=exc.usage.strip(),
            )
        else:
            print(exc.usage, end="", file=sys.stderr)
            print(f"docwen: error: {exc}", file=sys.stderr)
        return int(ExitCode.INVALID_INPUT)
    except SystemExit as e:
        # argparse exits on --help (code 0) or invalid args (code 2).
        code = e.code
        if isinstance(code, int):
            if code == 0:
                return int(ExitCode.OK)
            return int(ExitCode.INVALID_INPUT)
        return int(ExitCode.INVALID_INPUT)

    cmd = getattr(args, "command", "")
    if not cmd:
        parser.print_help()
        return int(ExitCode.INVALID_INPUT)

    public_cmd = _public_command(args, cmd)
    path_issue = first_namespace_path_issue(args)
    if path_issue is not None:
        if getattr(args, "json", False):
            from docwen_cli.presenters.json_presenter import JsonPresenter

            JsonPresenter().present_error(
                public_cmd,
                path_issue.message,
                error_code="invalid_path",
                details={"path": path_issue.path},
                hint="Shorten the path before retrying.",
            )
        else:
            print(f"Error: {path_issue.message}", file=sys.stderr)
        return int(ExitCode.INVALID_INPUT)

    _init_command_table()

    entry = _COMMAND_TABLE.get(cmd)
    if entry is None:
        return int(ExitCode.INTERNAL_ERROR)

    executor, needs_controller = entry

    if needs_controller and cmd in {"convert", "validate", "number", "merge", "split", "batch"}:
        from docwen_cli.commands.execution_v3 import preflight_execution

        preflight_exit = preflight_execution(args)
        if preflight_exit is not None:
            return preflight_exit
        public_cmd = _public_command(args, cmd)

    # Validate --jobs
    jobs = getattr(args, "jobs", 1)
    if jobs is None:
        jobs = 1
    if int(jobs) < 1:
        if getattr(args, "json", False):
            from docwen_cli.presenters.json_presenter import JsonPresenter

            p = JsonPresenter()
            p.present_error(public_cmd, "--jobs 必须 >= 1", error_code="invalid_input")
        else:
            print("错误: --jobs 必须 >= 1", file=sys.stderr)
        return int(ExitCode.INVALID_INPUT)

    owned_controller = False
    if needs_controller:
        try:
            if controller is None:
                controller = _create_controller(
                    args,
                    runtime_port=runtime_port,
                    config_port=config_port,
                    runtime_port_factory=runtime_port_factory,
                    config_port_factory=config_port_factory,
                )
                owned_controller = True
        except Exception as exc:
            return _handle_bootstrap_error(public_cmd, args, exc)

    command_service = controller
    if cmd == "serve":
        if machine_server_factory is None:
            return _handle_bootstrap_error(public_cmd, args, RuntimeError("Machine Protocol server is unavailable"))
        try:
            command_service = machine_server_factory()
        except Exception as exc:
            return _handle_bootstrap_error(public_cmd, args, exc)
    if cmd == "gui" and gui_control_port_factory is not None:
        try:
            command_service = gui_control_port_factory()
        except Exception as exc:
            return _handle_bootstrap_error(public_cmd, args, exc)

    # Execute
    try:
        return executor(args, command_service)
    except KeyboardInterrupt:
        if getattr(args, "json", False):
            from docwen_cli.presenters.json_presenter import JsonPresenter

            p = JsonPresenter()
            p.present_error(
                public_cmd,
                FAILURE_LABELS[FailureCategory.USER_INTERRUPT],
                error_code="operation_cancelled",
                details={"interrupted": True},
            )
        else:
            print(f"错误: {FAILURE_LABELS[FailureCategory.USER_INTERRUPT]}", file=sys.stderr)
        return int(ExitCode.CANCELLED)

    except NetworkAccessBlockedError as exc:
        return _handle_network_access_blocked(public_cmd, args, exc)

    except DocWenError as exc:
        return _handle_docwen_error(public_cmd, args, exc)

    except ValueError as exc:
        input_file = ""
        if hasattr(args, "files") and args.files:
            input_file = args.files[0]
        elif hasattr(args, "file") and args.file:
            input_file = str(args.file)
        return _handle_value_error(public_cmd, args, exc, input_file)

    except Exception as exc:
        return _handle_unknown_error(public_cmd, args, exc)

    finally:
        if owned_controller and controller is not None:
            with contextlib.suppress(Exception):
                controller.stop()


# ── Error handlers ──────────────────────────────────────────────────


def _handle_bootstrap_error(cmd: str, args: argparse.Namespace, exc: Exception) -> int:
    """Handle errors during controller bootstrap."""

    if isinstance(exc, NetworkAccessBlockedError):
        return _handle_network_access_blocked(cmd, args, exc)
    msg = f"{FAILURE_LABELS[FailureCategory.INITIALIZATION]}: {exc}"
    if getattr(args, "json", False):
        from docwen_cli.presenters.json_presenter import JsonPresenter

        p = JsonPresenter()
        p.present_error(cmd, msg, error_code="dependency_missing")
    else:
        print(f"错误: {msg}", file=sys.stderr)
    return int(ExitCode.DEPENDENCY_MISSING)


def _handle_network_access_blocked(
    cmd: str,
    args: argparse.Namespace,
    exc: NetworkAccessBlockedError,
) -> int:
    """Preserve the stable security error instead of folding it into internal."""

    if getattr(args, "json", False):
        from docwen_cli.presenters.json_presenter import JsonPresenter

        JsonPresenter().present_error(cmd, str(exc), error_code=exc.code)
    else:
        print(f"错误: {exc}", file=sys.stderr)
    return int(exit_code_from_error_code(exc.code))


def _handle_docwen_error(cmd: str, args: argparse.Namespace, exc: DocWenError) -> int:
    """Handle DocWenError exceptions."""
    code = getattr(exc, "code", "unknown_error")
    msg = str(exc)

    if getattr(args, "json", False):
        from docwen_cli.presenters.json_presenter import JsonPresenter

        p = JsonPresenter()
        p.present_error(cmd, msg, error_code=code)
    else:
        print(f"错误: {msg}", file=sys.stderr)

    return int(exit_code_from_error_code(code))


def _handle_value_error(
    cmd: str,
    args: argparse.Namespace,
    exc: ValueError,
    input_file: str = "",
) -> int:
    """Handle ValueError exceptions (typically invalid CLI input)."""
    prefix = cli_t("cli.messages.error_prefix")
    msg = str(exc)
    full_msg = msg if msg.startswith(prefix) else f"{prefix}: {msg}"

    if getattr(args, "json", False):
        from docwen_cli.presenters.json_presenter import JsonPresenter

        p = JsonPresenter()
        p.present_error(cmd, full_msg, error_code="invalid_input")
    else:
        print(f"错误: {full_msg}", file=sys.stderr)

    return int(ExitCode.INVALID_INPUT)


def _handle_unknown_error(
    cmd: str,
    args: argparse.Namespace,
    exc: Exception,
) -> int:
    """Handle unexpected exceptions."""
    if getattr(args, "json", False):
        from docwen_cli.presenters.json_presenter import JsonPresenter

        p = JsonPresenter()
        p.present_error(
            cmd,
            "An internal DocWen error occurred.",
            error_code="unknown_error",
            details={"exception_type": type(exc).__name__},
        )
    else:
        print("错误: DocWen 内部错误。请检查本地日志。", file=sys.stderr)

    return int(ExitCode.INTERNAL_ERROR)


# ── Direct invocation ───────────────────────────────────────────────

if __name__ == "__main__":
    sys.exit(main())
