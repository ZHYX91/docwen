"""Protocol 3 execution adaptation and destination preflight."""

from __future__ import annotations

import argparse
import os
import secrets
import sys
from pathlib import Path
from typing import Any

from docwen_cli.capabilities import normalize_target_format, resolve_optimization_action
from docwen_cli.commands.convert import execute_convert
from docwen_cli.commands.execution_routes import route_for_public_command
from docwen_cli.exit_codes import ExitCode
from docwen_cli.presenters.json_presenter import JsonPresenter
from docwen_core.models.document_node import canonical_source_tag, sanitize_node_label
from docwen_runtime.path_io import filesystem_path

_EXECUTION_DEFAULTS: dict[str, Any] = {
    "action": "",
    "to": "",
    "output": None,
    "output_path": None,
    "output_dir": None,
    "overwrite": False,
    "dry_run": False,
    "timeout": 600,
    "batch": False,
    "continue_on_error": False,
    "jobs": 1,
    "template": None,
    "optimization": None,
    "check": None,
    "extract_img": False,
    "no_extract_img": False,
    "ocr": False,
    "ocr_language": None,
    "image_mode": None,
    "image_link_style": None,
    "table_merge_strategy": None,
    "ocr_placement": None,
    "clean_numbering": None,
    "add_numbering": None,
    "heading_merge_mode": None,
    "heading_numbering_render_mode": None,
    "pages": None,
    "dpi": None,
    "mode": None,
    "keep_alpha": None,
    "spreadsheet_password_prompt": False,
    "allow_spreadsheet_protection_loss": False,
    "use_detected_format": False,
}

_DEFAULT_MARKDOWN_NUMBERING_SCHEME = "hierarchical_standard"


def _command_path(args: argparse.Namespace) -> str:
    command = str(args.command)
    child = getattr(args, f"{command}_command", None)
    return f"{command} {child}" if child else command


def _route(args: argparse.Namespace) -> str:
    path = _command_path(args)
    return route_for_public_command(path).action


def _prepare_args(args: argparse.Namespace) -> None:
    for name, value in _EXECUTION_DEFAULTS.items():
        if not hasattr(args, name):
            setattr(args, name, value)

    path = _command_path(args)
    action = _route(args)
    args.command_path = path
    args.action = action
    if path not in {"convert", "batch convert"}:
        # The resolved runtime route owns the unique target for named actions.
        args.to = ""
    args.batch = path.startswith("batch ")

    if not hasattr(args, "files"):
        args.files = [str(args.file)]

    if path == "validate":
        args.output_path = getattr(args, "report", None)
    elif path == "batch validate":
        args.output_dir = getattr(args, "report_dir", None)
    elif path == "split pdf" or path.startswith("batch "):
        args.output_dir = str(args.output_dir)
    else:
        output = getattr(args, "output", None)
        if path == "number markdown" and getattr(args, "in_place", False):
            output = str(Path(args.file))
            args.overwrite = True
            args.output_path = output
        elif output and (
            path == "number markdown"
            or (path == "convert" and normalize_target_format(str(getattr(args, "to", ""))) == "md")
        ):
            # Markdown publishes a complete document-node directory.  The CLI
            # destination therefore names its parent, never an exact .md file.
            args.output_dir = str(output)
            args.output_path = None
        else:
            args.output_path = str(output) if output else None

    if path == "number markdown":
        operation = str(getattr(args, "operation", ""))
        if operation == "add":
            args.clean_numbering = "keep"
            args.add_numbering = getattr(args, "scheme", None) or _DEFAULT_MARKDOWN_NUMBERING_SCHEME
        else:
            args.clean_numbering = "remove"
            args.add_numbering = "none"


def execute_execution(args: argparse.Namespace, controller: Any | None = None) -> int:
    """Adapt clean domain arguments to the existing application workflow."""

    if not getattr(args, "_docwen_preflight_done", False):
        result = preflight_execution(args)
        if result is not None:
            return result
    optimization = str(getattr(args, "optimization", "") or "")
    if optimization:
        args.action = resolve_optimization_action(controller, optimization)
    return execute_convert(args, controller)


def preflight_execution(args: argparse.Namespace) -> int | None:
    """Prepare and validate an execution request before runtime creation."""

    _prepare_args(args)
    if args.command_path == "number markdown":
        operation = str(getattr(args, "operation", ""))
        scheme = getattr(args, "scheme", None)
        if operation == "remove" and scheme:
            message = "--scheme is only valid with --operation add."
            if getattr(args, "json", False):
                JsonPresenter().present_error(
                    args.command_path,
                    message,
                    error_code="invalid_input",
                    details={"operation": operation, "scheme": str(scheme)},
                )
            else:
                print(f"Error: {message}", file=sys.stderr)
            return int(ExitCode.INVALID_INPUT)
    preflight = _preflight_destination(args)
    if preflight is not None:
        code, message, details = preflight
        if getattr(args, "json", False):
            JsonPresenter().present_error(
                args.command_path,
                message,
                error_code=code,
                details=details,
                hint="Use --overwrite only when replacing the target is intentional."
                if code == "output_exists"
                else None,
            )
        else:
            print(f"Error: {message}", file=sys.stderr)
        return int(ExitCode.CONFLICT if code in {"output_exists", "output_collision"} else ExitCode.INVALID_INPUT)
    args._docwen_preflight_done = True
    return None


def _preflight_destination(args: argparse.Namespace) -> tuple[str, str, dict[str, Any]] | None:
    """Fail before runtime startup or side effects when a destination is unsafe."""

    output_path = getattr(args, "output_path", None)
    if output_path:
        resolved = Path(output_path).expanduser().resolve(strict=False)
        parent = resolved.parent
        if not parent.is_dir():
            return "invalid_input", f"Output parent directory does not exist: {parent}", {"path": str(resolved)}
        inputs = [Path(value).expanduser().resolve(strict=False) for value in getattr(args, "files", [])]
        if resolved in inputs and not getattr(args, "in_place", False):
            return "invalid_input", "Output path must not replace an input file.", {"path": str(resolved)}
        if resolved.exists() and not getattr(args, "overwrite", False):
            return "output_exists", f"Output target already exists: {resolved}", {"path": str(resolved)}
        if not _directory_accepts_write_probe(parent):
            return "invalid_input", f"Output parent directory is not writable: {parent}", {"path": str(parent)}

    output_dir = getattr(args, "output_dir", None)
    if output_dir:
        resolved_dir = Path(output_dir).expanduser().resolve(strict=False)
        if resolved_dir.exists() and not resolved_dir.is_dir():
            return "invalid_input", f"Output directory is a file: {resolved_dir}", {"path": str(resolved_dir)}
        parent = resolved_dir.parent
        if not parent.is_dir():
            return "invalid_input", f"Output parent directory does not exist: {parent}", {"path": str(resolved_dir)}
        probe_dir = resolved_dir if resolved_dir.exists() else parent
        if not _directory_accepts_write_probe(probe_dir):
            target = "Output directory" if resolved_dir.exists() else "Output parent directory"
            return "invalid_input", f"{target} is not writable: {probe_dir}", {"path": str(probe_dir)}
        batch_collision = _preflight_batch_collisions(args, resolved_dir)
        if batch_collision is not None:
            return batch_collision
    return None


def _directory_accepts_write_probe(directory: Path) -> bool:
    """Check effective directory write/delete rights without trusting ``os.access``.

    On Windows, ``os.access(..., os.W_OK)`` can report a writable directory
    even when an explicit ACL denies the current user the rights needed to
    publish an output.  ``tempfile.TemporaryFile`` is intentionally not used:
    CPython retries candidate names indefinitely after ``PermissionError`` on
    an existing Windows directory.  A single exclusive ``os.open`` attempt
    fails immediately and ``O_TEMPORARY`` supplies delete-on-close semantics
    where Windows provides it.
    """

    probe_path = Path(directory) / f".__docwen-write-probe-{secrets.token_hex(12)}.tmp"
    native_probe_path = os.fspath(filesystem_path(probe_path, force_extended=sys.platform == "win32"))
    flags = os.O_CREAT | os.O_EXCL | os.O_WRONLY
    flags |= getattr(os, "O_BINARY", 0)
    flags |= getattr(os, "O_TEMPORARY", 0)
    descriptor: int | None = None
    created = False
    succeeded = False
    try:
        descriptor = os.open(native_probe_path, flags, 0o600)
        created = True
        succeeded = os.write(descriptor, b"\0") == 1
    except OSError:
        succeeded = False
    finally:
        if descriptor is not None:
            try:
                os.close(descriptor)
            except OSError:
                succeeded = False
        if created:
            try:
                os.unlink(native_probe_path)
            except FileNotFoundError:
                pass
            except OSError:
                succeeded = False
    return succeeded


def _preflight_batch_collisions(args: argparse.Namespace, output_dir: Path) -> tuple[str, str, dict[str, Any]] | None:
    """Reject deterministic many-to-one output collisions before runtime creation."""

    if getattr(args, "command_path", "") != "batch convert":
        return None
    target_format = normalize_target_format(str(getattr(args, "to", "")))
    markdown_target = target_format == "md"
    extension = target_format
    destinations: dict[str, tuple[Path, list[str]]] = {}
    for raw_input in getattr(args, "files", []):
        source = Path(raw_input).expanduser().resolve(strict=False)
        if markdown_target:
            source_tag = canonical_source_tag(source.suffix.lstrip(".") or "document")
            node_key = sanitize_node_label(f"{source.stem}_<timestamp>_from{source_tag}")
            destination = output_dir / node_key
        else:
            destination = output_dir / f"{source.stem}.{extension}"
        key = os.path.normcase(str(destination))
        if key not in destinations:
            destinations[key] = (destination, [])
        destinations[key][1].append(str(source))

    collisions = [inputs for _destination, inputs in destinations.values() if len(inputs) > 1]
    if collisions:
        return (
            "output_collision",
            "Multiple batch inputs resolve to the same output target.",
            {"collisions": collisions},
        )

    if not markdown_target and not getattr(args, "overwrite", False):
        existing = [str(path) for path, _inputs in destinations.values() if path.exists()]
        if existing:
            return (
                "output_exists",
                "One or more batch output targets already exist.",
                {"paths": existing},
            )
    return None


__all__ = ["execute_execution", "preflight_execution"]
