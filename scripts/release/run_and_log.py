from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
import subprocess
import sys
import time
from collections.abc import Mapping, Sequence
from datetime import UTC, datetime
from pathlib import Path
from typing import IO, NoReturn

LOG_SCHEMA_VERSION = 1
SPAWN_FAILURE_EXIT_CODE = 127
ALLOWED_ENVIRONMENT_OVERRIDES = frozenset(
    {
        "DOCWEN_BUILD_VERSION",
        "DOCWEN_CONFIG_DIR",
        "DOCWEN_LOG_DIR",
        "DOCWEN_LOG_TO_TEMP",
        "DOCWEN_PYTEST_REPORT_DIR",
        "DOCWEN_PYTEST_XDIST",
        "DOCWEN_PYTEST_XDIST_WORKERS",
        "PYTHONIOENCODING",
        "PYTHONUTF8",
        "QT_QPA_PLATFORM",
        "SOURCE_DATE_EPOCH",
    }
)

_REPARSE_ATTRIBUTE = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x0400)


class CommandLogError(RuntimeError):
    """A fail-closed command evidence logging error."""


def _fail(code: str, detail: str | None = None) -> NoReturn:
    if detail:
        raise CommandLogError(f"{code}: {detail}")
    raise CommandLogError(code)


def _is_reparse(stats: os.stat_result) -> bool:
    return bool(getattr(stats, "st_file_attributes", 0) & _REPARSE_ATTRIBUTE)


def _absolute_path(path: Path) -> Path:
    return Path(os.path.abspath(os.fspath(path)))


def _assert_existing_chain_without_links(path: Path) -> None:
    absolute = _absolute_path(path)
    current = Path(absolute.anchor)
    for part in absolute.parts[1:]:
        current /= part
        try:
            stats = current.lstat()
        except FileNotFoundError:
            _fail("path_missing", str(current))
        if stat.S_ISLNK(stats.st_mode) or _is_reparse(stats):
            _fail("linked_path_rejected", str(current))


def _safe_directory(path: Path, *, label: str) -> Path:
    absolute = _absolute_path(path)
    _assert_existing_chain_without_links(absolute)
    stats = absolute.lstat()
    if not stat.S_ISDIR(stats.st_mode):
        _fail("directory_required", f"{label}={absolute}")
    return absolute.resolve(strict=True)


def _claim_output(path: Path) -> tuple[Path, IO[str]]:
    absolute = _absolute_path(path)
    parent = _safe_directory(absolute.parent, label="output_parent")
    output = parent / absolute.name
    try:
        handle = output.open("x", encoding="utf-8", newline="\n")
    except FileExistsError:
        _fail("output_exists", str(output))
    return output, handle


def parse_environment_overrides(values: Sequence[str]) -> dict[str, str]:
    overrides: dict[str, str] = {}
    for item in values:
        if "=" not in item:
            _fail("invalid_environment_override", item)
        name, value = item.split("=", 1)
        if name not in ALLOWED_ENVIRONMENT_OVERRIDES:
            _fail("environment_override_not_allowed", name)
        if name in overrides:
            _fail("duplicate_environment_override", name)
        if "\0" in value:
            _fail("invalid_environment_override", name)
        overrides[name] = value
    return dict(sorted(overrides.items()))


def _utc_now() -> str:
    return datetime.now(UTC).isoformat().replace("+00:00", "Z")


def _stream_payload(value: bytes) -> dict[str, object]:
    try:
        text = value.decode("utf-8", errors="strict")
        decoding_errors = False
    except UnicodeDecodeError:
        text = value.decode("utf-8", errors="replace")
        decoding_errors = True
    return {
        "bytes": len(value),
        "sha256": hashlib.sha256(value).hexdigest(),
        "encoding": "utf-8",
        "decodingErrors": decoding_errors,
        "text": text,
    }


def _write_claimed_log(handle: IO[str], payload: Mapping[str, object]) -> None:
    json.dump(payload, handle, ensure_ascii=False, sort_keys=True, indent=2)
    handle.write("\n")
    handle.flush()
    os.fsync(handle.fileno())


def run_logged_command(
    argv: Sequence[str],
    *,
    cwd: Path,
    output: Path,
    environment_overrides: Mapping[str, str] | None = None,
) -> tuple[int, dict[str, object]]:
    command = list(argv)
    if not command or any("\0" in argument for argument in command):
        _fail("command_required")
    workdir = _safe_directory(cwd, label="cwd")
    overrides = parse_environment_overrides(
        [f"{name}={value}" for name, value in (environment_overrides or {}).items()]
    )
    output_path, handle = _claim_output(output)

    started_at = _utc_now()
    started_monotonic = time.monotonic_ns()
    stdout = b""
    stderr = b""
    exit_code: int | None = None
    spawn_error: dict[str, str] | None = None
    try:
        child_environment = os.environ.copy()
        child_environment.update(overrides)
        try:
            completed = subprocess.run(
                command,
                cwd=workdir,
                env=child_environment,
                check=False,
                capture_output=True,
            )
            stdout = completed.stdout
            stderr = completed.stderr
            exit_code = completed.returncode
            returned_exit_code = completed.returncode
        except OSError as exc:
            spawn_error = {"type": type(exc).__name__, "message": str(exc)}
            stderr = str(exc).encode("utf-8", errors="replace")
            returned_exit_code = SPAWN_FAILURE_EXIT_CODE
        ended_at = _utc_now()
        duration_seconds = (time.monotonic_ns() - started_monotonic) / 1_000_000_000
        payload: dict[str, object] = {
            "schemaVersion": LOG_SCHEMA_VERSION,
            "argv": command,
            "cwd": str(workdir),
            "environmentOverrides": overrides,
            "startedAt": started_at,
            "endedAt": ended_at,
            "durationSeconds": duration_seconds,
            "exitCode": exit_code,
            "stdout": _stream_payload(stdout),
            "stderr": _stream_payload(stderr),
            "spawnError": spawn_error,
        }
        _write_claimed_log(handle, payload)
    except BaseException:
        handle.close()
        raise
    else:
        handle.close()
    return returned_exit_code, {"output": str(output_path), **payload}


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run one command and write an exclusive, byte-accurate JSON evidence log.",
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--cwd", type=Path, required=True)
    parser.add_argument(
        "--env",
        action="append",
        default=[],
        metavar="NAME=VALUE",
        help="Recorded environment override from the fixed allow-list; may be repeated.",
    )
    parser.add_argument("command", nargs=argparse.REMAINDER, help="Command after --.")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    command = list(args.command)
    if command[:1] == ["--"]:
        command = command[1:]
    try:
        overrides = parse_environment_overrides(args.env)
        exit_code, _payload = run_logged_command(
            command,
            cwd=args.cwd,
            output=args.output,
            environment_overrides=overrides,
        )
    except CommandLogError as exc:
        print(f"run_and_log_error: {exc}", file=sys.stderr)
        return 2
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
