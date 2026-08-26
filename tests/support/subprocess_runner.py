from __future__ import annotations

import json
import os
import shlex
import subprocess
import time
from collections.abc import Sequence
from pathlib import Path
from typing import Any

_REPORT_DIR_ENV = "DOCWEN_PYTEST_REPORT_DIR"
DEFAULT_SUBPROCESS_TIMEOUT_SECONDS = 60.0


def _stringify_command(args: Sequence[str] | str) -> str:
    if isinstance(args, str):
        return args
    return " ".join(shlex.quote(str(part)) for part in args)


def _subprocess_record_path(report_dir: str | os.PathLike[str] | None = None) -> Path | None:
    resolved_dir = report_dir or os.environ.get(_REPORT_DIR_ENV)
    if not resolved_dir:
        return None
    worker_suffix = os.environ.get("PYTEST_XDIST_WORKER") or f"pid-{os.getpid()}"
    return Path(resolved_dir) / f"subprocess_runs-{worker_suffix}.jsonl"


def append_subprocess_run_record(record: dict[str, Any]) -> None:
    record_path = _subprocess_record_path()
    if record_path is None:
        return
    record_path.parent.mkdir(parents=True, exist_ok=True)
    with record_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, ensure_ascii=False) + "\n")


def clear_subprocess_run_records(report_dir: str | os.PathLike[str]) -> None:
    base = Path(report_dir)
    if not base.exists():
        return
    for path in base.glob("subprocess_runs-*.jsonl"):
        path.unlink()


def load_subprocess_run_records(report_dir: str | os.PathLike[str] | None = None) -> list[dict[str, Any]]:
    record_path = _subprocess_record_path(report_dir)
    if record_path is None:
        return []
    records: list[dict[str, Any]] = []
    for path in sorted(record_path.parent.glob("subprocess_runs-*.jsonl")):
        for line in path.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            records.append(json.loads(line))
    return records


def run_subprocess(
    args: Sequence[str] | str,
    *,
    cwd: str | os.PathLike[str] | None = None,
    env: dict[str, str] | None = None,
    timeout: float | None = DEFAULT_SUBPROCESS_TIMEOUT_SECONDS,
    capture_output: bool = True,
    text: bool = True,
    encoding: str = "utf-8",
    errors: str = "replace",
    check: bool = False,
    **kwargs: Any,
) -> subprocess.CompletedProcess[str] | subprocess.CompletedProcess[bytes]:
    if timeout is None or timeout <= 0:
        raise ValueError("test subprocess timeout must be a positive number")
    start = time.perf_counter()
    completed: subprocess.CompletedProcess[str] | subprocess.CompletedProcess[bytes] | None = None
    timed_out = False
    try:
        run_kwargs: dict[str, Any] = {
            "cwd": cwd,
            "env": env,
            "timeout": timeout,
            "capture_output": capture_output,
            "text": text,
            "check": check,
            **kwargs,
        }
        if text:
            run_kwargs["encoding"] = encoding
            run_kwargs["errors"] = errors
        completed = subprocess.run(args, **run_kwargs)
        return completed
    except subprocess.TimeoutExpired:
        timed_out = True
        raise
    finally:
        duration_seconds = round(time.perf_counter() - start, 6)
        append_subprocess_run_record(
            {
                "argv": [str(part) for part in args] if not isinstance(args, str) else [args],
                "command": _stringify_command(args),
                "cwd": str(Path(cwd).resolve()) if cwd is not None else None,
                "returncode": None if completed is None else completed.returncode,
                "timed_out": timed_out,
                "timeout_seconds": timeout,
                "duration_seconds": duration_seconds,
                "text": text,
                "capture_output": capture_output,
            }
        )
