"""Text output presenter — formats results for human-readable terminal output."""

from __future__ import annotations

import sys
from typing import Any


class TextPresenter:
    """Formats conversion results as plain text for terminal output.

    Writes result messages to stdout; progress/errors go to stderr
    following the stdout/stderr contract in CLI契约规范.md §7.
    """

    def __init__(self, *, quiet: bool = False, verbose: bool = False) -> None:
        self._quiet = quiet
        self._verbose = verbose

    # ── Single result ──────────────────────────────────────────────

    def present_single(self, result: Any) -> None:
        """Present a single-file conversion result."""
        success = getattr(result, "success", False)
        if success:
            self._output(self._format_success(result))
            self._present_warnings(result)
        else:
            self._error(self._format_failure(result))

    def _format_success(self, result: Any) -> str:
        """Format a successful conversion result."""
        artifacts = getattr(result, "artifacts", [])
        if artifacts:
            primary = artifacts[0]
            output_path = (
                getattr(primary, "staging_path", "") or getattr(primary, "path", "") or getattr(primary, "location", "")
            )
            if output_path:
                return f"转换成功 → {output_path}"
        return "转换成功"

    def _format_failure(self, result: Any) -> str:
        """Format a failed conversion result."""
        error = getattr(result, "error", None)
        if error is not None:
            msg = getattr(error, "message", str(error))
            return f"错误: {msg}"
        return "错误: 转换失败"

    # ── Batch result ──────────────────────────────────────────────

    def present_batch(self, results: list[Any], *, input_files: list[str] | None = None) -> None:
        """Present a batch of conversion results with summary."""
        total = len(results)
        success = sum(1 for r in results if getattr(r, "success", False))
        failed = total - success

        # Per-file results
        for idx, r in enumerate(results):
            if not self._quiet:
                input_file = input_files[idx] if input_files is not None and idx < len(input_files) else None
                self._present_batch_item(r, input_file=input_file)

        # Summary
        summary = f"\n总计: {total} 文件  |  成功: {success}  |  失败: {failed}"
        self._output(summary)

    def _present_batch_item(self, result: Any, *, input_file: str | None = None) -> None:
        """Present a single item in a batch."""
        input_path = input_file or getattr(result, "input_file", "") or getattr(result, "task_id", "?")

        if getattr(result, "success", False):
            self._output(f"  OK  {input_path}")
            self._present_warnings(result, input_path=input_path)
        else:
            error = getattr(result, "error", None)
            msg = getattr(error, "message", "未知错误") if error else "未知错误"
            self._error(f"  FAIL  {input_path}: {msg}")

    # ── Generic messages ──────────────────────────────────────────

    def present_message(self, message: str) -> None:
        """Print a plain informational message to stdout."""
        self._output(message)

    def present_error_message(self, message: str) -> None:
        """Print an error message to stderr (text mode contract)."""
        self._error(f"错误: {message}")

    def _present_warnings(self, result: Any, *, input_path: str = "") -> None:
        """Write successful-result warning diagnostics to stderr."""
        for diagnostic in getattr(result, "diagnostics", []):
            if getattr(diagnostic, "level", "") != "warning":
                continue
            code = str(getattr(diagnostic, "code", "") or "")
            message = str(getattr(diagnostic, "message", "") or "")
            location = str(getattr(diagnostic, "location", "") or "")
            code_label = f" [{code}]" if code else ""
            file_label = f"{input_path}: " if input_path else ""
            location_label = f" ({location})" if location else ""
            self._warning(f"警告{code_label}: {file_label}{message}{location_label}")

    # ── Internal I/O ──────────────────────────────────────────────

    def _output(self, text: str) -> None:
        if self._quiet:
            return
        print(text)

    def _error(self, text: str) -> None:
        if self._quiet:
            return
        print(text, file=sys.stderr)

    def _warning(self, text: str) -> None:
        if self._quiet:
            return
        print(text, file=sys.stderr)
