"""Machine output presenter for the DocWen CLI protocol 3 envelope."""

from __future__ import annotations

import json
import logging
import sys
from pathlib import Path
from typing import Any

from docwen_cli.error_registry import public_error_code
from docwen_cli.protocol import category_for_error_code, make_envelope
from docwen_runtime.path_io import filesystem_path

logger = logging.getLogger(__name__)

type JsonValue = None | bool | int | float | str | list[JsonValue] | dict[str, JsonValue]


def _try_embed_proofread_report(artifact: Any) -> dict[str, Any] | None:
    """If *artifact* is a JSON proofread report on disk, read and return it.

    Returns ``None`` when the artifact is not a proofread report or the file
    cannot be read.  The caller merges the returned dict into the envelope
    ``data.details.proofread`` field so downstream consumers (Obsidian plugin, AI
    agents) can display issues without reading the staging file themselves.
    """
    meta = getattr(artifact, "metadata", None)
    if not isinstance(meta, dict):
        return None
    if "issues_found" not in meta:
        return None

    staging = getattr(artifact, "staging_path", "")
    if not staging:
        return None
    try:
        raw = filesystem_path(staging).read_text(encoding="utf-8")
        return json.loads(raw)
    except Exception as exc:
        logger.debug("Failed to embed proofread report from %s: %s", staging, exc)
        return None


class JsonPresenter:
    """Format command results as the exact protocol 3 envelope."""

    def __init__(
        self,
        *,
        quiet: bool = False,
        verbose: bool = False,
        include_timing: bool = False,
    ) -> None:
        self._quiet = quiet
        self._verbose = verbose
        self._include_timing = include_timing
        self._warnings: list[JsonValue] = []

    @property
    def warnings(self) -> list[JsonValue]:
        return list(self._warnings)

    def add_warning(self, warning: JsonValue) -> None:
        """Add a warning to be included in the next envelope."""
        if isinstance(warning, dict):
            payload = dict(warning)
            payload.setdefault("code", "warning")
            payload.setdefault("message", "")
            self._warnings.append(payload)
            return
        self._warnings.append({"code": "warning", "message": str(warning)})

    # ── Envelope factory ─────────────────────────────────────────

    def _make_envelope(
        self,
        command: str,
        success: bool,
        data: Any | None = None,
        error: dict[str, Any] | None = None,
        meta: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Build a protocol 3 JSON envelope."""
        envelope = make_envelope(
            command=command,
            success=success,
            data=data,
            error=error,
            warnings=list(self._warnings),
            meta=meta,
        )
        self._warnings.clear()
        return envelope

    def _emit(self, envelope: dict[str, Any]) -> None:
        json.dump(envelope, sys.stdout, ensure_ascii=False, indent=2)
        sys.stdout.write("\n")

    # ── Single result ─────────────────────────────────────────────

    def present_single(
        self,
        result: Any,
        *,
        command: str = "convert",
        action_name: str = "",
        input_files: list[str] | None = None,
    ) -> None:
        """Emit a single-file result as a JSON envelope.

        Args:
            result: ConversionResult to present.
            command: Normalized CLI command path (e.g. ``"convert"``).
            action_name: Resolved business action (e.g. ``""``,
                ``"gongwen"``, ``"validate"``).
        """
        success = getattr(result, "success", False)
        error_obj = self._result_error(result) if not success else None
        del action_name  # internal runtime actions are not part of protocol 3
        data = self._single_data(result, input_files=input_files) if success else None
        meta = self._timing_info(result) if self._include_timing else None
        self._warnings.extend(self._result_warning_payloads(result))

        envelope = self._make_envelope(
            command=command,
            success=success,
            data=data,
            error=error_obj,
            meta=meta,
        )
        self._emit(envelope)

    def _single_data(
        self,
        result: Any,
        *,
        input_files: list[str] | None = None,
    ) -> dict[str, Any]:
        """Build a command-neutral artifact result without leaking runtime actions."""
        output_file = ""
        artifacts = getattr(result, "artifacts", [])
        primary_artifact = None
        if artifacts:
            primary_artifact = artifacts[0]
            output_file = (
                getattr(primary_artifact, "staging_path", "")
                or getattr(primary_artifact, "path", "")
                or getattr(primary_artifact, "location", "")
            )

        metadata: dict[str, Any] = {}

        # ── Embed proofread report inline when present ──────────────
        proofread: dict[str, Any] | None = self._proofread_report_from_metrics(result)
        if primary_artifact is not None:
            proofread = proofread or _try_embed_proofread_report(primary_artifact)
        if proofread is not None:
            metadata["proofread"] = proofread

        raw_inputs = list(input_files or [])
        if not raw_inputs:
            result_input = str(getattr(result, "input_file", "") or "")
            if result_input:
                raw_inputs.append(result_input)
        normalized_inputs = [str(Path(value).expanduser().resolve(strict=False)) for value in raw_inputs]
        normalized_output = str(Path(output_file).expanduser().resolve(strict=False)) if output_file else ""
        artifact_items = [
            {
                "path": str(Path(self._artifact_path(item)).expanduser().resolve(strict=False)),
                "kind": str(getattr(item, "kind", "") or ""),
                "media_type": str(getattr(item, "media_type", "") or ""),
                "primary": bool(getattr(item, "is_primary", False)),
            }
            for item in artifacts
            if self._artifact_path(item)
        ]
        data: dict[str, Any] = {
            "inputs": normalized_inputs,
            "output": normalized_output,
            "artifacts": artifact_items,
            "details": metadata,
        }
        return data

    @staticmethod
    def _proofread_report_from_metrics(result: Any) -> dict[str, Any] | None:
        metrics = getattr(result, "metrics", None)
        extra = getattr(metrics, "extra", None)
        if not isinstance(extra, dict):
            return None
        report = extra.get("proofread_report")
        return dict(report) if isinstance(report, dict) else None

    def _result_error(self, result: Any) -> dict[str, Any]:
        """Extract error info from a failed result."""
        err = getattr(result, "error", None)
        if err is None:
            return self._typed_error(
                "internal_error",
                "The operation failed without a structured runtime error.",
            )
        error_code = getattr(err, "error_type", "internal_error")
        return self._typed_error(
            error_code,
            getattr(err, "message", "转换失败"),
            details=getattr(err, "diagnostic_code", None) or None,
        )

    @staticmethod
    def _typed_error(
        error_code: str,
        message: str,
        *,
        details: JsonValue = None,
        hint: str | None = None,
    ) -> dict[str, Any]:
        """Build one registered protocol error object."""
        canonical_code = public_error_code(error_code)
        return {
            "category": category_for_error_code(canonical_code).value,
            "code": canonical_code,
            "message": message,
            "details": details,
            "hint": hint,
        }

    # ── Error envelope (no result object) ────────────────────────

    def present_error(
        self,
        command: str,
        message: str,
        *,
        error_code: str = "unknown_error",
        details: JsonValue = None,
        hint: str | None = None,
    ) -> None:
        """Emit an error envelope whose details follow the JSON-value schema."""
        envelope = self._make_envelope(
            command=command,
            success=False,
            data=None,
            error=self._typed_error(
                error_code,
                message,
                details=details,
                hint=hint,
            ),
        )
        self._emit(envelope)

    # ── Batch result ─────────────────────────────────────────────

    def present_batch(
        self,
        results: list[Any],
        *,
        command: str = "batch convert",
        action_name: str = "",
        interrupted: bool = False,
        input_files: list[str] | None = None,
    ) -> None:
        """Emit a batch result envelope.

        Args:
            results: List of ConversionResult objects.
            command: Normalized CLI command path (e.g. ``"batch convert"``).
            action_name: Resolved business action (e.g. ``""``,
                ``"gongwen"``, ``"validate"``).
            interrupted: Whether the operation was interrupted.
            input_files: Optional original input files, aligned by index with
                ``results``.  This avoids mutating slotted result objects.
        """
        del action_name  # internal runtime actions are not part of protocol 3
        total = len(results)
        success_count = sum(1 for r in results if getattr(r, "success", False))
        failed_count = total - success_count

        batch_items: list[dict[str, Any]] = []
        for idx, r in enumerate(results):
            input_file = input_files[idx] if input_files is not None and idx < len(input_files) else None
            batch_items.append(self._batch_item(r, input_file=input_file))
            warning_file = input_file or getattr(r, "input_file", "")
            self._warnings.extend(self._result_warning_payloads(r, file_path=warning_file))

        data: dict[str, Any] = {
            "total": total,
            "processed": total,
            "succeeded": success_count,
            "failed": failed_count,
            "interrupted": interrupted,
            "results": batch_items,
        }

        meta = None
        if self._include_timing:
            total_duration_ms = sum(self._duration_milliseconds(r) for r in results)
            meta = {"duration_ms": total_duration_ms}

        if interrupted:
            top_level_error = self._typed_error(
                "operation_cancelled",
                "The batch operation was interrupted.",
                details={
                    "total": total,
                    "processed": total,
                    "succeeded": success_count,
                    "failed": failed_count,
                },
            )
        elif failed_count and success_count:
            top_level_error = self._typed_error(
                "batch_partial_failure",
                "One or more batch items failed.",
                details={
                    "total": total,
                    "succeeded": success_count,
                    "failed": failed_count,
                },
            )
        elif failed_count:
            first_error = next(
                (item["error"] for item in batch_items if isinstance(item.get("error"), dict)),
                None,
            )
            top_level_error = first_error or self._typed_error(
                "internal_error",
                "Every batch item failed without a structured error.",
            )
        else:
            top_level_error = None

        envelope = self._make_envelope(
            command=command,
            success=(failed_count == 0 and not interrupted),
            data=data,
            error=top_level_error,
            meta=meta,
        )
        self._emit(envelope)

    def _batch_item(self, result: Any, *, input_file: str | None = None) -> dict[str, Any]:
        """Format a single item inside the batch ``results`` array."""
        success = getattr(result, "success", False)
        file_path = input_file or getattr(result, "input_file", "")
        item: dict[str, Any] = {
            "input": str(Path(file_path).expanduser().resolve(strict=False)) if file_path else "",
            "success": success,
            "output": self._normalized_output_path(result),
            "error": None,
            "details": {},
        }
        # ── Embed proofread report inline ───────────────────────────
        if success:
            proofread = self._proofread_report_from_metrics(result)
            artifacts = getattr(result, "artifacts", [])
            if proofread is None and artifacts:
                proofread = _try_embed_proofread_report(artifacts[0])
            if proofread is not None:
                item["details"]["proofread"] = proofread
        if not success:
            err = getattr(result, "error", None)
            if err is not None:
                error_code = str(getattr(err, "error_type", "internal_error") or "internal_error")
                item["error"] = self._typed_error(
                    error_code,
                    getattr(err, "message", "转换失败"),
                    details=getattr(err, "diagnostic_code", None) or None,
                )
            else:
                item["error"] = self._typed_error(
                    "internal_error",
                    "The batch item failed without a structured runtime error.",
                )
        return item

    # ── Generic data output ──────────────────────────────────────

    def present_data(
        self,
        command: str,
        data: Any,
        *,
        success: bool = True,
    ) -> None:
        """Emit an envelope with arbitrary data (for list/schema etc.)."""
        envelope = self._make_envelope(
            command=command,
            success=success,
            data=data,
        )
        self._emit(envelope)

    # ── Helpers ──────────────────────────────────────────────────

    @staticmethod
    def _timing_info(result: Any) -> dict[str, Any]:
        """Build timing info dict from a result's metrics."""
        return {
            "duration_ms": JsonPresenter._duration_milliseconds(result),
        }

    @staticmethod
    def _duration_milliseconds(result: Any) -> float:
        """Extract a stable millisecond duration from a result."""
        metrics = getattr(result, "metrics", None)
        if metrics is not None:
            return round(float(getattr(metrics, "duration_ms", 0.0) or 0.0), 3)
        return 0.0

    @staticmethod
    def _duration_seconds(result: Any) -> float:
        """Extract duration in seconds from a result's metrics."""
        metrics = getattr(result, "metrics", None)
        if metrics is not None:
            ms = getattr(metrics, "duration_ms", 0.0)
            if ms:
                return round(ms / 1000.0, 3)
        return 0.0

    @staticmethod
    def _get_output_path(result: Any) -> str:
        """Extract output path from a result."""
        artifacts = getattr(result, "artifacts", [])
        if artifacts:
            return (
                getattr(artifacts[0], "staging_path", "")
                or getattr(artifacts[0], "path", "")
                or getattr(artifacts[0], "location", "")
            )
        return ""

    @staticmethod
    def _artifact_path(artifact: Any) -> str:
        return str(
            getattr(artifact, "staging_path", "")
            or getattr(artifact, "path", "")
            or getattr(artifact, "location", "")
            or ""
        )

    @staticmethod
    def _normalized_output_path(result: Any) -> str:
        path = JsonPresenter._get_output_path(result)
        return str(Path(path).expanduser().resolve(strict=False)) if path else ""

    @staticmethod
    def _result_warning_payloads(result: Any, *, file_path: str = "") -> list[dict[str, JsonValue]]:
        """Project warning diagnostics into the JSON envelope warning surface."""
        payloads: list[dict[str, JsonValue]] = []
        for diagnostic in getattr(result, "diagnostics", []):
            if getattr(diagnostic, "level", "") != "warning":
                continue
            payload: dict[str, JsonValue] = {
                "level": "warning",
                "code": str(getattr(diagnostic, "code", "") or ""),
                "message": str(getattr(diagnostic, "message", "") or ""),
                "location": str(getattr(diagnostic, "location", "") or ""),
            }
            if file_path:
                payload["file"] = str(file_path)
            payloads.append(payload)
        return payloads
