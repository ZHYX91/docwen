"""Request construction helpers for protocol 3 execution commands.

This module owns the content-derived route, input-reference, option and output
policy projection used by the orchestration functions in :mod:`convert`.
"""

from __future__ import annotations

import argparse
from collections.abc import Collection
from typing import Any

from docwen_application.controller import CapabilityUnavailableError
from docwen_cli.i18n import cli_t, get_cli_locale
from docwen_core.detection import FileAdmissionError, inspect_file
from docwen_core.detection.ooxml_signature import OOXML_SIGNATURE_INFO_METADATA_KEY
from docwen_core.models.file_inspection import (
    FILE_ADMISSION_ACCEPTANCE_METADATA_KEY,
    FILE_INSPECTION_METADATA_KEY,
    FileInspection,
    make_admission_acceptance,
)
from docwen_core.models.request import FileRef, OutputPolicy


def resolve_cli_action(args: argparse.Namespace) -> str:
    """Return the resolved action name for this CLI invocation."""

    return str(getattr(args, "action", "") or "").strip().lower()


def _translation_or_default(key: str, default: str) -> str:
    value = cli_t(key, default="")
    return default if value == key or not value else value


def project_route_options(
    options: dict[str, Any],
    *,
    route_id: str,
    route_options: Collection[str],
    configured_ocr_language: str | None = None,
    ocr_requested: bool = False,
) -> dict[str, Any]:
    """Validate and project options through one canonical runtime route.

    Every key produced from an explicit CLI option must be declared by the
    resolved route. Runtime-derived defaults are only injected when that same
    route declares the key. This keeps all execution paths fail-closed without
    reconstructing route semantics from source categories or action names.
    """

    supported = frozenset(route_options)
    unsupported = sorted(set(options) - supported)
    if unsupported:
        raise ValueError(f"Canonical runtime route {route_id} does not declare option(s): {', '.join(unsupported)}")

    projected = dict(options)
    if "to_md_enable_ocr" in supported:
        # ``--ocr`` is deliberately opt-in at the CLI boundary.  Project the
        # negative default only after resolving a route that explicitly owns
        # the option, so a user config cannot silently turn OCR back on and
        # unrelated routes never receive a guessed option.
        projected.setdefault("to_md_enable_ocr", bool(ocr_requested))
    if ocr_requested and configured_ocr_language and "ocr_language" in supported:
        projected.setdefault("ocr_language", configured_ocr_language)
    if "locale" in supported:
        projected.setdefault("locale", get_cli_locale())
    if "yaml_key_labels" in supported:
        projected.setdefault(
            "yaml_key_labels",
            {
                "title": _translation_or_default("yaml_keys.title", "title"),
                "subtitle": _translation_or_default("yaml_keys.subtitle", "subtitle"),
            },
        )
    return projected


def redacted_options(options: dict[str, Any]) -> dict[str, Any]:
    """Return options safe for dry-run presentation."""

    redacted = dict(options)
    if "spreadsheet_password" in redacted:
        redacted["spreadsheet_password"] = "<redacted>"
    return redacted


def file_ref_for_runtime(
    file_path: str,
    inspection: FileInspection | None = None,
    *,
    explicit_acceptance: bool = False,
) -> FileRef:
    """Build a ``FileRef`` from one frozen content-derived inspection."""

    inspection = inspection or inspect_file(file_path)
    if not inspection.may_execute:
        raise FileAdmissionError(inspection)
    metadata: dict[str, Any] = {FILE_INSPECTION_METADATA_KEY: inspection.to_dict()}
    if inspection.requires_explicit_acceptance:
        if not explicit_acceptance:
            raise ValueError("This file requires explicit format acceptance before an execution request can be built.")
        metadata[FILE_ADMISSION_ACCEPTANCE_METADATA_KEY] = make_admission_acceptance(inspection)
    signature_info = dict(inspection.ooxml_signature)
    if signature_info.get("state") != "not_applicable":
        metadata[OOXML_SIGNATURE_INFO_METADATA_KEY] = signature_info
    return FileRef(
        path=inspection.file_path,
        format=inspection.detected_format,
        category=inspection.workflow_category,
        warning_message=inspection.warning_message,
        size_bytes=inspection.size_bytes,
        metadata=metadata,
    )


def configured_ocr_language(controller: Any) -> str | None:
    """Read the configured OCR language through the injected config port."""

    config_port = getattr(controller, "config_port", None)
    if config_port is None:
        return None
    try:
        value = config_port.get("image.ocr_language", None)
    except Exception as exc:
        raise CapabilityUnavailableError("OCR language configuration could not be read.") from exc
    return str(value).strip() if value else None


def public_command(args: argparse.Namespace) -> str:
    """Return the frozen protocol 3 command path for presentation."""

    return str(getattr(args, "command_path", None) or getattr(args, "command", "convert"))


def output_policy(args: argparse.Namespace) -> OutputPolicy:
    """Project explicit protocol 3 destination and overwrite intent."""

    output_path = getattr(args, "output_path", None)
    output_dir = getattr(args, "output_dir", None)
    command_path = str(getattr(args, "command_path", ""))
    is_validate = command_path in {"validate", "batch validate"}
    if output_path:
        return OutputPolicy(
            output_path=str(output_path),
            overwrite_mode="overwrite" if getattr(args, "overwrite", False) else "error",
        )
    if output_dir:
        return OutputPolicy(
            output_dir=str(output_dir),
            overwrite_mode="overwrite" if getattr(args, "overwrite", False) else "error",
        )
    if is_validate:
        return OutputPolicy(overwrite_mode="error", write_artifacts=False)
    return OutputPolicy(overwrite_mode="error")


__all__ = [
    "configured_ocr_language",
    "file_ref_for_runtime",
    "output_policy",
    "project_route_options",
    "public_command",
    "redacted_options",
    "resolve_cli_action",
]
