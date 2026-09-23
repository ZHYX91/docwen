"""Freeze GUI input and option state into runtime requests and safe history context.

This builder reads the view models through explicit dependencies. It has no
window, dialogs, thread ownership, or result presentation responsibilities.
"""

from __future__ import annotations

import logging
import uuid
from collections.abc import Callable, Sequence
from copy import deepcopy
from dataclasses import replace
from pathlib import Path
from typing import TYPE_CHECKING, Any, Literal

from docwen_application.postprocessing import prepare_postprocess_options
from docwen_core.models.request import POSTPROCESS_PROOFREAD_OPTION
from docwen_gui.file_admission_i18n import render_file_inspection_message
from docwen_gui.i18n import t as _t
from docwen_gui.path_identity import normalize_path

if TYPE_CHECKING:
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy
    from docwen_gui.view_models.batch_list_vm import BatchListViewModel
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel

logger = logging.getLogger(__name__)

_MARKDOWN_TARGET_FORMATS: frozenset[str] = frozenset({"md", "markdown"})
_DOCUMENT_TEMPLATE_TARGETS: frozenset[str] = frozenset({"docx", "doc", "odt", "rtf", "wps", "pdf"})
_SPREADSHEET_TEMPLATE_TARGETS: frozenset[str] = frozenset({"xlsx", "xls", "ods", "csv"})
_PROOFREAD_ACTIONS: frozenset[str] = frozenset({"validate"})
_PROOFREAD_GUI_OPTION_ALIASES: dict[str, str] = {
    "symbol_pairing": "enable_symbol_pairing",
    "symbol_correction": "enable_symbol_correction",
    "typos_rule": "enable_typos_rule",
    "sensitive_word": "enable_sensitive_word",
}
_OUTPUT_DATE_SUBFOLDER_TOKENS: dict[str, str] = {
    "%Y-%m-%d": "iso",
    "%Y%m%d": "compact",
    "%Y年%m月%d日": "chinese",
}


class OutputPolicyConfigError(RuntimeError):
    """Raised when persisted output settings cannot be read safely."""


def _redacted_request_options(options: dict[str, Any]) -> dict[str, Any]:
    """Keep execution secrets out of GUI retry/history context."""

    redacted = deepcopy(options)
    if "spreadsheet_password" in redacted:
        redacted["spreadsheet_password"] = "<redacted>"
    postprocess = redacted.pop(POSTPROCESS_PROOFREAD_OPTION, None)
    if isinstance(postprocess, dict):
        redacted["proofread"] = {
            "enabled": True,
            "options": deepcopy(postprocess),
        }
    return redacted


def _to_markdown_locale_options(
    options: dict[str, Any],
    *,
    target_format: str,
    action_name: str = "",
    route_options: Sequence[str] | None = None,
) -> dict[str, Any]:
    from docwen_gui.i18n import get_locale

    enriched = dict(options)
    if route_options is not None:
        supported = frozenset(route_options)
        if "locale" in supported:
            enriched.setdefault("locale", get_locale())
        if target_format not in _MARKDOWN_TARGET_FORMATS or "yaml_key_labels" not in supported:
            return enriched
    elif target_format not in _MARKDOWN_TARGET_FORMATS:
        return enriched
    elif action_name:
        # Named routes must provide their canonical option surface. Do not
        # infer support from an action string.
        return enriched
    else:
        enriched.setdefault("locale", get_locale())
    enriched.setdefault(
        "yaml_key_labels",
        {
            "title": _t("yaml_keys.title", default="title"),
            "subtitle": _t("yaml_keys.subtitle", default="subtitle"),
        },
    )
    return enriched


def _normalize_proofread_action_options(options: dict[str, Any], *, action_name: str) -> dict[str, Any]:
    if action_name and action_name not in _PROOFREAD_ACTIONS:
        return options
    normalized = dict(options)
    for gui_key, plugin_key in _PROOFREAD_GUI_OPTION_ALIASES.items():
        if gui_key not in normalized:
            continue
        value = normalized.pop(gui_key)
        normalized.setdefault(plugin_key, bool(value))
    return normalized


def _route_scoped_options(
    options: dict[str, Any],
    *,
    route_options: Sequence[str] | None,
) -> dict[str, Any]:
    if route_options is None:
        return dict(options)
    supported = frozenset(route_options)
    return {key: value for key, value in options.items() if key in supported or key == POSTPROCESS_PROOFREAD_OPTION}


def _output_date_subfolder_token(date_folder_format: str) -> str:
    """Map GUI output date formats to runtime output-policy tokens."""
    return _OUTPUT_DATE_SUBFOLDER_TOKENS.get(date_folder_format, date_folder_format)


class ExecutionRequestBuilder:
    """Build single, batch and aggregate requests with the same option policy."""

    def __init__(
        self,
        view_model: MainWindowViewModel,
        batch_list_vm: BatchListViewModel,
        *,
        file_contexts: Callable[[], dict[str, tuple[str, str]]],
        selected_template: Callable[[], tuple[str, str] | None],
    ) -> None:
        self._view_model = view_model
        self._batch_list_vm = batch_list_vm
        self._file_contexts = file_contexts
        self._selected_template = selected_template

    def single(
        self,
        *,
        file_path: str,
        target_format: str,
        action_name: str,
        options: dict[str, Any],
        route_options: Sequence[str] | None = None,
    ) -> tuple[ConversionRequest, dict[str, Any]]:
        return self._build(
            file_paths=[file_path],
            mode="single",
            target_format=target_format,
            action_name=action_name,
            options=options,
            route_options=route_options,
        )

    def batch(
        self,
        *,
        file_paths: Sequence[str],
        target_format: str,
        action_name: str,
        options: dict[str, Any],
        route_options: Sequence[str] | None = None,
    ) -> tuple[ConversionRequest, dict[str, Any]]:
        return self._build(
            file_paths=file_paths,
            mode="batch",
            target_format=target_format,
            action_name=action_name,
            options=options,
            route_options=route_options,
        )

    def aggregate(
        self,
        *,
        file_paths: Sequence[str],
        target_format: str,
        action_name: str,
        options: dict[str, Any],
        route_options: Sequence[str] | None = None,
    ) -> tuple[ConversionRequest, dict[str, Any]]:
        return self._build(
            file_paths=file_paths,
            mode="aggregate",
            target_format=target_format,
            action_name=action_name,
            options=options,
            route_options=route_options,
        )

    def _build(
        self,
        *,
        file_paths: Sequence[str],
        mode: Literal["single", "batch", "aggregate"],
        target_format: str,
        action_name: str,
        options: dict[str, Any],
        route_options: Sequence[str] | None,
    ) -> tuple[ConversionRequest, dict[str, Any]]:
        from docwen_core.models.request import ConversionRequest

        source_paths = [str(Path(path)) for path in file_paths]
        request_options = deepcopy(options)
        if mode != "aggregate":
            source_context = self.file_context(source_paths[0]) if source_paths else None
            request_options = self._merge_template_options(
                target_format,
                request_options,
                source_category=source_context[1] if source_context else "other",
                action_name=action_name,
            )
            request_options = _normalize_proofread_action_options(request_options, action_name=action_name)
            request_options = prepare_postprocess_options(
                request_options,
                source_format=source_context[0] if source_context else "",
                target_format=target_format,
                action_name=action_name,
            )
            request_options = _to_markdown_locale_options(
                request_options,
                target_format=target_format,
                action_name=action_name,
                route_options=route_options,
            )
        request_options = _route_scoped_options(request_options, route_options=route_options)
        output_policy = self.output_policy()
        request = ConversionRequest(
            request_id=str(uuid.uuid4()),
            input_refs=[self.file_ref(path) for path in source_paths],
            target_format=target_format,
            action_name=action_name,
            options=request_options,
            output_policy=output_policy,
        )
        normalized_paths = [normalize_path(path) for path in source_paths]
        context: dict[str, Any] = {
            "request_id": request.request_id,
            "file_path": normalized_paths[0] if normalized_paths else "",
            "target_format": target_format,
            "action_name": action_name,
            "options": _redacted_request_options(request_options),
            "open_after_done": output_policy.open_after_done,
        }
        if mode == "single":
            context["display_name"] = Path(source_paths[0]).name
        else:
            context.update(
                file_paths=normalized_paths,
                total_count=len(source_paths),
                display_name=(
                    _t("info_area.batch_name", "Batch processing ({count} files)", count=len(source_paths))
                    if mode == "batch"
                    else _t("info_area.aggregate_name", "Merge ({count} files)", count=len(source_paths))
                ),
            )
            context[mode] = True
        return request, context

    def file_context(self, file_path: str) -> tuple[str, str] | None:
        """Read current routing state, including lists replaced after construction."""
        context = self._file_contexts().get(normalize_path(file_path))
        if context is not None:
            return context
        entry = self._batch_list_vm.get_file_entry(file_path)
        if entry is not None:
            return entry.detected_format.lower(), entry.workflow_category.lower()
        return None

    def _merge_template_options(
        self,
        target_format: str,
        options: dict[str, Any],
        *,
        source_category: str,
        action_name: str,
    ) -> dict[str, Any]:
        """Add selected template metadata for Markdown document/spreadsheet targets."""
        merged = dict(options)
        target = str(target_format or "").lower()
        if action_name or "template_name" in merged:
            return merged
        from .view_models.interaction import FileCapability, resolve_capabilities

        if FileCapability.TEMPLATE_SELECTION not in resolve_capabilities(source_category):
            return merged
        selected = self._selected_template()
        if selected is None:
            raise ValueError(_t("settings.templates.choose_template", "Choose an enabled template in the right panel"))
        template_type, template_id = selected
        if target in _DOCUMENT_TEMPLATE_TARGETS and template_type == "docx":
            merged["template_name"] = template_id
        if target in _SPREADSHEET_TEMPLATE_TARGETS and template_type == "xlsx":
            merged["template_name"] = template_id
        return merged

    def file_ref(
        self,
        source_path: str,
    ) -> FileRef:
        """Build a runtime ref without discarding the ingress inspection.

        Routing may normalize the runtime format/category (notably TXT to the
        Markdown workflow), but warning and inspection facts must remain
        attached so application/runtime admission can enforce the same
        decision without opening and guessing the file again.
        """
        from docwen_core.models.file_ref import FileRef

        normalized = normalize_path(source_path)
        source_ref = next(
            (ref for ref in self._view_model.files if normalize_path(getattr(ref, "path", "")) == normalized),
            None,
        )
        if source_ref is not None:
            return replace(source_ref, path=source_path, metadata=deepcopy(source_ref.metadata))

        entry = self._batch_list_vm.get_file_entry(source_path)
        if entry is not None:
            return FileRef(
                path=source_path,
                format=entry.detected_format,
                category=entry.workflow_category,
                warning_message=entry.warning_message or "",
                size_bytes=entry.size_bytes,
                metadata=deepcopy(entry.metadata),
            )

        # Programmatic callers that bypass the visual list still cross the
        # same Core admission boundary here; no suffix-derived FileRef is ever
        # manufactured.
        from docwen_core.detection import FileAdmissionError, inspect_file
        from docwen_core.detection.ooxml_signature import OOXML_SIGNATURE_INFO_METADATA_KEY
        from docwen_core.models import FILE_INSPECTION_METADATA_KEY

        inspection = inspect_file(source_path)
        if not inspection.may_execute:
            raise FileAdmissionError(inspection)
        return FileRef(
            path=inspection.file_path,
            format=inspection.detected_format,
            category=inspection.workflow_category,
            warning_message=render_file_inspection_message(inspection),
            size_bytes=inspection.size_bytes,
            metadata={
                FILE_INSPECTION_METADATA_KEY: inspection.to_dict(),
                OOXML_SIGNATURE_INFO_METADATA_KEY: dict(inspection.ooxml_signature),
            },
        )

    def output_policy(self) -> OutputPolicy:
        from docwen_core.models.request import OutputPolicy

        controller = self._view_model.controller
        cfg_port = getattr(controller, "config_port", None) if controller is not None else None
        if cfg_port is None:
            return OutputPolicy()

        try:
            mode = str(cfg_port.get("output.directory.mode", "source") or "source")
            custom_path = str(cfg_port.get("output.directory.custom_path", "") or "").strip()
            create_date_subfolder = bool(cfg_port.get("output.directory.create_date_subfolder", False))
            date_folder_format = str(cfg_port.get("output.directory.date_folder_format", "%Y-%m-%d") or "%Y-%m-%d")
            auto_open_folder = bool(cfg_port.get("output.behavior.auto_open_folder", False))
        except Exception as exc:
            logger.exception("Unable to read persisted output settings")
            raise OutputPolicyConfigError("Persisted output settings are unavailable") from exc

        try:
            output_dir = (
                str(Path(custom_path).expanduser().resolve(strict=False)) if mode == "custom" and custom_path else None
            )
        except (OSError, RuntimeError, ValueError) as exc:
            logger.exception("Unable to resolve persisted custom output path")
            raise OutputPolicyConfigError("Persisted custom output path is invalid") from exc
        date_subfolder = _output_date_subfolder_token(date_folder_format) if create_date_subfolder else ""
        return OutputPolicy(
            output_dir=output_dir,
            date_subfolder=date_subfolder,
            overwrite_mode="rename",
            open_after_done=auto_open_folder,
        )
