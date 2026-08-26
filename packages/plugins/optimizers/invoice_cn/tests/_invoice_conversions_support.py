"""Golden-style semantic tests for the Invoice plugin routes.

Covered routes:
- pdf → md with action_name="invoice_cn"  (InvoiceCnConverter)
- ofd → md with action_name="invoice_cn"  (InvoiceCnConverter / InvoiceData.xml)
- ofd → md with action_name="invoice_cn"  (fallback: content.xml)
- image → md with action_name="invoice_cn" (OCR-backed; can fail if OCR models/input fail)
- OCR option (to_md_enable_ocr=True) is accepted for scan-based PDFs
"""

from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.contract

_PROJECT_ROOT = Path(__file__).resolve().parents[5]

_INVOICE_CN_OLD_SYSTEM_FIXTURE = (
    _PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_invoice_cn_semantics.json"
)


def _ocr_outcome(status: str, text: str = "", message: str = "") -> Any:
    from docwen_core.text.ocr import OcrOutcome, OcrStatus

    return OcrOutcome(OcrStatus(status), text=text, message=message)


def _load_invoice_cn_old_system_fixture() -> dict[str, Any]:
    return json.loads(_INVOICE_CN_OLD_SYSTEM_FIXTURE.read_text(encoding="utf-8"))


def _parse_invoice_markdown(markdown: str) -> tuple[dict[str, str], list[str]]:
    yaml_fields: dict[str, str] = {}
    if markdown.startswith("---"):
        end = markdown.find("\n---", 3)
        if end != -1:
            for line in markdown[3:end].splitlines():
                if ":" in line and not line.startswith("  -"):
                    key, value = line.split(":", 1)
                    yaml_fields[key.strip()] = value.strip().strip("\"'")

    table_rows = [line for line in markdown.splitlines() if line.startswith("|") and "---" not in line]
    return yaml_fields, table_rows


def _build_fake_context(
    input_path: str,
    staging_dir: str,
    target_format: str,
    options: dict[str, Any] | None = None,
    action_name: str = "",
    source_format: str = "",
    *,
    pre_cancelled: bool = False,
) -> Any:
    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    @dataclass
    class FakeWorkspaceHandle:
        input_path: str
        staging_dir: str
        _counter: list[int] = field(default_factory=lambda: [0])
        _artifacts: list[ArtifactManifest] = field(default_factory=list)

        def create_artifact_path(self, kind: str, suffix: str) -> str:
            self._counter[0] += 1
            return str(Path(self.staging_dir) / f"{kind}_{self._counter[0]}{suffix}")

        def add_artifact(self, manifest: ArtifactManifest) -> None:
            self._artifacts.append(manifest)

    @dataclass
    class FakeProgressSink:
        events: list[tuple[float, str]] = field(default_factory=list)
        artifacts: list[tuple[str, str]] = field(default_factory=list)
        diagnostics: list[tuple[str, str, str, str]] = field(default_factory=list)

        def report_progress(self, percent: float, message: str = "") -> None:
            self.events.append((percent, message))

        def report_diagnostic(self, level: str, message: str, code: str = "", location: str = "") -> None:
            self.diagnostics.append((level, message, code, location))

        def report_artifact_ready(self, artifact_id: str, suggested_name: str) -> None:
            self.artifacts.append((artifact_id, suggested_name))

    @dataclass
    class FakeLogger:
        messages: list[str] = field(default_factory=list)

        def debug(self, msg: str, **kwargs: Any) -> None:
            pass

        def info(self, msg: str, **kwargs: Any) -> None:
            self.messages.append(msg)

        def warning(self, msg: str, **kwargs: Any) -> None:
            pass

        def error(self, msg: str, **kwargs: Any) -> None:
            self.messages.append(msg)

    @dataclass
    class FakeContext:
        request: ConversionRequest
        workspace: FakeWorkspaceHandle
        config: Any
        progress: FakeProgressSink
        cancellation: CancellationToken
        logger: FakeLogger

    detected_format = source_format or Path(input_path).suffix.lstrip(".")
    file_refs = [
        FileRef(
            path=input_path,
            format=detected_format,
            category="layout",
        )
    ]
    request = ConversionRequest(
        request_id="test-invoice-001",
        input_refs=file_refs,
        target_format=target_format,
        action_name=action_name,
        options=options or {},
        output_policy=OutputPolicy(),
    )
    config = type("FakeConfig", (), {"get": lambda self, k, d=None: d})()
    token = CancellationToken()
    if pre_cancelled:
        token.cancel("test cancellation")
    return FakeContext(
        request,
        FakeWorkspaceHandle(input_path, staging_dir),
        config,
        FakeProgressSink(),
        token,
        FakeLogger(),
    )


def _convert_invoice_fixture(input_path: Path, source_format: str) -> tuple[Any, str]:
    from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import (
        InvoiceCnConverter,
    )

    with tempfile.TemporaryDirectory() as staging:
        context = _build_fake_context(
            str(input_path),
            staging,
            "md",
            action_name="invoice_cn",
            source_format=source_format,
        )
        result = InvoiceCnConverter().convert(context)
        assert result.success is True, f"unexpected error: {result.error}"
        markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        return result, markdown


def _execute_invoice_runtime_fixture(
    *,
    input_path: Path,
    source_format: str,
    output_dir: Path,
    workspace_root: Path,
    request_id: str,
    options: dict[str, Any] | None = None,
) -> Any:
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy
    from docwen_plugin_optimizer_invoice_cn.plugin import InvoicePlugin
    from docwen_runtime.engine.route_resolver import RouteResolver
    from docwen_runtime.engine.task_manager import TaskManager
    from docwen_runtime.output.finalizer import OutputFinalizer
    from docwen_runtime.plugin_registry.registry import PluginRegistry
    from docwen_runtime.workspace.manager import WorkspaceManager

    registry = PluginRegistry()
    registry.register(InvoicePlugin())
    task_mgr = TaskManager(
        registry,
        RouteResolver(registry),
        WorkspaceManager(root_dir=str(workspace_root)),
        OutputFinalizer(),
    )
    request = ConversionRequest(
        request_id=request_id,
        input_refs=[
            FileRef(
                path=str(input_path),
                format=source_format,
                category="layout",
                size_bytes=input_path.stat().st_size,
            )
        ],
        target_format="md",
        action_name="invoice_cn",
        options=options or {},
        output_policy=OutputPolicy(output_dir=str(output_dir)),
    )
    return task_mgr.execute_single(request)


__all__ = (
    "Any",
    "Path",
    "_build_fake_context",
    "_convert_invoice_fixture",
    "_execute_invoice_runtime_fixture",
    "_load_invoice_cn_old_system_fixture",
    "_ocr_outcome",
    "_parse_invoice_markdown",
    "os",
    "pytest",
    "pytestmark",
    "tempfile",
)
