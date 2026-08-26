"""Print plugin route resolution and unavailable-capability tests.

Verifies that:
- docx/xlsx → pdf routes resolve correctly at the plugin level
- document → OFD stays unavailable and outside the executable route table
- can_handle() and convert() dispatch agree with manifest
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.unit


# ── Fake context builder ──────────────────────────────────────────────


def _build_fake_context(
    source_format: str,
    target_format: str,
    tmp_path: Path,
    *,
    action_name: str = "",
    input_bytes: bytes = b"dummy",
    config_values: dict[str, Any] | None = None,
) -> Any:
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    file_refs = [
        FileRef(
            path="/fake/input." + source_format,
            format=source_format,
            category="document" if source_format in ("docx", "doc", "odt", "rtf", "wps", "document") else "spreadsheet",
        )
    ]
    request = ConversionRequest(
        request_id="test-print-001",
        input_refs=file_refs,
        target_format=target_format,
        action_name=action_name,
        output_policy=OutputPolicy(),
    )
    config = FakeConfigView(config_values)
    token = CancellationToken()
    staging = tmp_path / "staging"
    staging.mkdir()
    return FakeExecutionContext(
        request,
        FakeWorkspaceHandle("/fake/input." + source_format, str(staging)),
        config,
        FakeProgressSink(),
        token,
        FakePluginLogger(),
    )


# ── Route resolution tests ───────────────────────────────────────────


class TestPrintCanHandle:
    """Verify can_handle() matches manifest routes exactly."""

    def test_document_family_to_pdf_accepted(self) -> None:
        from docwen_plugin_print import PrintPlugin

        plugin = PrintPlugin()
        for src in ("docx", "doc", "odt", "rtf", "wps", "document"):
            assert plugin.can_handle(src, "pdf") is True

    def test_document_family_to_ofd_is_not_executable(self) -> None:
        from docwen_plugin_print import PrintPlugin

        plugin = PrintPlugin()
        for src in ("docx", "doc", "odt", "rtf", "wps", "document"):
            assert plugin.can_handle(src, "ofd") is False

    def test_spreadsheet_family_to_pdf_accepted(self) -> None:
        from docwen_plugin_print import PrintPlugin

        plugin = PrintPlugin()
        for src in ("xlsx", "xls", "ods", "et", "csv", "spreadsheet"):
            assert plugin.can_handle(src, "pdf") is True

    def test_unregistered_routes_rejected(self) -> None:
        from docwen_plugin_print import PrintPlugin

        plugin = PrintPlugin()
        assert plugin.can_handle("tsv", "pdf") is False
        assert plugin.can_handle("pdf", "md") is False
        assert plugin.can_handle("docx", "md") is False


class TestPrintPdfBridge:
    """Verify implemented PDF routes call the Office bridge with route-specific settings."""

    def test_document_to_pdf_uses_default_wps_bridge(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_print import PrintPlugin
        from docwen_plugin_print.paged_output import converter as print_converter

        calls: list[dict[str, Any]] = []

        def fake_convert_with_backend_priority(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            Path(output_path).write_bytes(b"%PDF-1.4\nfake document pdf")
            calls.append({"input_path": input_path, "output_path": output_path, **kwargs})
            return BridgeResult(True, output_path=output_path, backend="fake-word", message="ok")

        monkeypatch.setattr(print_converter, "convert_with_backend_priority", fake_convert_with_backend_priority)

        result = PrintPlugin().convert(_build_fake_context("docx", "pdf", tmp_path))

        assert result.success is True
        assert result.artifacts[0].suggested_name == "input.pdf"
        assert result.metrics.extra["engine"] == "office_bridge"
        assert result.metrics.extra["backend"] == "fake-word"
        assert calls[0]["source_format"] == "docx"
        assert calls[0]["backend_priority"] == ["wps_writer", "msoffice_word", "libreoffice"]
        assert calls[0]["libreoffice_format"] == "pdf:writer_pdf_Export"
        assert {key: candidate.prog_id for key, candidate in calls[0]["com_candidates"].items()} == {
            "wps_writer": "KWPS.Application",
            "msoffice_word": "Word.Application",
        }

    def test_spreadsheet_to_pdf_uses_calc_bridge(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_print import PrintPlugin
        from docwen_plugin_print.paged_output import converter as print_converter

        calls: list[dict[str, Any]] = []

        def fake_convert_with_backend_priority(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            Path(output_path).write_bytes(b"%PDF-1.4\nfake spreadsheet pdf")
            calls.append({"input_path": input_path, "output_path": output_path, **kwargs})
            return BridgeResult(True, output_path=output_path, backend="fake-calc", message="ok")

        monkeypatch.setattr(print_converter, "convert_with_backend_priority", fake_convert_with_backend_priority)

        result = PrintPlugin().convert(
            _build_fake_context(
                "xlsx",
                "pdf",
                tmp_path,
                config_values={
                    "software": {
                        "special_conversions": {
                            "spreadsheet_to_pdf": [
                                "wps_spreadsheets",
                                "msoffice_excel",
                                "libreoffice",
                            ]
                        }
                    }
                },
            )
        )

        assert result.success is True
        assert result.artifacts[0].suggested_name == "input.pdf"
        assert result.metrics.extra["engine"] == "office_bridge"
        assert result.metrics.extra["backend"] == "fake-calc"
        assert calls[0]["source_format"] == "xlsx"
        assert calls[0]["backend_priority"] == ["wps_spreadsheets", "msoffice_excel", "libreoffice"]
        assert calls[0]["libreoffice_format"] == "pdf:calc_pdf_Export"
        assert {key: candidate.prog_id for key, candidate in calls[0]["com_candidates"].items()} == {
            "wps_spreadsheets": "KET.Application",
            "msoffice_excel": "Excel.Application",
        }

    def test_spreadsheet_to_pdf_honors_libreoffice_first(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_print import PrintPlugin
        from docwen_plugin_print.paged_output import converter as print_converter

        calls: list[dict[str, Any]] = []

        def fake_convert_with_backend_priority(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            Path(output_path).write_bytes(b"%PDF-1.4\nfake libreoffice pdf")
            calls.append({"input_path": input_path, "output_path": output_path, **kwargs})
            return BridgeResult(True, output_path=output_path, backend="LibreOffice", message="ok")

        monkeypatch.setattr(print_converter, "convert_with_backend_priority", fake_convert_with_backend_priority)
        context = _build_fake_context(
            "xlsx",
            "pdf",
            tmp_path,
            config_values={
                "software": {
                    "special_conversions": {
                        "spreadsheet_to_pdf": [
                            "libreoffice",
                            "msoffice_excel",
                            "wps_spreadsheets",
                        ]
                    }
                }
            },
        )

        result = PrintPlugin().convert(context)

        assert result.success is True
        assert len(calls) == 1
        assert calls[0]["backend_priority"] == ["libreoffice", "msoffice_excel", "wps_spreadsheets"]
        assert calls[0]["libreoffice_format"] == "pdf:calc_pdf_Export"

    def test_document_to_pdf_honors_microsoft_word_first(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_print import PrintPlugin
        from docwen_plugin_print.paged_output import converter as print_converter

        calls: list[dict[str, Any]] = []

        def fake_convert_with_backend_priority(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            Path(output_path).write_bytes(b"%PDF-1.4\nfake word pdf")
            calls.append({"input_path": input_path, "output_path": output_path, **kwargs})
            return BridgeResult(True, output_path=output_path, backend="Microsoft Word", message="ok")

        monkeypatch.setattr(print_converter, "convert_with_backend_priority", fake_convert_with_backend_priority)
        context = _build_fake_context(
            "docx",
            "pdf",
            tmp_path,
            config_values={
                "software": {
                    "special_conversions": {
                        "document_to_pdf": [
                            "msoffice_word",
                            "wps_writer",
                            "libreoffice",
                        ]
                    }
                }
            },
        )

        result = PrintPlugin().convert(context)

        assert result.success is True
        assert len(calls) == 1
        assert calls[0]["backend_priority"] == ["msoffice_word", "wps_writer", "libreoffice"]
        assert calls[0]["com_candidates"]["msoffice_word"].prog_id == "Word.Application"
        assert calls[0]["libreoffice_format"] == "pdf:writer_pdf_Export"

    def test_pdf_bridge_failure_is_structured(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_print import PrintPlugin
        from docwen_plugin_print.paged_output import converter as print_converter

        def fake_convert_with_backend_priority(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            return BridgeResult(False, backend="fake-word", message="Install LibreOffice.")

        monkeypatch.setattr(print_converter, "convert_with_backend_priority", fake_convert_with_backend_priority)

        result = PrintPlugin().convert(_build_fake_context("docx", "pdf", tmp_path))

        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "conversion_failed"
        assert result.error.diagnostic_code == "OFFICE2PDF-BRIDGE-FAILED"
        assert "Install LibreOffice." in result.error.message


class TestPrintUnavailableRoutes:
    """Unavailable capabilities never enter the executable route table."""

    def test_docx_to_ofd_direct_call_returns_unsupported_route(self, tmp_path: Path) -> None:
        from docwen_plugin_print import PrintPlugin

        plugin = PrintPlugin()
        context = _build_fake_context("docx", "ofd", tmp_path)
        result = plugin.convert(context)

        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "unsupported_route"
        assert result.error.diagnostic_code == "PRINT-UNSUPPORTED-ROUTE"
        assert "docx→ofd" in result.error.message

    def test_manifest_exposes_ofd_only_as_unavailable_capability(self) -> None:
        from docwen_plugin_print import PrintPlugin

        plugin = PrintPlugin()
        unavailable = plugin.manifest.extra["unavailable_routes"]
        assert [(item.source, item.targets) for item in unavailable] == [("document", ["ofd"])]
        assert all(route.target_format != "ofd" for route in plugin.manifest.routes)

    def test_unknown_route_returns_unsupported_route(self, tmp_path: Path) -> None:
        from docwen_plugin_print import PrintPlugin

        plugin = PrintPlugin()
        context = _build_fake_context("pdf", "md", tmp_path)
        result = plugin.convert(context)

        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "unsupported_route"
