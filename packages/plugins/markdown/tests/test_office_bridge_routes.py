"""Tests for Markdown routes backed by the Office bridge."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_core.office_bridge import BridgeResult
from docwen_plugin_markdown import MarkdownPlugin

from .conftest import make_context, write_temp_md

pytestmark = pytest.mark.contract


class TestOfficeBridgeRoutes:
    """Office-backed Markdown routes use generated intermediate files."""

    @pytest.mark.parametrize(
        ("target_format", "config_values", "expected_priority", "expected_candidates"),
        [
            (
                "doc",
                {"software": {"default_priority": {"word_processors": ["msoffice_word", "libreoffice", "wps_writer"]}}},
                ["msoffice_word", "libreoffice", "wps_writer"],
                {"wps_writer", "msoffice_word"},
            ),
            (
                "odt",
                {"software": {"special_conversions": {"odt": ["libreoffice", "msoffice_word"]}}},
                ["libreoffice", "msoffice_word"],
                {"msoffice_word"},
            ),
            (
                "pdf",
                {
                    "software": {
                        "special_conversions": {"document_to_pdf": ["msoffice_word", "wps_writer", "libreoffice"]}
                    }
                },
                ["msoffice_word", "wps_writer", "libreoffice"],
                {"wps_writer", "msoffice_word"},
            ),
            (
                "xls",
                {
                    "software": {
                        "default_priority": {
                            "spreadsheet_processors": ["msoffice_excel", "libreoffice", "wps_spreadsheets"]
                        }
                    }
                },
                ["msoffice_excel", "libreoffice", "wps_spreadsheets"],
                {"wps_spreadsheets", "msoffice_excel"},
            ),
            (
                "ods",
                {"software": {"special_conversions": {"ods": ["libreoffice", "msoffice_excel"]}}},
                ["libreoffice", "msoffice_excel"],
                {"msoffice_excel"},
            ),
        ],
    )
    def test_markdown_office_bridge_consumes_configured_backend_priority(
        self,
        monkeypatch: pytest.MonkeyPatch,
        target_format: str,
        config_values: dict[str, object],
        expected_priority: list[str],
        expected_candidates: set[str],
    ) -> None:
        """Every Markdown Office route must consume its authoritative software key."""
        calls: list[tuple[str, list[str], set[str]]] = []

        def fake_convert_with_backend_priority(
            input_path,
            output_path,
            *,
            source_format,
            backend_priority,
            com_candidates,
            libreoffice_format,
            **kwargs,
        ):
            del input_path, libreoffice_format, kwargs
            calls.append((source_format, list(backend_priority), set(com_candidates)))
            Path(output_path).write_bytes(b"configured office output")
            return BridgeResult(success=True, output_path=str(output_path), backend="configured-office")

        monkeypatch.setattr(
            "docwen_plugin_markdown.office_bridge.converter.convert_with_backend_priority",
            fake_convert_with_backend_priority,
        )
        md_path = write_temp_md("# Configured priority\n\n| A |\n|---|\n| 1 |\n")
        ctx, _workspace = make_context(
            md_path,
            target_format=target_format,
            config_values=config_values,
        )

        result = MarkdownPlugin().convert(ctx)

        assert result.success
        expected_source_format = "xlsx" if target_format in {"xls", "ods"} else "docx"
        assert calls == [(expected_source_format, expected_priority, expected_candidates)]

    @pytest.mark.parametrize(
        ("target_format", "intermediate_suffix", "bridge_suffix", "libreoffice_format"),
        [
            ("doc", ".docx", ".doc", "doc"),
            ("odt", ".docx", ".odt", "odt"),
            ("rtf", ".docx", ".rtf", "rtf"),
            ("wps", ".docx", ".doc", "doc"),
            ("pdf", ".docx", ".pdf", "pdf"),
        ],
    )
    def test_md_to_document_uses_office_bridge(
        self,
        monkeypatch,
        target_format: str,
        intermediate_suffix: str,
        bridge_suffix: str,
        libreoffice_format: str,
    ):
        """MD->DOC/ODT/RTF/WPS/PDF renders DOCX first, then uses the Office bridge."""
        calls: list[tuple[Path, Path, str, str]] = []

        def fake_convert_with_fallback(input_path, output_path, *, libreoffice_format, **kwargs):
            input_path = Path(input_path)
            output_path = Path(output_path)
            calls.append((input_path, output_path, libreoffice_format, kwargs["source_format"]))
            output_path.write_bytes(b"office document")
            return BridgeResult(success=True, output_path=str(output_path), backend="fake-office")

        monkeypatch.setattr(
            "docwen_plugin_markdown.office_bridge.converter.convert_with_backend_priority",
            fake_convert_with_fallback,
        )
        md_path = write_temp_md("# Title\n\nText.")
        ctx, _workspace = make_context(md_path, target_format=target_format)
        plugin = MarkdownPlugin()

        result = plugin.convert(ctx)

        assert result.success
        assert result.error is None
        assert len(result.artifacts) == 1
        artifact = result.artifacts[0]
        assert artifact.suggested_name.endswith(f".{target_format}")
        assert Path(artifact.staging_path).suffix == f".{target_format}"
        assert Path(artifact.staging_path).exists()
        assert calls
        assert calls[0][0].suffix == intermediate_suffix
        assert calls[0][1].suffix == bridge_suffix
        assert calls[0][2] == libreoffice_format
        assert calls[0][3] == "docx"
        assert result.metrics.extra["engine"] == "office_bridge"
        assert result.metrics.extra["backend"] == "fake-office"

    @pytest.mark.parametrize(
        ("target_format", "intermediate_suffix"),
        [("xls", ".xlsx"), ("ods", ".xlsx")],
    )
    def test_md_to_spreadsheet_uses_office_bridge(self, monkeypatch, target_format: str, intermediate_suffix: str):
        """MD->XLS/ODS renders XLSX first, then uses the Office bridge."""
        calls: list[tuple[Path, Path, str, str]] = []

        def fake_convert_with_fallback(input_path, output_path, *, libreoffice_format, **kwargs):
            input_path = Path(input_path)
            output_path = Path(output_path)
            calls.append((input_path, output_path, libreoffice_format, kwargs["source_format"]))
            output_path.write_bytes(b"office spreadsheet")
            return BridgeResult(success=True, output_path=str(output_path), backend="fake-office")

        monkeypatch.setattr(
            "docwen_plugin_markdown.office_bridge.converter.convert_with_backend_priority",
            fake_convert_with_fallback,
        )
        md_path = write_temp_md("# Title\n\n| A |\n|---|\n| 1 |\n")
        ctx, _workspace = make_context(md_path, target_format=target_format)
        plugin = MarkdownPlugin()

        result = plugin.convert(ctx)

        assert result.success
        assert result.error is None
        assert len(result.artifacts) == 1
        artifact = result.artifacts[0]
        assert artifact.suggested_name.endswith(f".{target_format}")
        assert Path(artifact.staging_path).suffix == f".{target_format}"
        assert Path(artifact.staging_path).exists()
        assert calls
        assert calls[0][0].suffix == intermediate_suffix
        assert calls[0][1].suffix == f".{target_format}"
        assert calls[0][2] == target_format
        assert calls[0][3] == "xlsx"
        assert result.metrics.extra["engine"] == "office_bridge"
        assert result.metrics.extra["backend"] == "fake-office"

    def test_office_bridge_failure_is_diagnostic(self, monkeypatch):
        """Bridge failures return structured conversion errors."""

        def fake_convert_with_fallback(input_path, output_path, *, libreoffice_format, **kwargs):
            return BridgeResult(
                success=False,
                output_path=None,
                backend="fake-office",
                message="Office conversion failed",
            )

        monkeypatch.setattr(
            "docwen_plugin_markdown.office_bridge.converter.convert_with_backend_priority",
            fake_convert_with_fallback,
        )
        md_path = write_temp_md("# Title\n")
        ctx, _workspace = make_context(md_path, target_format="doc")
        plugin = MarkdownPlugin()

        result = plugin.convert(ctx)

        assert not result.success
        assert result.error is not None
        assert result.error.error_type == "office_bridge_failed"
        assert result.error.diagnostic_code == "MD-OFFICE-BRIDGE-FAILED"
        assert result.artifacts == []
