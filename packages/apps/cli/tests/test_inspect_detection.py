"""Tests for inspect command using content-based format detection.

Verifies that ``docwen inspect`` uses ``docwen_core.detection`` to report
actual format (content-based) vs. extension-based format, and correctly
flags mismatches.

Exercises the canonical ``inspect_file`` result through the CLI user path.
"""

from __future__ import annotations

import json
import os
import tempfile
import zipfile
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from docwen_cli.commands.inspect import _discover_supported_actions, execute_inspect
from docwen_core.detection import OOXML_SIGNATURE_VALIDATION_UNAVAILABLE

pytestmark = pytest.mark.contract

# ── Helpers ───────────────────────────────────────────────────────────


def _make_args(file_path: str, *, json_mode: bool = False) -> MagicMock:
    """Create a mock argparse.Namespace for inspect command."""
    args = MagicMock()
    args.file = file_path
    args.json = json_mode
    return args


class _CapabilityController:
    def __init__(self, sources: list[dict[str, object]]) -> None:
        routes: list[dict[str, object]] = []
        for source in sources:
            source_routes = source["routes"]
            assert isinstance(source_routes, list)
            for route in source_routes:
                assert isinstance(route, dict)
                routes.append(route)
        self._projection = {
            "resource": "formats",
            "contract": {"id": "docwen.runtime-capabilities", "version": 1},
            "runtime": {"state": "available", "platform": "windows"},
            "security": {"dependency_egress_guard": {}},
            "gates": [],
            "sources": sources,
            "counts": {
                "sources": len(sources),
                "routes": len(routes),
                "available_routes": sum(bool(route["available"]) for route in routes),
                "unavailable_routes": sum(not bool(route["available"]) for route in routes),
                "actions": sum(route["operation"] == "action" for route in routes),
            },
        }

    def describe_runtime_capabilities(self) -> dict[str, object]:
        return self._projection


def _route(
    route_id: str,
    *,
    source: str,
    target: str,
    action: str | None = None,
    available: bool = True,
) -> dict[str, object]:
    return {
        "id": route_id,
        "operation": "action" if action else "conversion",
        "source": source,
        "target": target,
        "action": action,
        "available": available,
        "state": "available" if available else "unavailable",
        "options": [],
    }


def _source(source_id: str, category: str, routes: list[dict[str, object]]) -> dict[str, object]:
    return {
        "id": source_id,
        "category": category,
        "available": any(bool(route["available"]) for route in routes),
        "routes": routes,
    }


def _write_signed_docx(path: Path, *, signed: bool) -> None:
    entries: dict[str, str | bytes] = {
        "word/document.xml": "<document/>",
        "_rels/.rels": (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + (
                '<Relationship Id="origin" '
                'Type="http://schemas.openxmlformats.org/package/2006/relationships/'
                'digital-signature/origin" Target="_xmlsignatures/origin.sigs"/>'
                if signed
                else ""
            )
            + "</Relationships>"
        ),
    }
    content_types = [
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    ]
    if signed:
        content_types.extend(
            [
                '<Default Extension="sigs" '
                'ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>',
                '<Override PartName="/_xmlsignatures/sig1.xml" '
                'ContentType="application/vnd.openxmlformats-package.'
                'digital-signature-xmlsignature+xml"/>',
            ]
        )
        entries.update(
            {
                "_xmlsignatures/origin.sigs": b"",
                "_xmlsignatures/sig1.xml": "<Signature/>",
                "_xmlsignatures/_rels/origin.sigs.rels": (
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="sig1" '
                    'Type="http://schemas.openxmlformats.org/package/2006/relationships/'
                    'digital-signature/signature" Target="sig1.xml"/>'
                    "</Relationships>"
                ),
            }
        )
    entries["[Content_Types].xml"] = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + "".join(content_types)
        + "</Types>"
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as package:
        for name, payload in entries.items():
            package.writestr(name, payload)


# ── Tests ─────────────────────────────────────────────────────────────


class TestInspectContentBasedDetection:
    """Inspect uses content-based detection for the detected format."""

    def test_inspect_pdf_actual_format(self, capsys):
        """PDF content is detected regardless of extension."""
        fd, path = tempfile.mkstemp(suffix=".dat")
        os.write(fd, b"%PDF-1.4\n%\x80\x80\x80\x80")
        os.close(fd)
        try:
            args = _make_args(path)
            result = execute_inspect(args)
            assert result == 0
            captured = capsys.readouterr().out
            # Extension-based format is "dat", but actual format is "pdf"
            assert "实际格式：pdf" in captured
            assert "实际类别：layout" in captured
            assert "格式关系：unrecognized_extension" in captured
            assert "准入决定：require_explicit_acceptance" in captured
        finally:
            os.unlink(path)

    def test_inspect_png_matching(self, capsys):
        """PNG with correct extension — content matches extension."""
        fd, path = tempfile.mkstemp(suffix=".png")
        os.write(fd, b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR")
        os.close(fd)
        try:
            args = _make_args(path)
            result = execute_inspect(args)
            assert result == 0
            captured = capsys.readouterr().out
            assert "实际格式：png" in captured
            # No mismatch warning when content matches extension
            assert "不匹配" not in captured
        finally:
            os.unlink(path)

    def test_inspect_jpeg_mismatch(self, capsys):
        """JPEG content with .txt extension — mismatch detected."""
        fd, path = tempfile.mkstemp(suffix=".txt")
        os.write(fd, b"\xff\xd8\xff\xe0\x00\x10JFIF\x00")
        os.close(fd)
        try:
            args = _make_args(path)
            result = execute_inspect(args)
            assert result == 0
            captured = capsys.readouterr().out
            assert "实际格式：jpeg" in captured
            assert "格式关系：cross_family_mismatch" in captured
            assert "准入决定：require_explicit_acceptance" in captured
        finally:
            os.unlink(path)

    def test_inspect_markdown_actual_format(self, capsys):
        """Markdown content detected via text format sniffing."""
        fd, path = tempfile.mkstemp(suffix=".txt")
        os.write(fd, b"# Markdown Heading\n\nContent with **bold**.\n")
        os.close(fd)
        try:
            args = _make_args(path)
            result = execute_inspect(args)
            assert result == 0
            captured = capsys.readouterr().out
            assert "实际格式：markdown" in captured
        finally:
            os.unlink(path)


class TestInspectJsonOutput:
    """Inspect JSON output includes content-based detection fields."""

    def test_json_includes_canonical_detected_format(self, capsys):
        """JSON output uses canonical inspection names without aliases."""
        fd, path = tempfile.mkstemp(suffix=".pdf")
        os.write(fd, b"%PDF-1.4\n")
        os.close(fd)
        try:
            args = _make_args(path, json_mode=True)
            result = execute_inspect(args)
            assert result == 0
            captured = capsys.readouterr().out
            data = json.loads(captured)
            inspect_data = data.get("data", data)
            assert inspect_data["detected_format"] == "pdf"
            assert inspect_data["relation"] == "exact_match"
            assert "actual_format" not in inspect_data
            assert "extension_matches_content" not in inspect_data
            assert "warning_message" in inspect_data
        finally:
            os.unlink(path)

    def test_json_mismatch_warning(self, capsys):
        """JSON output with mismatch includes warning_message."""
        fd, path = tempfile.mkstemp(suffix=".txt")
        os.write(fd, b"%PDF-1.4\n")
        os.close(fd)
        try:
            args = _make_args(path, json_mode=True)
            result = execute_inspect(args)
            assert result == 0
            captured = capsys.readouterr().out
            data = json.loads(captured)
            inspect_data = data.get("data", data)
            assert inspect_data["relation"] == "cross_family_mismatch"
            assert inspect_data["warning_message"] is not None
            assert "pdf" in inspect_data["warning_message"].lower()
        finally:
            os.unlink(path)

    def test_json_missing_file(self, capsys):
        """JSON output for nonexistent file."""
        args = _make_args("/nonexistent/inspect_test_12345.xyz", json_mode=True)
        result = execute_inspect(args)
        assert result == 2  # ExitCode.INVALID_INPUT
        captured = capsys.readouterr().out
        data = json.loads(captured)
        assert "error" in data


class TestInspectOoxmlSignaturePresence:
    """Inspect projects the shared presence-only signature fact."""

    def test_json_exposes_typed_warning_and_structural_state(
        self,
        tmp_path: Path,
        capsys,
    ) -> None:
        source = tmp_path / "signed.docx"
        _write_signed_docx(source, signed=True)

        result = execute_inspect(_make_args(str(source), json_mode=True))

        assert result == 0
        payload = json.loads(capsys.readouterr().out)
        data = payload.get("data", payload)
        assert data["ooxml_signature"]["state"] == "complete"
        assert [warning["code"] for warning in data["warnings"]] == [OOXML_SIGNATURE_VALIDATION_UNAVAILABLE]
        assert "intact and tampered inputs cannot be distinguished" in data["warnings"][0]["message"]

    def test_text_exposes_code_and_limitation(
        self,
        tmp_path: Path,
        capsys,
    ) -> None:
        source = tmp_path / "signed.docx"
        _write_signed_docx(source, signed=True)

        result = execute_inspect(_make_args(str(source)))

        assert result == 0
        output = capsys.readouterr().out
        assert OOXML_SIGNATURE_VALIDATION_UNAVAILABLE in output
        assert "intact and tampered inputs cannot be distinguished" in output

    def test_unsigned_control_has_no_signature_warning(
        self,
        tmp_path: Path,
        capsys,
    ) -> None:
        source = tmp_path / "unsigned.docx"
        _write_signed_docx(source, signed=False)

        result = execute_inspect(_make_args(str(source), json_mode=True))

        assert result == 0
        payload = json.loads(capsys.readouterr().out)
        data = payload.get("data", payload)
        assert data["ooxml_signature"]["state"] == "unsigned"
        assert data["warnings"] == []


class TestInspectErrorPaths:
    """Error handling in inspect command."""

    def test_inspect_nonexistent_file(self, capsys):
        """Inspect nonexistent file returns error exit code."""
        args = _make_args("/nonexistent/inspect_test_99999.xyz")
        result = execute_inspect(args)
        assert result == 2  # ExitCode.INVALID_INPUT
        captured_err = capsys.readouterr().err
        assert "文件不存在" in captured_err or "error" in captured_err.lower()

    def test_inspect_empty_file_arg(self, capsys):
        """Inspect with empty file argument."""
        args = _make_args("")
        result = execute_inspect(args)
        assert result == 2


class TestInspectInfoIntegration:
    """Verify :func:`inspect_file` integration in the inspect user path.

    These tests confirm that the structured inspection is the single source
    of truth, rather than an ad-hoc CLI projection.
    """

    def test_info_pdf_matching_extension(self, capsys):
        """PDF with .pdf extension → info reports match, supported, valid."""
        fd, path = tempfile.mkstemp(suffix=".pdf")
        os.write(fd, b"%PDF-1.4\n%\x80\x80\x80\x80")
        os.close(fd)
        try:
            args = _make_args(path)
            result = execute_inspect(args)
            assert result == 0
            captured = capsys.readouterr().out
            assert "扩展名格式：pdf" in captured
            assert "实际格式：pdf" in captured
            assert "实际类别：layout" in captured
            # No mismatch warning when content matches extension
            assert "不匹配" not in captured
        finally:
            os.unlink(path)

    def test_info_markdown_content_in_txt(self, capsys):
        """Markdown content in .txt → info shows mismatch with actual markdown category."""
        fd, path = tempfile.mkstemp(suffix=".txt")
        os.write(fd, b"# Heading\n\n**bold** and *italic*.\n")
        os.close(fd)
        try:
            args = _make_args(path)
            result = execute_inspect(args)
            assert result == 0
            captured = capsys.readouterr().out
            # Extension-based format is "txt", but actual is "markdown"
            assert "扩展名格式：txt" in captured
            assert "实际格式：markdown" in captured
            assert "格式关系：compatible_text" in captured
            assert "准入决定：allow_with_warning" in captured
        finally:
            os.unlink(path)

    def test_info_docx_zip_content(self, capsys):
        """DOCX file detected via ZIP inspection — info path."""
        from docx import Document

        fd, path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        try:
            Document().save(path)
            args = _make_args(path)
            result = execute_inspect(args)
            assert result == 0
            captured = capsys.readouterr().out
            assert "扩展名格式：docx" in captured
            assert "实际格式：docx" in captured
            assert "实际类别：document" in captured
        finally:
            os.unlink(path)

    def test_info_csv_content_spreadsheet_category(self, capsys):
        """CSV content → spreadsheet category via inspect_file."""
        fd, path = tempfile.mkstemp(suffix=".csv")
        os.write(fd, b"a,b,c\n1,2,3\n4,5,6\n")
        os.close(fd)
        try:
            args = _make_args(path)
            result = execute_inspect(args)
            assert result == 0
            captured = capsys.readouterr().out
            assert "实际格式：csv" in captured
            assert "实际类别：spreadsheet" in captured
            # No mismatch warning
            assert "不匹配" not in captured
        finally:
            os.unlink(path)

    @pytest.mark.parametrize(
        "payload",
        [
            "city;value\n北京;1\n上海;2\n".encode(),
            "city,value\n北京,1\n上海,2\n".encode("utf-16"),
        ],
    )
    def test_info_json_admits_semicolon_and_utf16_csv(self, capsys, payload: bytes):
        fd, path = tempfile.mkstemp(suffix=".csv")
        os.write(fd, payload)
        os.close(fd)
        try:
            result = execute_inspect(_make_args(path, json_mode=True))
            assert result == 0
            data = json.loads(capsys.readouterr().out)["data"]
            assert data["detected_format"] == "csv"
            assert data["detected_category"] == "spreadsheet"
            assert data["relation"] == "exact_match"
            assert data["decision"] == "allow"
        finally:
            os.unlink(path)

    def test_info_json_output_uses_file_inspection_keys(self, capsys):
        """JSON output mirrors the canonical FileInspection names."""
        fd, path = tempfile.mkstemp(suffix=".pdf")
        os.write(fd, b"%PDF-1.4\n%\x80\x80\x80\x80")
        os.close(fd)
        try:
            args = _make_args(path, json_mode=True)
            result = execute_inspect(args)
            assert result == 0
            captured = capsys.readouterr().out
            data = json.loads(captured)
            inspect_data = data.get("data", data)
            # Canonical inspection fields are present without compatibility aliases.
            for key in (
                "detected_format",
                "detected_category",
                "declared_format",
                "declared_category",
                "extension",
                "relation",
                "warning_message",
            ):
                assert key in inspect_data, f"Missing key: {key}"
            assert inspect_data["relation"] == "exact_match"
            assert "actual_format" not in inspect_data
        finally:
            os.unlink(path)

    def test_info_unsupported_extension_fallback(self, capsys):
        """File with unrecognized extension → info still reports actual content."""
        fd, path = tempfile.mkstemp(suffix=".oddball")
        os.write(fd, b"%PDF-1.4\n")
        os.close(fd)
        try:
            args = _make_args(path)
            result = execute_inspect(args)
            assert result == 0
            captured = capsys.readouterr().out
            assert "实际格式：pdf" in captured
            assert "格式关系：unrecognized_extension" in captured
            assert "准入决定：require_explicit_acceptance" in captured
        finally:
            os.unlink(path)

    def test_info_html_content_markup_category(self, capsys):
        """HTML content → markup category via inspect_file."""
        fd, path = tempfile.mkstemp(suffix=".html")
        os.write(fd, b"<!DOCTYPE html>\n<html><body>Test</body></html>\n")
        os.close(fd)
        try:
            args = _make_args(path)
            result = execute_inspect(args)
            assert result == 0
            captured = capsys.readouterr().out
            assert "实际格式：html" in captured
            assert "实际类别：markup" in captured
        finally:
            os.unlink(path)


class TestInspectSupportedActionsContract:
    """Inspect joins Runtime routes to public commands without format tables."""

    def test_exact_routes_then_category_fallback_are_projected(self) -> None:
        controller = _CapabilityController(
            [
                _source("xlsx", "spreadsheet", [_route("xlsx-md", source="xlsx", target="md")]),
                _source(
                    "spreadsheet",
                    "spreadsheet",
                    [_route("merge", source="spreadsheet", target="xlsx", action="merge_tables")],
                ),
            ]
        )

        actions, discovery = _discover_supported_actions(controller, "xlsx", "spreadsheet")

        assert actions == ["inspect", "convert", "merge tables"]
        assert discovery == {
            "state": "available",
            "matched_by": "detected_format_then_workflow_category",
            "source_ids": ["xlsx", "spreadsheet"],
            "error": None,
        }

    def test_unavailable_and_internal_actions_are_not_advertised(self) -> None:
        controller = _CapabilityController(
            [
                _source(
                    "pdf",
                    "layout",
                    [
                        _route("split", source="pdf", target="pdf", action="split_pdf", available=False),
                        _route("private", source="pdf", target="md", action="new_manifest_action"),
                    ],
                )
            ]
        )

        actions, discovery = _discover_supported_actions(controller, "pdf", "layout")

        assert actions == ["inspect"]
        assert discovery["state"] == "available"
        assert "split_pdf" not in actions
        assert "new_manifest_action" not in actions

    def test_successful_empty_discovery_is_not_failure(self) -> None:
        actions, discovery = _discover_supported_actions(_CapabilityController([]), "pdf", "layout")

        assert actions == ["inspect"]
        assert discovery == {
            "state": "available",
            "matched_by": "none",
            "source_ids": [],
            "error": None,
        }

    def test_execute_marks_discovery_failure_without_falling_back_to_a_guess_table(
        self,
        tmp_path: Path,
        capsys,
    ) -> None:
        source = tmp_path / "sample.pdf"
        source.write_bytes(b"%PDF-1.4\n")

        assert execute_inspect(_make_args(str(source), json_mode=True), controller=None) == 0

        payload = json.loads(capsys.readouterr().out)
        assert payload["data"]["supported_actions"] == ["inspect"]
        assert payload["data"]["supported_actions_discovery"]["state"] == "unavailable"
        assert payload["data"]["supported_actions_discovery"]["error"]["code"] == "capability_unavailable"
