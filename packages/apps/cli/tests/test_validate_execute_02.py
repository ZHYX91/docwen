"""Focused tests split from test_validate_execute.py."""

from __future__ import annotations

import pytest

from ._validate_execute_support import (
    _validation_controller,
    argparse,
    patch,
)

pytestmark = pytest.mark.unit


class TestExecuteConvertValidateRouting:
    """Verify execute_convert routes validate → validate for markdown."""

    def _fake_args(self, **overrides) -> argparse.Namespace:
        ns = argparse.Namespace()
        ns.action = "validate"
        ns.json = False
        ns.quiet = False
        ns.verbose = False
        ns.timing = False
        ns.batch = False
        ns.jobs = 1
        ns.continue_on_error = False
        ns.output = None
        ns.dry_run = False
        ns.to = None
        ns.template = None
        ns.check = []
        ns.extract_img = False
        ns.no_extract_img = False
        ns.ocr = False
        ns.image_mode = None
        ns.ocr_placement = None
        ns.clean_numbering = None
        ns.add_numbering = None
        ns.heading_merge_mode = None
        ns.files = []
        ns.file = None
        ns.pages = None
        ns.dpi = None
        ns.mode = None
        ns.keep_alpha = False
        for k, v in overrides.items():
            setattr(ns, k, v)
        return ns

    def test_validate_markdown_file_routes_to_validate(self, tmp_path) -> None:
        """The public validate command routes Markdown to the validate action."""

        from docwen_cli.commands.convert import execute_convert

        md_file = tmp_path / "test.md"
        md_file.write_text("# Test\n", encoding="utf-8")

        mock_controller = _validation_controller()

        args = self._fake_args(
            action="validate",
            files=[str(md_file)],
        )

        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(md_file)], [], []),
        ):
            exit_code = execute_convert(args, controller=mock_controller)

        assert exit_code == 0
        # Verify the controller was called with validate action
        call_args = mock_controller.execute_single.call_args[0][0]
        assert call_args.action_name == "validate"
        assert call_args.input_refs[0].format == "markdown"
        assert call_args.target_format == "markdown"
        assert "locale" not in call_args.options
        assert "yaml_key_labels" not in call_args.options

    def test_validate_markdown_file_carries_only_proofread_check_options(self, tmp_path) -> None:
        """Proofread markdown action should not receive Markdown export metadata options."""

        from docwen_cli.commands.convert import execute_convert
        from docwen_cli.i18n import init_cli_locale

        md_file = tmp_path / "test.md"
        md_file.write_text("# Test\n", encoding="utf-8")

        mock_controller = _validation_controller()

        try:
            init_cli_locale("de_DE")
            args = self._fake_args(
                action="validate",
                files=[str(md_file)],
                check=["typo", "sensitive"],
            )

            with patch(
                "docwen_cli.commands.convert.validate_files",
                return_value=([str(md_file)], [], []),
            ):
                exit_code = execute_convert(args, controller=mock_controller)
        finally:
            init_cli_locale("zh_CN")

        assert exit_code == 0
        call_args = mock_controller.execute_single.call_args[0][0]
        assert call_args.action_name == "validate"
        assert call_args.target_format == "markdown"
        assert call_args.options == {
            "enable_symbol_pairing": False,
            "enable_symbol_correction": False,
            "enable_typos_rule": True,
            "enable_sensitive_word": True,
        }

    def test_validate_docx_file_stays_validate(self, tmp_path) -> None:
        """The public validate command keeps DOCX on the direct route."""

        from docx import Document

        from docwen_cli.commands.convert import execute_convert

        docx_file = tmp_path / "test.docx"
        Document().save(docx_file)

        mock_controller = _validation_controller()

        args = self._fake_args(
            action="validate",
            files=[str(docx_file)],
        )

        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(docx_file)], [], []),
        ):
            exit_code = execute_convert(args, controller=mock_controller)

        assert exit_code == 0
        call_args = mock_controller.execute_single.call_args[0][0]
        assert call_args.action_name == "validate"
        assert call_args.input_refs[0].format == "docx"
        assert call_args.input_refs[0].category == "document"
        assert call_args.target_format == "docx"

    def test_validate_rejects_legacy_document_category_target(self, tmp_path) -> None:
        """Protocol 3 does not expose the old category-valued validation target."""

        from docx import Document

        from docwen_cli.commands.convert import execute_convert

        docx_file = tmp_path / "report.docx"
        Document().save(docx_file)

        mock_controller = _validation_controller()

        args = self._fake_args(
            action="validate",
            files=[str(docx_file)],
            to="document",
        )

        with patch(
            "docwen_cli.commands.convert.validate_files",
            return_value=([str(docx_file)], [], []),
        ):
            exit_code = execute_convert(args, controller=mock_controller)

        assert exit_code == 2
        mock_controller.execute_single.assert_not_called()

    @pytest.mark.parametrize("bridge_available", [True, False])
    def test_validate_routes_legacy_office_content_through_document_capability(
        self, tmp_path, bridge_available
    ) -> None:
        from docwen_cli.commands.convert import execute_convert
        from docwen_core.models import (
            AdmissionDecision,
            DetectionConfidence,
            DetectionMethod,
            FileInspection,
            FormatRelation,
            StructureStatus,
        )

        source = tmp_path / "legacy.doc"
        source.write_bytes(b"legacy-office-content")
        inspection = FileInspection(
            file_path=str(source),
            size_bytes=source.stat().st_size,
            mtime_ns=source.stat().st_mtime_ns,
            extension=".doc",
            declared_format="doc",
            declared_category="document",
            detected_format="doc",
            detected_category="document",
            workflow_category="document",
            detection_method=DetectionMethod.SIGNATURE,
            confidence=DetectionConfidence.CERTAIN,
            structure_status=StructureStatus.VALID,
            relation=FormatRelation.EXACT_MATCH,
            decision=AdmissionDecision.ALLOW,
            declared_supported=True,
            detected_supported=True,
        )
        controller = _validation_controller()
        if not bridge_available:
            projection = controller.describe_runtime_capabilities.return_value
            source_projection = next(item for item in projection["sources"] if item["id"] == "doc")
            source_projection["available"] = False
            source_projection["routes"][0].update(available=False, state="unavailable")
            projection["counts"].update(available_routes=3, unavailable_routes=1)
        args = self._fake_args(action="validate", files=[str(source)], to="docx")

        with (
            patch("docwen_cli.commands.convert.validate_files", return_value=([str(source)], [], [])),
            patch("docwen_cli.commands.convert._inspection_for", return_value=inspection),
        ):
            exit_code = execute_convert(args, controller=controller)

        if not bridge_available:
            assert exit_code == 6
            controller.execute_single.assert_not_called()
            return
        assert exit_code == 0
        request = controller.execute_single.call_args.args[0]
        assert request.action_name == "validate"
        assert request.input_refs[0].format == "doc"
        assert request.input_refs[0].category == "document"
        assert request.target_format == "docx"


@pytest.mark.integration
@pytest.mark.parametrize(
    ("filename", "text", "detected", "expected_offsets"),
    [
        ("plain.md", "Ordinary Markdown paragraph.\n", "txt", [0]),
        ("inline.md", "A paragraph with `inline code`.\n", "txt", []),
        ("protected.md", "Ordinary `Ordinary`.\n", "txt", [0]),
        ("list.md", "- Ordinary Markdown paragraph.\n", "txt", [2]),
        ("heading.md", "# Ordinary\n", "markdown", [2]),
        ("plain.txt", "Ordinary Markdown paragraph.\n", "txt", [0]),
    ],
)
def test_admitted_markdown_workflow_reaches_real_report(
    tmp_path, monkeypatch, capsys, filename, text, detected, expected_offsets
):
    import hashlib
    import json

    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_cli.main import main
    from docwen_core.detection import inspect_file
    from docwen_runtime.config import ConfigLoader

    source = tmp_path / filename
    source.write_bytes(text.encode("utf-8"))
    original = source.read_bytes()
    report_path = tmp_path / "report.json"
    monkeypatch.setenv("DOCWEN_RUNTIME_ROOT", str(tmp_path / "runtime"))
    loader = ConfigLoader(
        user_dir=tmp_path / "config",
        runtime_overrides={
            "logger": {"console_enable": False},
            "proofread": {"typos": {"entries": {"Expected": ["Ordinary"]}}},
        },
    )
    inspection = inspect_file(str(source))
    assert inspection.detected_format == detected
    assert inspection.workflow_category == "markdown"
    assert inspection.may_execute and not inspection.requires_explicit_acceptance
    code = main(
        ["validate", str(source), "--report", str(report_path), "--check", "typo", "--json"],
        runtime_port_factory=lambda: create_runtime_port(config_loader=loader),
        config_port_factory=lambda: ConfigPortAdapter(loader),
    )
    payload = json.loads(capsys.readouterr().out)
    assert code == 0, payload
    assert payload["success"]
    # Inspect final published bytes independently of presenter embedding.
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report == payload["data"]["details"]["proofread"]
    assert report["schema"] == "docwen.proofread_report.v2"
    assert report["file"] == filename
    assert report["source"]["content_sha256"] == hashlib.sha256(original).hexdigest()
    assert report["location_contract"]["coordinate_system"] == "unicode_code_point"
    assert [issue["range"]["start"]["offset"] for issue in report["issues"]] == expected_offsets
    for issue, offset in zip(report["issues"], expected_offsets, strict=True):
        assert issue["matched_text"] == "Ordinary"
        assert issue["range"]["end"]["offset"] == offset + len("Ordinary")
        assert text[offset : offset + len("Ordinary")] == issue["matched_text"]
    assert source.read_bytes() == original


@pytest.mark.integration
@pytest.mark.parametrize(
    "body",
    ["<!doctype html><html><body><h1>Ordinary</h1></body></html>\n", "Name,Amount\nAlice,7\nBob,9\n"],
)
def test_wrong_family_markdown_suffix_still_requires_explicit_acceptance(tmp_path, capsys, body):
    import json

    from docwen_cli.main import main
    from docwen_core.detection import inspect_file

    source = tmp_path / "wrong.md"
    source.write_bytes(body.encode("utf-8"))
    original = source.read_bytes()
    report = tmp_path / "must-not-publish.json"
    inspection = inspect_file(str(source))
    assert inspection.workflow_category != "markdown"
    assert inspection.requires_explicit_acceptance
    code = main(["validate", str(source), "--report", str(report), "--json"])
    payload = json.loads(capsys.readouterr().out)
    assert code == 2 and payload["error"]["code"] == "file_format_confirmation_required"
    assert not report.exists()
    assert source.read_bytes() == original
