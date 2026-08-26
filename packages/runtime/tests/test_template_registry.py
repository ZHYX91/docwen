"""Tests for strict canonical TemplateRegistry resolution."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

pytestmark = pytest.mark.integration

from docwen_runtime.templates.registry import (
    TemplateIdentityConflictError,
    TemplateNotFoundError,
    TemplateRegistry,
    TemplateResolutionError,
    is_canonical_template_id,
    validate_template_path,
)

# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _make_registry(files: list[str], tmp_path: Path) -> TemplateRegistry:
    """Create a registry populated with structurally valid OOXML packages."""
    for filename in files:
        path = tmp_path / filename
        target = "xlsx" if path.suffix.casefold() == ".xlsx" else "docx"
        _write_ooxml_template(path, target)
    return TemplateRegistry(tmp_path)


def _write_ooxml_template(path: Path, target: str) -> None:
    main_part = "word/document.xml" if target == "docx" else "xl/workbook.xml"
    content_type = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        if target == "docx"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    )
    root = "document" if target == "docx" else "workbook"
    namespace = (
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        if target == "docx"
        else "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    )
    with zipfile.ZipFile(path, "w") as package:
        package.writestr(
            "[Content_Types].xml",
            (
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                f'<Override PartName="/{main_part}" ContentType="{content_type}"/>'
                "</Types>"
            ),
        )
        package.writestr(
            "_rels/.rels",
            (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
                f'Target="{main_part}"/>'
                "</Relationships>"
            ),
        )
        package.writestr(main_part, f'<{root} xmlns="{namespace}"/>')


# ---------------------------------------------------------------------------
# exact match (case-sensitive)
# ---------------------------------------------------------------------------


class TestExactMatch:
    def test_same_registry_instance_observes_added_and_removed_templates(self, tmp_path: Path) -> None:
        first = tmp_path / "First.docx"
        second = tmp_path / "Second.docx"
        _write_ooxml_template(first, "docx")
        registry = TemplateRegistry(tmp_path)

        assert [template.name for template in registry.list_templates()] == ["First"]

        _write_ooxml_template(second, "docx")
        assert {template.name for template in registry.list_templates()} == {"First", "Second"}

        first.unlink()
        assert [template.name for template in registry.list_templates()] == ["Second"]

    def test_canonical_id_is_stable_and_resolves_without_display_name_inference(self, tmp_path: Path) -> None:
        first_dir = tmp_path / "first"
        second_dir = tmp_path / "second"
        first_dir.mkdir()
        second_dir.mkdir()
        _write_ooxml_template(first_dir / "Corporate Report.docx", "docx")
        _write_ooxml_template(second_dir / "Corporate Report.bin", "docx")
        with zipfile.ZipFile(second_dir / "Corporate Report.bin", "a") as package:
            package.writestr("custom/revision.xml", "<revision>2</revision>")

        first = TemplateRegistry(first_dir).list_templates()[0]
        second = TemplateRegistry(second_dir).list_templates()[0]

        assert first.id == "template.docx.4fd9bdd9a72279293b810865533f9662a4668e281a891e43d4d6f5939adc5c09"
        assert second.id == first.id
        assert TemplateRegistry(first_dir).get_template(first.id) == first
        assert first.to_dict()["id"] == first.id
        assert is_canonical_template_id(first.id)

    def test_unknown_canonical_id_never_matches_a_display_name(self, tmp_path: Path) -> None:
        canonical_looking_name = f"template.docx.{'0' * 64}"
        _write_ooxml_template(tmp_path / f"{canonical_looking_name}.docx", "docx")
        registry = TemplateRegistry(tmp_path)

        with pytest.raises(TemplateNotFoundError, match="资源 ID 不存在"):
            registry.get_template(canonical_looking_name, target_type="docx")

    def test_canonical_id_can_be_scoped_to_target_type(self, tmp_path: Path) -> None:
        reg = _make_registry(["Report.docx", "Report.xlsx"], tmp_path)
        xlsx = next(template for template in reg.list_templates() if template.target == "xlsx")
        tpl = reg.get_template(xlsx.id, target_type="xlsx")
        assert tpl.path.name == "Report.xlsx"
        assert tpl.target == "xlsx"


# ---------------------------------------------------------------------------
# error cases
# ---------------------------------------------------------------------------


class TestErrors:
    @pytest.mark.parametrize(
        ("first_name", "second_name"),
        [
            ("Report.docx", "report.DOCX"),
            ("Caf\u00e9.docx", "Cafe\u0301.docx"),
        ],
    )
    def test_normalized_identity_collision_fails_closed(
        self,
        tmp_path: Path,
        first_name: str,
        second_name: str,
    ) -> None:
        first_dir = tmp_path / "first"
        second_dir = tmp_path / "second"
        first_dir.mkdir()
        second_dir.mkdir()
        _write_ooxml_template(first_dir / first_name, "docx")
        _write_ooxml_template(second_dir / second_name, "docx")

        with pytest.raises(TemplateIdentityConflictError, match="Conflicting canonical template identity"):
            TemplateRegistry(first_dir, extra_paths=[second_dir]).list_templates()

    @pytest.mark.parametrize("invalid_id", ["", "   ", "GongWen", "GongWen.docx", "template.docx." + "A" * 64])
    def test_noncanonical_id_raises(self, tmp_path: Path, invalid_id: str) -> None:
        reg = _make_registry(["GongWen.docx"], tmp_path)
        with pytest.raises(TemplateNotFoundError, match="模板资源 ID 无效"):
            reg.get_template(invalid_id)

    def test_unknown_canonical_id_raises_without_name_suggestions(self, tmp_path: Path) -> None:
        reg = _make_registry(["GongWen.docx", "Report.xlsx"], tmp_path)
        missing = f"template.docx.{'f' * 64}"
        with pytest.raises(TemplateNotFoundError, match="资源 ID 不存在") as exc_info:
            reg.get_template(missing)
        assert "GongWen" not in str(exc_info.value)
        assert "Report" not in str(exc_info.value)

    def test_no_templates_raises(self, tmp_path: Path) -> None:
        reg = TemplateRegistry(tmp_path)
        with pytest.raises(TemplateNotFoundError, match="没有可用模板"):
            reg.get_template(f"template.docx.{'f' * 64}")


class TestContentValidation:
    def test_discovers_docx_content_with_misleading_suffix(self, tmp_path: Path) -> None:
        path = tmp_path / "Renamed Template.xlsx"
        _write_ooxml_template(path, "docx")

        templates = TemplateRegistry(tmp_path).list_templates()

        assert [(template.name, template.target) for template in templates] == [("Renamed Template", "docx")]

    def test_ignores_invalid_or_non_template_files(self, tmp_path: Path) -> None:
        (tmp_path / "broken.docx").write_bytes(b"not a package")
        (tmp_path / "notes.txt").write_text("hello", encoding="utf-8")

        assert TemplateRegistry(tmp_path).list_templates() == []

    def test_validate_template_path_rejects_wrong_content_target(self, tmp_path: Path) -> None:
        path = tmp_path / "looks-like-docx.docx"
        _write_ooxml_template(path, "xlsx")

        with pytest.raises(TemplateResolutionError) as exc_info:
            validate_template_path(path, expected_target="docx")

        assert exc_info.value.diagnostic_code == "TEMPLATE_FORMAT_MISMATCH"

    def test_validate_template_path_accepts_wrong_suffix_when_content_matches(self, tmp_path: Path) -> None:
        path = tmp_path / "template.bin"
        _write_ooxml_template(path, "xlsx")

        assert validate_template_path(path, expected_target="xlsx") == path.resolve()
