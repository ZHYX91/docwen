"""Tests for writable user-template management and MSIX-safe discovery."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docwen_runtime.templates import (
    TemplateManagementError,
    TemplateManager,
    TemplateNotFoundError,
    TemplateRegistry,
    TemplateStateStore,
)

pytestmark = pytest.mark.integration


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
    path.parent.mkdir(parents=True, exist_ok=True)
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


def _manager(tmp_path: Path) -> tuple[TemplateManager, Path, Path]:
    builtin_dir = tmp_path / "bundled" / "templates"
    user_dir = tmp_path / "user" / "templates"
    builtin_dir.mkdir(parents=True)
    user_dir.mkdir(parents=True)
    state = TemplateStateStore(tmp_path / "user" / "template-state.json", user_dir=user_dir)
    registry = TemplateRegistry(
        builtin_dir,
        extra_paths=[user_dir],
        state_store=state,
        managed_user_dir=user_dir,
    )
    return TemplateManager(registry, state_store=state, user_dir=user_dir), builtin_dir, user_dir


def test_registry_merges_bundled_and_user_templates(tmp_path: Path) -> None:
    manager, builtin_dir, user_dir = _manager(tmp_path)
    _write_ooxml_template(builtin_dir / "Built In.docx", "docx")
    _write_ooxml_template(user_dir / "Mine.xlsx", "xlsx")

    templates = manager.list_templates()

    assert [(item.name, item.target) for item in templates] == [("Built In", "docx"), ("Mine", "xlsx")]
    assert manager.is_custom(templates[0]) is False
    assert manager.is_custom(templates[1]) is True


def test_import_keeps_display_names_unambiguous_against_bundled_template(tmp_path: Path) -> None:
    manager, builtin_dir, _user_dir = _manager(tmp_path)
    _write_ooxml_template(builtin_dir / "Report.docx", "docx")
    incoming = tmp_path / "incoming" / "Report.docx"
    _write_ooxml_template(incoming, "docx")

    imported = manager.import_template(incoming)

    assert imported.name == "Report (2)"
    assert manager.is_custom(imported)
    assert imported.path.parent == manager.user_dir


def test_user_identity_survives_managed_rename(tmp_path: Path) -> None:
    manager, _builtin_dir, _user_dir = _manager(tmp_path)
    incoming = tmp_path / "incoming.docx"
    _write_ooxml_template(incoming, "docx")
    imported = manager.import_template(incoming)

    renamed = manager.rename_custom(imported.id, "Renamed")

    assert renamed.id == imported.id
    assert renamed.name == "Renamed"
    assert renamed.path.name == "Renamed.docx"
    assert manager.registry.get_template(imported.id).name == "Renamed"


def test_disabled_template_is_hidden_from_normal_registry_but_retained_for_management(tmp_path: Path) -> None:
    manager, builtin_dir, _user_dir = _manager(tmp_path)
    _write_ooxml_template(builtin_dir / "Default.docx", "docx")
    template = manager.list_templates()[0]

    manager.set_enabled(template.id, False)

    assert manager.registry.list_templates("docx") == []
    assert [item.id for item in manager.list_templates("docx", include_disabled=True)] == [template.id]
    with pytest.raises(TemplateNotFoundError):
        manager.registry.get_template(template.id, "docx")


def test_copy_builtin_creates_independent_enabled_custom_template_after_source(tmp_path: Path) -> None:
    manager, builtin_dir, _user_dir = _manager(tmp_path)
    _write_ooxml_template(builtin_dir / "Alpha.docx", "docx")
    _write_ooxml_template(builtin_dir / "Beta.docx", "docx")
    alpha = manager.list_templates("docx")[0]

    copied = manager.copy_builtin_as_custom(alpha.id, custom_name="Alpha Custom")
    ordered = manager.list_templates("docx", include_disabled=True)

    assert manager.is_custom(copied)
    assert manager.is_enabled(copied.id)
    assert [item.name for item in ordered] == ["Alpha", "Alpha Custom", "Beta"]
    assert copied.path.read_bytes() == alpha.path.read_bytes()


def test_move_and_disabled_state_survive_new_registry_instance(tmp_path: Path) -> None:
    manager, builtin_dir, user_dir = _manager(tmp_path)
    _write_ooxml_template(builtin_dir / "Alpha.docx", "docx")
    _write_ooxml_template(builtin_dir / "Beta.docx", "docx")
    alpha, beta = manager.list_templates("docx")

    manager.move(beta.id, -1)
    manager.set_enabled(alpha.id, False)

    state = TemplateStateStore(tmp_path / "user" / "template-state.json", user_dir=user_dir)
    registry = TemplateRegistry(
        builtin_dir,
        extra_paths=[user_dir],
        state_store=state,
        managed_user_dir=user_dir,
    )
    assert [item.id for item in registry.list_templates("docx", include_disabled=True)] == [beta.id, alpha.id]
    assert [item.id for item in registry.list_templates("docx")] == [beta.id]


def test_built_in_template_cannot_be_renamed_or_exported_as_custom(tmp_path: Path) -> None:
    manager, builtin_dir, _user_dir = _manager(tmp_path)
    _write_ooxml_template(builtin_dir / "Default.docx", "docx")
    template = manager.list_templates("docx")[0]

    with pytest.raises(TemplateManagementError, match="Built-in templates are read-only"):
        manager.rename_custom(template.id, "Changed")
    with pytest.raises(TemplateManagementError, match="Built-in templates are read-only"):
        manager.export_custom(template.id, tmp_path / "export.docx")
