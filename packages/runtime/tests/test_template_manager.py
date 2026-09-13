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


def test_corrupt_state_is_preserved_and_blocks_discovery(tmp_path: Path) -> None:
    manager, _, user_dir = _manager(tmp_path)
    _write_ooxml_template(user_dir / "Mine.docx", "docx")
    manager.state_store.path.parent.mkdir(parents=True, exist_ok=True)
    manager.state_store.path.write_text("{broken", encoding="utf-8")
    with pytest.raises(ValueError, match="Cannot read template state"):
        manager.list_templates()
    assert manager.state_store.path.read_text(encoding="utf-8") == "{broken"


def test_failed_rename_restores_file_and_identity(tmp_path: Path, monkeypatch) -> None:
    manager, _, user_dir = _manager(tmp_path)
    _write_ooxml_template(user_dir / "Mine.docx", "docx")
    template = manager.list_templates()[0]
    before = manager.state_store.path.read_bytes()

    def fail_save(_state):
        raise OSError("disk unavailable")

    monkeypatch.setattr(manager.state_store, "save", fail_save)
    with pytest.raises(OSError, match="disk unavailable"):
        manager.rename_custom(template.id, "Renamed")
    assert template.path.is_file()
    assert not (user_dir / "Renamed.docx").exists()
    assert manager.state_store.path.read_bytes() == before


def test_external_rename_retains_identity_and_disable_state(tmp_path: Path) -> None:
    manager, _, user_dir = _manager(tmp_path)
    _write_ooxml_template(user_dir / "Mine.docx", "docx")
    template = manager.list_templates()[0]
    manager.set_enabled(template.id, False)
    template.path.rename(user_dir / "External.docx")
    renamed = manager.list_templates()[0]
    assert renamed.id == template.id
    assert not manager.is_enabled(renamed.id)


def test_disabling_default_clears_it_without_selecting_another(tmp_path: Path) -> None:
    manager, builtin_dir, _ = _manager(tmp_path)
    _write_ooxml_template(builtin_dir / "Alpha.docx", "docx")
    _write_ooxml_template(builtin_dir / "Beta.docx", "docx")
    template = manager.list_templates()[0]
    manager.set_default(template.id)
    assert manager.state_store.default_id("docx") == template.id
    manager.set_enabled(template.id, False)
    assert manager.state_store.default_id("docx") is None
    assert len(manager.registry.list_templates()) == 1


def test_same_display_name_keeps_distinct_resource_ids(tmp_path: Path) -> None:
    manager, builtin_dir, user_dir = _manager(tmp_path)
    _write_ooxml_template(builtin_dir / "Report.docx", "docx")
    _write_ooxml_template(user_dir / "Report.docx", "docx")
    templates = manager.list_templates()
    assert len(templates) == 2
    assert templates[0].id != templates[1].id


def test_reorder_rejects_stale_catalog(tmp_path: Path) -> None:
    manager, builtin_dir, _ = _manager(tmp_path)
    _write_ooxml_template(builtin_dir / "Alpha.docx", "docx")
    with pytest.raises(TemplateManagementError, match="catalog changed"):
        manager.set_order("docx", [])


def test_prepared_journal_recovers_template_bytes_and_state(tmp_path):
    from docwen_runtime import file_transactions as transactions

    manager, _, user_dir = _manager(tmp_path)
    _write_ooxml_template(user_dir / "Mine.docx", "docx")
    original = manager.list_templates()[0]
    before_bytes = original.path.read_bytes()
    before_state = manager.state_store.load()
    snapshots = [transactions.capture_user_file_preimage(path) for path in (manager.state_store.path, original.path)]
    transactions.write_transaction_journal(manager.state_store.path.parent, "templates", snapshots, state="PREPARED")
    original.path.write_bytes(b"interrupted replacement")
    recovered = manager.list_templates()[0]
    assert recovered.id == original.id
    assert recovered.path.read_bytes() == before_bytes
    recovered_state = manager.state_store.load()
    for entry in recovered_state["user_identities"].values():
        entry.pop("file_key", None)
    for entry in before_state["user_identities"].values():
        entry.pop("file_key", None)
    assert recovered_state == before_state
    assert not (manager.state_store.path.parent / transactions.CONFIG_JOURNAL_NAME).exists()


def test_replace_preserves_identity_order_enablement_and_default(tmp_path):
    manager, _, user_dir = _manager(tmp_path)
    _write_ooxml_template(user_dir / "Mine.docx", "docx")
    original = manager.list_templates()[0]
    manager.set_default(original.id)
    incoming = tmp_path / "incoming" / "Mine.docx"
    _write_ooxml_template(incoming, "docx")
    with zipfile.ZipFile(incoming, "a") as package:
        package.writestr("custom-marker.txt", "replacement")
    result = manager.import_template(incoming, replace_id=original.id)
    assert result.id == original.id
    assert result.path.read_bytes() == incoming.read_bytes()
    assert manager.state_store.default_id("docx") == original.id
    manager.set_enabled(original.id, False)
    manager.import_template(incoming, replace_id=original.id)
    assert not manager.is_enabled(original.id)


def test_trash_failure_restores_template_and_default(tmp_path, monkeypatch):
    manager, _, user_dir = _manager(tmp_path)
    _write_ooxml_template(user_dir / "Mine.docx", "docx")
    original = manager.list_templates()[0]
    manager.set_default(original.id)
    before = manager.state_store.path.read_bytes()

    def fail_trash(path):
        raise OSError("trash unavailable")

    monkeypatch.setattr("docwen_runtime.templates.manager._move_to_recycle_bin", fail_trash)
    with pytest.raises(TemplateManagementError, match="recycle bin"):
        manager.delete_custom(original.id)
    assert original.path.is_file()
    assert manager.state_store.path.read_bytes() == before


def test_package_upgrade_keeps_builtin_state_and_independent_custom_copy(tmp_path):
    manager, builtin_dir, user_dir = _manager(tmp_path)
    _write_ooxml_template(builtin_dir / "Report.docx", "docx")
    builtin = manager.list_templates()[0]
    custom = manager.copy_builtin_as_custom(builtin.id)
    custom_bytes = custom.path.read_bytes()
    manager.set_enabled(builtin.id, False)
    manager.set_default(custom.id)
    upgraded = tmp_path / "new-package-version" / "templates"
    _write_ooxml_template(upgraded / "Report.docx", "docx")
    registry = TemplateRegistry(
        upgraded, extra_paths=[user_dir], state_store=manager.state_store, managed_user_dir=user_dir
    )
    assert [item.id for item in registry.list_templates()] == [custom.id]
    assert manager.state_store.default_id("docx") == custom.id
    assert custom.path.read_bytes() == custom_bytes


@pytest.mark.parametrize("name", ["CON", "NUL.docx", "Report.", "bad\x01name"])
def test_invalid_portable_name_does_not_create_template(tmp_path, name):
    manager, builtin_dir, user_dir = _manager(tmp_path)
    _write_ooxml_template(builtin_dir / "Report.docx", "docx")
    builtin = manager.list_templates()[0]
    with pytest.raises(TemplateManagementError):
        manager.copy_builtin_as_custom(builtin.id, custom_name=name)
    assert not list(user_dir.iterdir())


def test_concurrent_process_imports_preserve_every_identity_and_order(tmp_path):
    import os
    import sys
    from concurrent.futures import ThreadPoolExecutor

    from tests.support.subprocess_runner import run_subprocess

    manager, builtin_dir, user_dir = _manager(tmp_path)
    incoming = tmp_path / "incoming" / "Report.docx"
    _write_ooxml_template(incoming, "docx")
    script = """
import sys
from pathlib import Path
from docwen_runtime.templates import TemplateManager, TemplateRegistry, TemplateStateStore
builtin, user, source = map(Path, sys.argv[1:])
state = TemplateStateStore(user.parent / 'template-state.json', user_dir=user)
manager = TemplateManager(
    TemplateRegistry(builtin, extra_paths=[user], state_store=state, managed_user_dir=user),
    state_store=state, user_dir=user,
)
for _ in range(4):
    manager.import_template(source)
"""
    command = [sys.executable, "-c", script, str(builtin_dir), str(user_dir), str(incoming)]
    with ThreadPoolExecutor(max_workers=2) as pool:
        pending = [pool.submit(run_subprocess, command, env=dict(os.environ), timeout=30) for _ in range(2)]
        results = [future.result() for future in pending]
    assert all(result.returncode == 0 for result in results), [result.stderr for result in results]
    templates = manager.list_templates()
    assert len(templates) == 8
    assert len({item.id for item in templates}) == 8
    assert len({item.name for item in templates}) == 8
    assert manager.state_store.load()["order"]["docx"] == [item.id for item in templates]
