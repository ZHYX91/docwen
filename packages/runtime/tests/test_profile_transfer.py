"""Current profile transfer preserves user configuration and template identity."""

import shutil

import pytest

from docwen_runtime import profile_paths

pytestmark = pytest.mark.integration


@pytest.mark.parametrize("portable", [False, True])
def test_copying_profile_preserves_rules_template_ids_order_and_default(tmp_path, monkeypatch, portable):
    from docx import Document

    from docwen_runtime.config import ConfigLoader
    from docwen_runtime.templates import TemplateManager

    for key in ("DOCWEN_CONFIG_DIR", "DOCWEN_DATA_DIR", "DOCWEN_LOG_DIR", "DOCWEN_LOG_TO_TEMP"):
        monkeypatch.delenv(key, raising=False)
    monkeypatch.setattr(profile_paths.sys, "frozen", portable, raising=False)
    monkeypatch.setattr(profile_paths, "_windows_has_package_identity", lambda: False)
    source, destination = (tmp_path / name / "data" for name in ("original", "copy"))

    def select(root):
        if portable:
            monkeypatch.setattr(profile_paths.sys, "executable", str(root.parent / "DocWenCLI.exe"))
        else:
            monkeypatch.setenv("DOCWEN_DATA_DIR", str(root))

    def catalog(manager):
        return [
            (item.id, item.name, item.is_default, manager.is_enabled(item.id))
            for item in manager.list_templates("docx")
        ]

    overrides = {"logger": {"enable": False, "console_enable": False}}
    select(source)
    with profile_paths.bind_process_profile():
        config = ConfigLoader(runtime_overrides=overrides)
        assert config.set_values({"gui.language.locale": "en_US", "proofread.typos.entries.correct": ["mistkae"]})
        manager = TemplateManager.default()
        imported = []
        for name in ("Personal A", "Personal B"):
            incoming = tmp_path / f"{name}.docx"
            document = Document()
            document.add_paragraph(name)
            document.save(str(incoming))
            imported.append(manager.import_template(incoming))
        manager.set_default(imported[1].id)
        manager.set_enabled(imported[0].id, False)
        manager.set_order("docx", list(reversed([item.id for item in manager.list_templates("docx")])))
        expected = catalog(manager)
        original_instance = profile_paths.profile_instance_name()
    before = {path.relative_to(source): path.read_bytes() for path in source.rglob("*") if path.is_file()}
    shutil.copytree(source, destination)
    select(destination)
    with profile_paths.bind_process_profile() as profile:
        config = ConfigLoader(runtime_overrides=overrides)
        assert config.user_dir == destination / "configs" and profile.data_dir == destination
        assert config.config.gui.language.locale == "en_US"
        assert config.config.proofread.typos.entries.correct == ["mistkae"]
        manager = TemplateManager.default()
        assert catalog(manager) == expected
        assert all(
            item.path.parent == destination / "templates"
            for item in manager.list_templates()
            if item.origin == "custom"
        )
        assert profile_paths.profile_instance_name() != original_instance
    assert {path.relative_to(source): path.read_bytes() for path in source.rglob("*") if path.is_file()} == before
