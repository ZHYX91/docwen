"""Focused tests split from test_config_loader.py."""

from __future__ import annotations

import pytest

from ._config_loader_support import (
    Path,
    write_minimal_base_config_tree,
)

pytestmark = pytest.mark.unit


class TestSetValue:
    """Writing individual config values via set_value."""

    def test_set_value_writes_and_reloads(self, tmp_path):
        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")

        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.set_value("gui.window.default_mode", "batch") is True
        assert loader.config.gui.window.default_mode == "batch"

    def test_set_value_routes_to_correct_file(self, tmp_path):
        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")
        (base_dir / "numbering").mkdir(parents=True, exist_ok=True)
        (base_dir / "numbering" / "add.toml").write_text(
            '[settings]\ndefault_scheme = "gongwen_standard"\n', encoding="utf-8"
        )

        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

        # GUI setting → gui.toml
        assert loader.set_value("gui.window.default_mode", "batch") is True
        assert loader.config.gui.window.default_mode == "batch"

        # Numbering setting → numbering/add.toml
        assert loader.set_value("numbering.add.settings.default_scheme", "legal_standard") is True
        assert loader.config.numbering.add.settings.default_scheme == "legal_standard"

    def test_set_values_groups_writes_and_reloads_once(self, tmp_path, monkeypatch):
        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")
        (base_dir / "output.toml").write_text('[directory]\nmode = "source"\n', encoding="utf-8")

        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        reload_count = 0
        original_reload = loader.reload

        def counting_reload() -> None:
            nonlocal reload_count
            reload_count += 1
            original_reload()

        monkeypatch.setattr(loader, "reload", counting_reload)

        assert (
            loader.set_values(
                {
                    "gui.window.default_mode": "batch",
                    "output.directory.mode": "custom",
                    "output.directory.custom_path": "D:/out",
                }
            )
            is True
        )

        assert reload_count == 1
        assert loader.config.gui.window.default_mode == "batch"
        assert loader.config.output.directory.mode == "custom"
        assert loader.config.output.directory.custom_path == "D:/out"
        assert (user_dir / "gui.toml").exists()
        assert (user_dir / "output.toml").exists()

    def test_set_values_later_file_failure_restores_missing_preimages(self, tmp_path, monkeypatch):
        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")
        (base_dir / "output.toml").write_text('[directory]\nmode = "source"\n', encoding="utf-8")

        from docwen_runtime.config import loader as loader_module
        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        original_write = loader_module.write_toml_file

        def fail_output_write(path, data) -> None:
            if path.name == "output.toml":
                raise OSError("simulated later-file failure")
            original_write(path, data)

        monkeypatch.setattr(loader_module, "write_toml_file", fail_output_write)

        assert (
            loader.set_values(
                {
                    "gui.window.default_mode": "batch",
                    "output.directory.mode": "custom",
                }
            )
            is False
        )

        assert not (user_dir / "gui.toml").exists()
        assert not (user_dir / "output.toml").exists()
        assert loader.config.gui.window.default_mode == "single"
        assert loader.config.output.directory.mode == "source"

    def test_set_values_reload_failure_restores_disk_and_cache(self, tmp_path, monkeypatch):
        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")

        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        original_reload = loader.reload
        reload_count = 0

        def fail_once_then_reload() -> None:
            nonlocal reload_count
            reload_count += 1
            if reload_count == 1:
                raise OSError("simulated transient reload failure")
            original_reload()

        monkeypatch.setattr(loader, "reload", fail_once_then_reload)

        assert loader.set_values({"gui.window.default_mode": "batch"}) is False
        assert reload_count == 2
        assert loader.config.gui.window.default_mode == "single"
        assert not (user_dir / "gui.toml").exists()

    def test_set_value_unknown_root_returns_false(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.set_value("nonexistent_root.some_key", "value") is False

    def test_reset_values_deletes_selected_user_overrides_only(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "software.toml").write_text(
            "\n".join(
                [
                    "[default_priority]",
                    'word_processors = ["wps_writer", "msoffice_word", "libreoffice"]',
                    'spreadsheet_processors = ["wps_spreadsheets", "msoffice_excel", "libreoffice"]',
                    "",
                ]
            ),
            encoding="utf-8",
        )
        (user_dir / "software.toml").write_text(
            "\n".join(
                [
                    "[default_priority]",
                    'word_processors = ["libreoffice", "msoffice_word", "wps_writer"]',
                    'spreadsheet_processors = ["libreoffice", "msoffice_excel", "wps_spreadsheets"]',
                    "",
                ]
            ),
            encoding="utf-8",
        )

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.config.software.default_priority.word_processors[0] == "libreoffice"
        assert loader.config.software.default_priority.spreadsheet_processors[0] == "libreoffice"

        assert loader.reset_values(["software.default_priority.word_processors"]) is True

        assert loader.config.software.default_priority.word_processors == [
            "wps_writer",
            "msoffice_word",
            "libreoffice",
        ]
        assert loader.config.software.default_priority.spreadsheet_processors == [
            "libreoffice",
            "msoffice_excel",
            "wps_spreadsheets",
        ]
        user_text = (user_dir / "software.toml").read_text(encoding="utf-8")
        assert "word_processors" not in user_text
        assert "spreadsheet_processors" in user_text


class TestDeepMerge:
    def test_scalar_overwrite(self) -> None:
        from docwen_runtime.config.loader import deep_merge

        assert deep_merge({"a": 1}, {"a": 2}) == {"a": 2}

    def test_nested_merge(self) -> None:
        from docwen_runtime.config.loader import deep_merge

        result = deep_merge(
            {"a": {"b": 1, "c": 2}},
            {"a": {"b": 99}, "d": 3},
        )
        assert result == {"a": {"b": 99, "c": 2}, "d": 3}

    def test_new_key_added(self) -> None:
        from docwen_runtime.config.loader import deep_merge

        result = deep_merge({"a": 1}, {"b": 2})
        assert result == {"a": 1, "b": 2}

    def test_default_unchanged(self) -> None:
        from docwen_runtime.config.loader import deep_merge

        default = {"a": {"b": 1}}
        user = {"a": {"c": 2}}
        deep_merge(default, user)
        assert default == {"a": {"b": 1}}, "deep_merge must not mutate default"


class TestConfigLoaderUpdateSectionsPreservingComments:
    """ConfigLoader.update_file_sections preserves comments in untouched sections."""

    def test_preserves_comments_in_untouched_sections(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        (user_dir / "gui.toml").parent.mkdir(parents=True, exist_ok=True)
        (user_dir / "gui.toml").write_text(
            '# user-comment\n[window]\ndefault_mode = "single"\n[theme]\n# theme-comment\ndefault_theme = "light"\n',
            encoding="utf-8",
        )

        ok = loader.update_file_sections(
            "gui.toml",
            {"window": {"default_mode": "batch"}},
        )
        assert ok is True

        text = (user_dir / "gui.toml").read_text(encoding="utf-8")
        assert "# user-comment" in text
        assert "# theme-comment" in text
        assert 'default_mode = "batch"' in text

    def test_unknown_rel_path_returns_false(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.update_file_sections("nonexistent.toml", {}) is False

    def test_empty_sections_is_noop_but_still_writes(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        (user_dir / "gui.toml").parent.mkdir(parents=True, exist_ok=True)
        (user_dir / "gui.toml").write_text(
            '# keep\n[window]\ndefault_mode = "single"\n',
            encoding="utf-8",
        )

        assert loader.update_file_sections("gui.toml", {}) is True
        assert "# keep" in (user_dir / "gui.toml").read_text(encoding="utf-8")


class TestConfigLoaderUpdateFileDocument:
    """ConfigLoader.update_file_document hands the tomlkit document to the caller
    for fine-grained mutation with per-value inline comment preservation."""

    def test_preserves_inline_comments_via_mutate_callback(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        # typos.toml is RESET_EXCLUDED, but reload still needs base file to exist
        (base_dir / "proofread").mkdir(parents=True, exist_ok=True)
        (base_dir / "proofread" / "typos.toml").write_text("", encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

        (user_dir / "proofread").mkdir(parents=True, exist_ok=True)
        (user_dir / "proofread" / "typos.toml").write_text(
            '[typos]\nwrong = ["correct"]  # user-remark\n',
            encoding="utf-8",
        )

        def mutate(doc):
            import tomlkit

            tbl = tomlkit.table()
            val = tomlkit.item(["correct1", "correct2"])
            val.comment("new-remark")
            tbl["newword"] = val
            doc["typos"] = tbl

        ok = loader.update_file_document("proofread/typos.toml", mutate)
        assert ok is True

        text = (user_dir / "proofread" / "typos.toml").read_text(encoding="utf-8")
        assert "newword" in text
        assert "# new-remark" in text

    def test_unknown_rel_path_returns_false(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.update_file_document("nonexistent.toml", lambda doc: None) is False

    def test_mutate_exception_returns_false_no_corrupt(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "proofread").mkdir(parents=True, exist_ok=True)
        (base_dir / "proofread" / "typos.toml").write_text("", encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

        original = '[typos]\nwrong = ["correct"]\n'
        (user_dir / "proofread").mkdir(parents=True, exist_ok=True)
        (user_dir / "proofread" / "typos.toml").write_text(original, encoding="utf-8")

        def bad_mutate(doc):
            raise RuntimeError("boom")

        assert loader.update_file_document("proofread/typos.toml", bad_mutate) is False
        assert (user_dir / "proofread" / "typos.toml").read_text(encoding="utf-8") == original

    def test_transient_reload_failure_returns_false_and_restores_document(
        self,
        tmp_path,
        monkeypatch,
    ):
        import tomlkit

        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text(
            '[window]\ndefault_mode = "single"\n',
            encoding="utf-8",
        )
        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        original_reload = loader.reload
        reload_count = 0

        def fail_once_then_reload() -> None:
            nonlocal reload_count
            reload_count += 1
            if reload_count == 1:
                raise OSError("simulated transient document reload failure")
            original_reload()

        def mutate(doc) -> None:
            window = tomlkit.table()
            window["default_mode"] = "batch"
            doc["window"] = window

        monkeypatch.setattr(loader, "reload", fail_once_then_reload)

        assert loader.update_file_document("gui.toml", mutate) is False
        assert reload_count == 2
        assert loader.config.gui.window.default_mode == "single"
        assert not (user_dir / "gui.toml").exists()

    def test_update_file_document_contains_parent_directory_creation_failure(
        self,
        tmp_path,
        monkeypatch,
    ):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        original_mkdir = Path.mkdir

        def fail_user_directory(path: Path, *args, **kwargs) -> None:
            if path == user_dir:
                raise OSError("simulated user config directory failure")
            original_mkdir(path, *args, **kwargs)

        monkeypatch.setattr(Path, "mkdir", fail_user_directory)

        assert loader.update_file_document("gui.toml", lambda _doc: None) is False
        assert not user_dir.exists()


class TestConfigLoaderGetFileDict:
    """get_file_dict reads raw on-disk user override content."""

    def test_reads_raw_user_content_without_base_defaults(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "link.toml").write_text('[format]\nimage_link_style = "wiki_embed"\n', encoding="utf-8")
        # Write user override with only one key
        (user_dir / "link.toml").parent.mkdir(parents=True, exist_ok=True)
        (user_dir / "link.toml").write_text('[format]\nimage_link_style = "markdown_link"\n', encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        raw = loader.get_file_dict("link.toml")
        assert raw.get("format", {}).get("image_link_style") == "markdown_link"

    def test_get_base_file_dict_returns_base_content(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "link.toml").write_text('[format]\nimage_link_style = "wiki_embed"\n', encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        base_raw = loader.get_base_file_dict("link.toml")
        assert base_raw.get("format", {}).get("image_link_style") == "wiki_embed"

    def test_missing_user_file_returns_empty_dict(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.get_file_dict("link.toml") == {}

    def test_unknown_rel_path_returns_empty_dict(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.get_file_dict("nonexistent.toml") == {}
