"""Focused tests split from test_config_loader.py."""

from __future__ import annotations

from ._config_loader_support import (
    PROJECT_CONFIGS,
    Path,
    pytest,
    write_minimal_base_config_tree,
)

pytestmark = pytest.mark.unit


def test_default_user_config_dir_honors_internal_isolation_override(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    from docwen_runtime.config.loader import _default_user_config_dir

    isolated = tmp_path / "isolated-config"
    monkeypatch.setenv("DOCWEN_CONFIG_DIR", str(isolated))

    assert _default_user_config_dir() == isolated


class TestConfigLoadFromRealDirectory:
    """Verify the config system loads from the project configs/ dir."""

    def test_loads_real_config_with_expected_sections(self) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=PROJECT_CONFIGS)
        config = loader.config

        # Core sections from the new layout
        assert config.gui.window.default_mode in {"single", "batch"}
        assert config.output.directory.mode in {"source", "custom"}
        assert config.logger.enable in {True, False}

        # Domain-root sections
        assert isinstance(config.document.as_dict(), dict)
        assert isinstance(config.numbering.as_dict(), dict)
        assert isinstance(config.link.as_dict(), dict)
        assert isinstance(config.text.as_dict(), dict)
        assert isinstance(config.proofread.as_dict(), dict)

    def test_markdown_defaults_do_not_embed_cleanup(self) -> None:
        # The registry model puts clean rules in numbering/cleanup.toml
        # and add schemes in numbering/add.toml — two separate files.
        import tomllib

        add_content = tomllib.loads((PROJECT_CONFIGS / "numbering" / "add.toml").read_text(encoding="utf-8"))
        assert "cleanup" not in add_content

    def test_cleanup_rules_are_built_from_snapshot_without_core_mutation(self) -> None:
        from docwen_core.text import heading_numbering
        from docwen_runtime.config import build_heading_cleanup_rules
        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=PROJECT_CONFIGS)
        rules = build_heading_cleanup_rules(loader.config.as_dict())
        assert len(rules) == 8, f"Expected 8 request rules, got {len(rules)}"
        assert rules[0][0] == "legal_unit"
        assert not hasattr(heading_numbering, "_INJECTED_RULES")

    def test_legacy_cleanup_override_retains_new_shipped_rule(self, tmp_path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        user_dir = tmp_path / "user"
        user_cleanup = user_dir / "numbering" / "cleanup.toml"
        user_cleanup.parent.mkdir(parents=True)
        user_cleanup.write_text(
            """
[settings]
order = ["arabic_separator"]

[[rules]]
id = "arabic_separator"
enabled = false
pattern = "^USER$"
level = 3
""".strip()
            + "\n",
            encoding="utf-8",
        )

        loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
        cleanup = loader.config.as_dict()["numbering"]["cleanup"]
        rules_by_id = {rule["id"]: rule for rule in cleanup["rules"]}

        assert rules_by_id["arabic_separator"]["enabled"] is False
        assert rules_by_id["arabic_separator"]["pattern"] == "^USER$"
        assert rules_by_id["arabic_space"]["enabled"] is True
        editable_text = loader.get_file_text("numbering/cleanup.toml")
        assert editable_text is not None
        assert "arabic_space" in editable_text

    def test_config_tree_attribute_access(self) -> None:
        from docwen_runtime.config.loader import DocWenConfig

        cfg = DocWenConfig({"a": {"b": 42}, "c": "hello"})
        assert cfg.a.b == 42
        assert cfg.c == "hello"
        with pytest.raises(AttributeError):
            _ = cfg.nonexistent

    def test_config_as_dict_returns_deep_copy(self) -> None:
        from docwen_runtime.config.loader import DocWenConfig

        cfg = DocWenConfig({"a": {"b": 42}})
        d1 = cfg.as_dict()
        d2 = cfg.as_dict()
        d1["a"]["b"] = 99
        assert d2["a"]["b"] == 42
        assert d1 is not d2


class TestConfigReset:
    """Reset single file — delete user override to reveal base defaults."""

    def test_reset_file_deletes_user_override(self, tmp_path):
        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")
        (user_dir / "gui.toml").parent.mkdir(parents=True, exist_ok=True)
        (user_dir / "gui.toml").write_text('[window]\ndefault_mode = "batch"\n', encoding="utf-8")

        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.config.gui.window.default_mode == "batch"

        result = loader.reset_file("gui.toml")
        assert result is True
        assert loader.config.gui.window.default_mode == "single"
        assert not (user_dir / "gui.toml").exists()

    def test_reset_file_missing_user_file_is_ok(self, tmp_path):
        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.reset_file("gui.toml") is True

    def test_reset_excluded_files_are_not_reset(self, tmp_path):
        from docwen_runtime.config.loader import RESET_EXCLUDED, ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        for excluded in RESET_EXCLUDED:
            result = loader.reset_file(excluded)
            assert result is False, f"RESET_EXCLUDED file {excluded!r} must not be resettable"

    def test_reset_unknown_file_returns_false(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.reset_file("nonexistent.toml") is False

    def test_reset_unknown_group_returns_false(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.reset_group("nonexistent") is False

    def test_reset_group_middle_unlink_failure_restores_all_preimages(self, tmp_path, monkeypatch):
        from docwen_runtime.config.loader import ConfigLoader
        from docwen_runtime.config.registry import reset_plan_for_group

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        plan = reset_plan_for_group("conversion_defaults")
        assert len(plan.files) >= 3
        preimages: dict[str, bytes] = {}
        for index, rel_path in enumerate(plan.files[:3]):
            user_path = user_dir / rel_path
            user_path.parent.mkdir(parents=True, exist_ok=True)
            user_path.write_text(f"# preimage {index}\n", encoding="utf-8")
            preimages[rel_path] = user_path.read_bytes()

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        original_unlink = Path.unlink
        unlink_count = 0
        reset_targets = {user_dir / rel_path for rel_path in preimages}

        def fail_second_unlink(path: Path, *args, **kwargs) -> None:
            nonlocal unlink_count
            if path in reset_targets:
                unlink_count += 1
                if unlink_count == 2:
                    raise OSError("simulated middle-file failure")
            original_unlink(path, *args, **kwargs)

        monkeypatch.setattr(Path, "unlink", fail_second_unlink)

        assert loader.reset_group("conversion_defaults") is False
        assert unlink_count == 2
        for rel_path, preimage in preimages.items():
            assert (user_dir / rel_path).read_bytes() == preimage

    def test_text_reset_reload_failure_restores_all_dotted_file_preimages(
        self,
        tmp_path,
        monkeypatch,
    ):
        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "text.toml").write_text(
            'remove_numbering = false\nadd_numbering = false\nnumbering_scheme = "base"\n',
            encoding="utf-8",
        )
        (base_dir / "gui.toml").write_text(
            '[template]\nmd_default_template = "docx"\n',
            encoding="utf-8",
        )
        (base_dir / "numbering" / "add.toml").write_text(
            '[settings]\ndefault_scheme = "base"\n',
            encoding="utf-8",
        )
        (user_dir / "numbering").mkdir(parents=True)
        (user_dir / "text.toml").write_text(
            'remove_numbering = true\nadd_numbering = true\nnumbering_scheme = "custom"\n',
            encoding="utf-8",
        )
        (user_dir / "gui.toml").write_text(
            '[template]\nmd_default_template = "xlsx"\n',
            encoding="utf-8",
        )
        (user_dir / "numbering" / "add.toml").write_text(
            '[settings]\ndefault_scheme = "custom"\n',
            encoding="utf-8",
        )

        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        original_reload = loader.reload
        reload_count = 0

        def fail_once_then_reload() -> None:
            nonlocal reload_count
            reload_count += 1
            if reload_count == 1:
                raise OSError("simulated transient reset reload failure")
            original_reload()

        monkeypatch.setattr(loader, "reload", fail_once_then_reload)

        assert loader.reset_group("text") is False
        assert reload_count == 2
        assert loader.config.text.remove_numbering is True
        assert loader.config.text.add_numbering is True
        assert loader.config.text.numbering_scheme == "custom"
        assert loader.config.gui.template.md_default_template == "xlsx"
        assert loader.config.numbering.add.settings.default_scheme == "custom"

    def test_reset_all_middle_unlink_failure_restores_disk_and_cache(self, tmp_path, monkeypatch):
        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text(
            '[window]\ndefault_mode = "single"\n',
            encoding="utf-8",
        )
        (base_dir / "output.toml").write_text(
            '[directory]\nmode = "source"\n',
            encoding="utf-8",
        )
        (base_dir / "link.toml").write_text(
            '[format]\nimage_link_style = "wiki_embed"\n',
            encoding="utf-8",
        )
        (user_dir / "gui.toml").write_text(
            '[window]\ndefault_mode = "batch"\n',
            encoding="utf-8",
        )
        (user_dir / "output.toml").write_text(
            '[directory]\nmode = "custom"\n',
            encoding="utf-8",
        )
        (user_dir / "link.toml").write_text(
            '[format]\nimage_link_style = "markdown_link"\n',
            encoding="utf-8",
        )

        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        original_unlink = Path.unlink
        failed = False

        def fail_output_owner(path: Path, *args, **kwargs) -> None:
            nonlocal failed
            if path == user_dir / "output.toml" and not failed:
                failed = True
                raise OSError("simulated output reset failure")
            original_unlink(path, *args, **kwargs)

        monkeypatch.setattr(Path, "unlink", fail_output_owner)

        assert loader.reset_all() is False
        assert failed is True
        assert loader.config.gui.window.default_mode == "batch"
        assert loader.config.output.directory.mode == "custom"
        assert loader.config.link.format.image_link_style == "markdown_link"
        assert (user_dir / "gui.toml").exists()
        assert (user_dir / "output.toml").exists()
        assert (user_dir / "link.toml").exists()

    def test_reset_proofread_group_preserves_curated_user_dictionaries(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "proofread" / "engine.toml").write_text(
            "enable_symbol_pairing = true\n",
            encoding="utf-8",
        )
        (base_dir / "proofread" / "skip.toml").write_text(
            "code_blocks = true\nquote_blocks = false\nlog_skipped = true\n",
            encoding="utf-8",
        )
        (user_dir / "proofread").mkdir(parents=True)
        (user_dir / "proofread" / "engine.toml").write_text(
            "enable_symbol_pairing = false\n",
            encoding="utf-8",
        )
        (user_dir / "proofread" / "typos.toml").write_text(
            '[entries]\nteh = ["the"]\n',
            encoding="utf-8",
        )
        (user_dir / "proofread" / "sensitive_words.toml").write_text(
            'words = ["keep-me"]\n',
            encoding="utf-8",
        )
        (user_dir / "proofread" / "skip.toml").write_text(
            "code_blocks = false\nlog_skipped = false\n",
            encoding="utf-8",
        )
        (user_dir / "proofread" / "pairs.toml").write_text(
            'items = [["<", ">"]]\n',
            encoding="utf-8",
        )
        (user_dir / "proofread" / "symbol_map.toml").write_text(
            '[entries]\n"!" = ["！"]\n',
            encoding="utf-8",
        )

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.config.proofread.engine.enable_symbol_pairing is False

        assert loader.reset_group("proofread") is True

        assert loader.config.proofread.engine.enable_symbol_pairing is True
        assert (user_dir / "proofread" / "engine.toml").read_text(encoding="utf-8").strip() == ""
        assert loader.config.proofread.skip.code_blocks is True
        assert loader.config.proofread.skip.log_skipped is False
        assert loader.config.proofread.pairs.items == [["<", ">"]]
        assert loader.config.proofread.symbol_map.entries.as_dict() == {"!": ["！"]}
        assert (user_dir / "proofread" / "pairs.toml").exists()
        assert (user_dir / "proofread" / "symbol_map.toml").exists()
        assert (user_dir / "proofread" / "typos.toml").exists()
        assert (user_dir / "proofread" / "sensitive_words.toml").exists()

    def test_reset_all_resets_pairs_but_byte_preserves_curated_dictionaries(self, tmp_path):
        from docwen_runtime.config.loader import ConfigLoader

        user_dir = tmp_path / "user"
        loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
        assert loader.set_value("proofread.pairs.items", [["<", ">"]]) is True
        assert loader.set_value("proofread.symbol_map.entries.!", ["！"]) is True
        assert loader.set_value("proofread.typos.entries.correct", ["mistkae"]) is True
        assert loader.set_value("proofread.sensitive_words.entries.secret", ["allowed"]) is True

        protected_paths = (
            user_dir / "proofread" / "symbol_map.toml",
            user_dir / "proofread" / "typos.toml",
            user_dir / "proofread" / "sensitive_words.toml",
        )
        protected_bytes = {path: path.read_bytes() for path in protected_paths}

        assert loader.reset_all() is True

        assert not (user_dir / "proofread" / "pairs.toml").exists()
        assert loader.config.proofread.pairs.items != [["<", ">"]]
        assert {path: path.read_bytes() for path in protected_paths} == protected_bytes
        proofread = loader.config.as_dict()["proofread"]
        assert proofread["symbol_map"]["entries"]["!"] == ["！"]
        assert proofread["typos"]["entries"]["correct"] == ["mistkae"]
        assert proofread["sensitive_words"]["entries"]["secret"] == ["allowed"]

    def test_reset_section_deletes_user_override(self, tmp_path):
        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "link.toml").write_text('[format]\nimage_link_style = "wiki_embed"\n', encoding="utf-8")
        (user_dir / "link.toml").parent.mkdir(parents=True, exist_ok=True)
        (user_dir / "link.toml").write_text('[format]\nimage_link_style = "markdown_link"\n', encoding="utf-8")

        from docwen_runtime.config.loader import ConfigLoader

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.config.link.format.image_link_style == "markdown_link"

        assert loader.reset_section("link") is True
        assert not (user_dir / "link.toml").exists()
        assert loader.config.link.format.image_link_style == "wiki_embed"
