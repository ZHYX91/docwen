"""Focused tests split from test_config_loader.py."""

from __future__ import annotations

from ._config_loader_support import (
    Path,
    pytest,
    write_minimal_base_config_tree,
)

pytestmark = pytest.mark.unit


class TestConfigLoaderEditableFileText:
    def test_missing_user_file_returns_shipped_source_verbatim(self, tmp_path: Path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        source = '# shipped remark\n[entries]\nfoo = ["bar"] # item remark\n'
        (base_dir / "proofread" / "typos.toml").write_text(source, encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

        assert loader.get_file_text("proofread/typos.toml") == source

    def test_sparse_user_source_overlays_base_with_comments(self, tmp_path: Path) -> None:
        import tomllib

        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "link.toml").write_text(
            '# shipped\n[format]\nimage_link_style = "wiki_embed"\nmd_file_link_style = "wiki_link"\n',
            encoding="utf-8",
        )
        user_dir.mkdir()
        (user_dir / "link.toml").write_text(
            '[format]\nimage_link_style = "markdown_link" # user choice\n',
            encoding="utf-8",
        )

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        source = loader.get_file_text("link.toml")

        assert source is not None
        parsed = tomllib.loads(source)
        assert parsed["format"] == {
            "image_link_style": "markdown_link",
            "md_file_link_style": "wiki_link",
        }
        assert "# shipped" in source
        assert "# user choice" in source

    def test_curated_entries_replace_base_so_editor_deletion_round_trips(self, tmp_path: Path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "proofread" / "symbol_map.toml").write_text(
            '[entries]\nzero = ["fullwidth-zero"]\none = ["fullwidth-one"]\n',
            encoding="utf-8",
        )
        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

        assert loader.save_file_text(
            "proofread/symbol_map.toml",
            '# user header\n[entries]\n# kept entry\none = ["fullwidth-one"] # user inline\n',
        )

        entries = loader.config.as_dict()["proofread"]["symbol_map"]["entries"]
        assert entries == {"one": ["fullwidth-one"]}
        source = loader.get_file_text("proofread/symbol_map.toml")
        assert source is not None
        assert "zero" not in source
        assert "# user header" in source
        assert "# kept entry" in source
        assert "# user inline" in source
        assert "__docwen_" not in (user_dir / "proofread" / "symbol_map.toml").read_text(encoding="utf-8")

    def test_existing_curated_user_section_is_complete_across_shipped_upgrades(
        self,
        tmp_path: Path,
    ) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "proofread" / "symbol_map.toml").write_text(
            '[entries]\nzero = ["shipped-zero"]\none = ["shipped-one"]\n',
            encoding="utf-8",
        )
        (user_dir / "proofread").mkdir(parents=True)
        (user_dir / "proofread" / "symbol_map.toml").write_text(
            '# user header\n[entries]\n# user entry lead\none = ["legacy-override"] # user inline\n',
            encoding="utf-8",
        )

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

        assert loader.config.as_dict()["proofread"]["symbol_map"]["entries"] == {
            "one": ["legacy-override"],
        }
        source = loader.get_file_text("proofread/symbol_map.toml")
        assert source is not None
        assert "shipped-zero" not in source
        assert "legacy-override" in source
        assert "# user header" in source
        assert "# user entry lead" in source
        assert "# user inline" in source

        (base_dir / "proofread" / "symbol_map.toml").write_text(
            '[entries]\nzero = ["shipped-zero-v2"]\none = ["shipped-one-v2"]\ntwo = ["new-in-upgrade"]\n',
            encoding="utf-8",
        )
        loader.reload()

        assert loader.config.as_dict()["proofread"]["symbol_map"]["entries"] == {
            "one": ["legacy-override"],
        }

    def test_empty_curated_section_suppresses_shipped_entries(self, tmp_path: Path) -> None:
        import tomllib

        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "proofread" / "typos.toml").write_text(
            '[entries]\nzero = ["base-zero"]\none = ["base-one"]\n',
            encoding="utf-8",
        )
        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

        assert loader.save_file_text("proofread/typos.toml", "[entries]\n")

        assert loader.config.as_dict()["proofread"]["typos"]["entries"] == {}
        source = loader.get_file_text("proofread/typos.toml")
        assert source is not None
        assert tomllib.loads(source)["entries"] == {}

    def test_replacement_wins_if_a_future_section_is_also_keyed(self) -> None:
        from docwen_runtime.config.loader import _merge_file_layers  # pyright: ignore[reportPrivateUsage]
        from docwen_runtime.config.registry import ConfigFileSpec

        spec = ConfigFileSpec(
            "future.toml",
            ("future",),
            replace_sections=frozenset({"rules"}),
            keyed_list_sections=frozenset({"rules"}),
        )

        merged = _merge_file_layers(
            spec,
            {"rules": [{"id": "shipped", "enabled": True}]},
            {"rules": [{"id": "user", "enabled": False}]},
        )

        assert merged["rules"] == [{"id": "user", "enabled": False}]

    def test_whole_section_writes_use_registry_replacement(self, tmp_path: Path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "proofread" / "typos.toml").write_text(
            '[entries]\nshipped = ["base"]\n',
            encoding="utf-8",
        )
        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

        assert loader.set_value("proofread.typos.entries", {"custom": ["fix"]})
        assert loader.config.as_dict()["proofread"]["typos"]["entries"] == {
            "custom": ["fix"],
        }

        raw = loader.get_file_dict("proofread/typos.toml")
        assert loader.write_file_content("proofread/typos.toml", raw)
        assert loader.config.as_dict()["proofread"]["typos"]["entries"] == {
            "custom": ["fix"],
        }

    def test_batched_leaf_writes_create_one_complete_replacement(self, tmp_path: Path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "proofread" / "typos.toml").write_text(
            '[entries]\nshipped = ["base"]\n',
            encoding="utf-8",
        )
        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

        assert loader.set_values(
            {
                "proofread.typos.entries.first": ["one"],
                "proofread.typos.entries.second": ["two"],
            }
        )

        assert loader.config.as_dict()["proofread"]["typos"]["entries"] == {
            "first": ["one"],
            "second": ["two"],
        }
        assert loader.get_file_dict("proofread/typos.toml")["entries"] == {
            "first": ["one"],
            "second": ["two"],
        }

    def test_reset_complete_section_reveals_base_then_leaf_write_starts_new_replacement(
        self,
        tmp_path: Path,
    ) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "proofread" / "symbol_map.toml").write_text(
            '[entries]\nzero = ["shipped-zero"]\none = ["shipped-one"]\n',
            encoding="utf-8",
        )
        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

        assert loader.set_value(
            "proofread.symbol_map.entries",
            {"custom": ["complete"]},
        )
        assert loader.reset_values(("proofread.symbol_map.entries",))
        assert loader.config.as_dict()["proofread"]["symbol_map"]["entries"] == {
            "zero": ["shipped-zero"],
            "one": ["shipped-one"],
        }
        assert loader.set_value(
            "proofread.symbol_map.entries.later",
            ["replacement-leaf"],
        )

        assert loader.config.as_dict()["proofread"]["symbol_map"]["entries"] == {
            "later": ["replacement-leaf"],
        }
        user_text = (user_dir / "proofread" / "symbol_map.toml").read_text(encoding="utf-8")
        assert "__docwen_" not in user_text

    def test_reset_replacement_base_leaf_materializes_default_without_reviving_siblings(
        self,
        tmp_path: Path,
    ) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "proofread" / "symbol_map.toml").write_text(
            '[entries]\nzero = ["shipped-zero"]\none = ["shipped-one"]\n',
            encoding="utf-8",
        )
        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.set_value(
            "proofread.symbol_map.entries",
            {"zero": ["user-zero"]},
        )

        assert loader.reset_values(("proofread.symbol_map.entries.zero",))

        assert loader.config.as_dict()["proofread"]["symbol_map"]["entries"] == {
            "zero": ["shipped-zero"],
        }
        user_text = (user_dir / "proofread" / "symbol_map.toml").read_text(encoding="utf-8")
        assert "[entries]" in user_text
        assert "shipped-zero" in user_text
        assert "shipped-one" not in user_text
        assert "__docwen_" not in user_text

    def test_reset_replacement_custom_leaf_keeps_shipped_entries_deleted(
        self,
        tmp_path: Path,
    ) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "proofread" / "typos.toml").write_text(
            '[entries]\nzero = ["base-zero"]\none = ["base-one"]\n',
            encoding="utf-8",
        )
        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.set_value(
            "proofread.typos.entries",
            {"custom": ["fix"]},
        )

        assert loader.reset_values(("proofread.typos.entries.custom",))

        assert loader.config.as_dict()["proofread"]["typos"]["entries"] == {}
        user_text = (user_dir / "proofread" / "typos.toml").read_text(encoding="utf-8")
        assert "[entries]" in user_text
        assert "custom" not in user_text
        assert "__docwen_" not in user_text

    def test_leaf_write_creates_complete_replacement_without_hidden_metadata(
        self,
        tmp_path: Path,
    ) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "proofread" / "typos.toml").write_text(
            '[entries]\nshipped = ["base"]\n',
            encoding="utf-8",
        )
        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

        assert loader.save_file_text("proofread/typos.toml", "# no owned section\n")
        assert loader.set_value("proofread.typos.entries.custom", ["leaf"])

        assert loader.config.as_dict()["proofread"]["typos"]["entries"] == {
            "custom": ["leaf"],
        }
        user_text = (user_dir / "proofread" / "typos.toml").read_text(encoding="utf-8")
        assert "__docwen_" not in user_text

    def test_editor_save_writes_no_internal_metadata(self, tmp_path: Path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

        assert loader.save_file_text("proofread/typos.toml", '[entries]\ncustom = ["fix"]\n')

        user_text = (user_dir / "proofread" / "typos.toml").read_text(encoding="utf-8")
        assert "__docwen_" not in user_text
        assert loader.get_file_dict("proofread/typos.toml") == {"entries": {"custom": ["fix"]}}
        assert loader.config.as_dict()["proofread"]["typos"] == {"entries": {"custom": ["fix"]}}

    def test_same_base_and_user_directory_writes_plain_toml(
        self,
        tmp_path: Path,
    ) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        config_dir = tmp_path / "configs"
        config_dir.mkdir()
        write_minimal_base_config_tree(config_dir)
        loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)

        assert loader.save_file_text(
            "proofread/typos.toml",
            '[entries]\ncustom = ["fix"]\n',
        )

        physical = (config_dir / "proofread" / "typos.toml").read_text(encoding="utf-8")
        editable = loader.get_file_text("proofread/typos.toml")
        assert physical == '[entries]\ncustom = ["fix"]\n'
        assert loader.config.as_dict()["proofread"]["typos"] == {"entries": {"custom": ["fix"]}}
        assert editable is not None
        assert "__docwen_" not in editable

    def test_removed_internal_marker_is_rejected_fail_closed(self, tmp_path: Path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        user_file = user_dir / "proofread" / "typos.toml"
        user_file.parent.mkdir(parents=True)
        user_file.write_text(
            '__docwen_replace_sections__ = ["entries"]\n[entries]\ncustom = ["fix"]\n',
            encoding="utf-8",
        )

        with pytest.raises(ValueError, match="Reserved internal keys"):
            ConfigLoader(base_dir=base_dir, user_dir=user_dir)


def test_config_registry_has_exactly_23_files() -> None:
    from docwen_runtime.config.registry import CONFIG_FILES

    assert len(CONFIG_FILES) == 23
    assert {spec.rel_path for spec in CONFIG_FILES} == {
        "gui.toml",
        "output.toml",
        "logger.toml",
        "conversion.toml",
        "export.toml",
        "other.toml",
        "document.toml",
        "text.toml",
        "layout.toml",
        "spreadsheet.toml",
        "image.toml",
        "link.toml",
        "software.toml",
        "optimize.toml",
        "field_processors.toml",
        "numbering/add.toml",
        "numbering/cleanup.toml",
        "proofread/engine.toml",
        "proofread/skip.toml",
        "proofread/pairs.toml",
        "proofread/symbol_map.toml",
        "proofread/typos.toml",
        "proofread/sensitive_words.toml",
    }


@pytest.mark.parametrize(
    ("dotted_key", "expected_rel_path"),
    [
        ("link.format.image_link_style", "link.toml"),
        ("text.remove_numbering", "text.toml"),
        ("export.to_md_image_extraction_mode", "export.toml"),
        ("document.keep_images", "document.toml"),
        ("numbering.add.settings.default_scheme", "numbering/add.toml"),
        ("proofread.engine.enable_typos_rule", "proofread/engine.toml"),
    ],
)
def test_registry_routes_dotted_keys_by_namespace(
    dotted_key: str,
    expected_rel_path: str,
) -> None:
    from docwen_runtime.config.registry import spec_for_dotted_key

    spec = spec_for_dotted_key(dotted_key)

    assert spec is not None
    assert spec.rel_path == expected_rel_path


def test_registry_returns_none_for_an_unknown_dotted_key() -> None:
    from docwen_runtime.config.registry import spec_for_dotted_key

    assert spec_for_dotted_key("unknown.setting") is None


def test_registry_groups_specs_by_group() -> None:
    from docwen_runtime.config.registry import specs_for_group

    assert {s.rel_path for s in specs_for_group("proofread")} == {
        "proofread/engine.toml",
        "proofread/skip.toml",
        "proofread/pairs.toml",
        "proofread/symbol_map.toml",
        "proofread/typos.toml",
        "proofread/sensitive_words.toml",
    }
    assert {s.rel_path for s in specs_for_group("text")} == {
        "field_processors.toml",
        "text.toml",
        "numbering/add.toml",
        "numbering/cleanup.toml",
    }
