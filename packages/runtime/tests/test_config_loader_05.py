"""Focused tests split from test_config_loader.py."""

from __future__ import annotations

from ._config_loader_support import (
    PROJECT_CONFIGS,
    Any,
    Path,
    pytest,
    write_minimal_base_config_tree,
)

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    "operation",
    (
        "set_value",
        "set_values",
        "write_file_content",
        "save_file_text",
        "update_file_sections",
        "update_file_document",
    ),
)
def test_semantic_invalid_persistence_restores_exact_preimage(
    tmp_path: Path,
    operation: str,
) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    user_dir.mkdir()
    user_file = user_dir / "output.toml"
    preimage = b'[directory]\nmode = "custom"\ncustom_path = "kept"\n'
    user_file.write_bytes(preimage)
    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)

    if operation == "set_value":
        result = loader.set_value("output.directory.mode", "not-a-mode")
    elif operation == "set_values":
        result = loader.set_values({"output.directory.mode": "not-a-mode"})
    elif operation == "write_file_content":
        result = loader.write_file_content("output.toml", {"directory": {"mode": "not-a-mode"}})
    elif operation == "save_file_text":
        result = loader.save_file_text("output.toml", '[directory]\nmode = "not-a-mode"\n')
    elif operation == "update_file_sections":
        result = loader.update_file_sections("output.toml", {"directory": {"mode": "not-a-mode"}})
    else:

        def _make_invalid(doc: Any) -> None:
            doc["directory"]["mode"] = "not-a-mode"

        result = loader.update_file_document("output.toml", _make_invalid)

    assert result is False
    assert user_file.read_bytes() == preimage
    assert loader.config.output.directory.mode == "custom"
    assert loader.config_state_trusted is True


def test_invalid_shipped_config_fails_closed_without_recovery(tmp_path: Path) -> None:
    from docwen_runtime.config.loader import ConfigLoader
    from docwen_runtime.config.validation import ConfigSemanticError

    base_dir = tmp_path / "base"
    user_dir = tmp_path / "user"
    base_dir.mkdir()
    write_minimal_base_config_tree(base_dir)
    invalid_bytes = b'level = "verbose"\n'
    (base_dir / "logger.toml").write_bytes(invalid_bytes)

    with pytest.raises(ConfigSemanticError, match=r"logger\.level"):
        ConfigLoader(base_dir=base_dir, user_dir=user_dir)

    assert (base_dir / "logger.toml").read_bytes() == invalid_bytes
    assert not list(base_dir.glob("logger.toml.bak_*"))


def test_invalid_runtime_override_fails_closed_without_touching_user_files(tmp_path: Path) -> None:
    from docwen_runtime.config.loader import ConfigLoader
    from docwen_runtime.config.validation import ConfigSemanticError

    user_dir = tmp_path / "user"
    with pytest.raises(ConfigSemanticError, match=r"output\.directory\.mode"):
        ConfigLoader(
            base_dir=PROJECT_CONFIGS,
            user_dir=user_dir,
            runtime_overrides={"output": {"directory": {"mode": "not-a-mode"}}},
        )

    assert not user_dir.exists()


def test_runtime_overrides_override_both_base_and_user(tmp_path) -> None:
    base_dir = tmp_path / "base"
    user_dir = tmp_path / "user"
    base_dir.mkdir()
    user_dir.mkdir()
    write_minimal_base_config_tree(base_dir)
    (base_dir / "link.toml").write_text('[format]\nimage_link_style = "wiki_embed"\n', encoding="utf-8")
    (user_dir / "link.toml").parent.mkdir(parents=True, exist_ok=True)
    (user_dir / "link.toml").write_text('[format]\nimage_link_style = "markdown_link"\n', encoding="utf-8")

    from docwen_runtime.config.loader import ConfigLoader

    loader = ConfigLoader(
        base_dir=base_dir,
        user_dir=user_dir,
        runtime_overrides={"link": {"format": {"image_link_style": "wiki_link"}}},
    )

    assert loader.config.link.format.image_link_style == "wiki_link"


def test_user_file_is_not_backfilled_with_base_defaults(tmp_path) -> None:
    base_dir = tmp_path / "base"
    user_dir = tmp_path / "user"
    base_dir.mkdir()
    user_dir.mkdir()
    write_minimal_base_config_tree(base_dir)
    (base_dir / "numbering" / "add.toml").write_text(
        '[settings]\ndefault_scheme = "gongwen_standard"\n'
        '[schemes.gongwen_standard]\nname = "公文"\nis_system = true\n'
        '[schemes.gongwen_standard.level_1]\nformat = "{1.arabic_half} "\n',
        encoding="utf-8",
    )
    (user_dir / "numbering").mkdir(parents=True)
    (user_dir / "numbering" / "add.toml").write_text(
        '[schemes.my_custom]\nname = "自定义"\nis_system = false\n'
        '[schemes.my_custom.level_1]\nformat = "{1.arabic_half} "\n',
        encoding="utf-8",
    )

    from docwen_runtime.config.loader import ConfigLoader

    loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)

    user_text = (user_dir / "numbering" / "add.toml").read_text(encoding="utf-8")
    assert "gongwen_standard" not in user_text
    assert loader.config.numbering.add.schemes.gongwen_standard.is_system is True
    assert loader.config.numbering.add.schemes.my_custom.is_system is False


def test_get_base_file_dict_returns_raw_base(tmp_path) -> None:
    base_dir = tmp_path / "base"
    user_dir = tmp_path / "user"
    base_dir.mkdir()
    user_dir.mkdir()
    write_minimal_base_config_tree(base_dir)
    (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")

    from docwen_runtime.config.loader import ConfigLoader

    loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
    raw = loader.get_base_file_dict("gui.toml")
    assert raw.get("window", {}).get("default_mode") == "single"
