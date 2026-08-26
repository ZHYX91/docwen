"""Permanent admission tests for deep runtime configuration contracts."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_CONFIGS = Path(__file__).resolve().parents[3] / "configs"


@pytest.mark.parametrize(
    ("rel_path", "invalid_text"),
    (
        ("gui.toml", '[history]\nrecent_files = ["ok.docx", 7]\n'),
        (
            "conversion.toml",
            "[ocr_output.blockquote_title_override_by_locale]\nzh_CN = 7\n",
        ),
        ("export.toml", 'to_md_image_extraction_mode = "directory"\n'),
        ("export.toml", 'to_md_image_extraction_mode = " BASE64 "\n'),
        ("export.toml", 'to_md_ocr_placement_mode = "sidecar"\n'),
        ("link.toml", '[format]\nmd_file_link_style = "markdown_embed"\n'),
        ("link.toml", '[non_embed_links]\nwiki_mode = "render"\n'),
        ("link.toml", "[embedding]\nmax_depth = 0\n"),
        ("link.toml", '[path_resolution]\nsearch_dirs = ["assets", 7]\n'),
        (
            "software.toml",
            '[default_priority]\nword_processors = ["libreoffice", 7]\n',
        ),
        ("optimize.toml", '[types.custom]\nenabled = "yes"\n'),
        (
            "field_processors.toml",
            "[processors.custom]\nmodule = 7\nenabled = true\n",
        ),
        (
            "field_processors.toml",
            '[processors.custom]\nmodule = "example.custom"\nlocales = []\nenabled = true\n',
        ),
        ("numbering/add.toml", '[settings]\norder = ["custom", "custom"]\n'),
        (
            "numbering/add.toml",
            '[schemes.custom]\nenabled = true\nlocales = ["*"]\n[schemes.custom.level_1]\nformat = "{1.unknown}"\n',
        ),
        (
            "numbering/add.toml",
            '[schemes.custom]\nenabled = true\nlocales = []\n[schemes.custom.level_1]\nformat = "{1.arabic_half}"\n',
        ),
        ("numbering/add.toml", '[schemes.custom]\nenabled = true\nlocales = ["*"]\n'),
        (
            "numbering/cleanup.toml",
            '[[rules]]\nid = "custom"\npattern = "("\nlevel = 1\n',
        ),
        (
            "numbering/cleanup.toml",
            '[[rules]]\nid = "custom"\npattern = "^x"\nlevel = 6\n',
        ),
        ("proofread/pairs.toml", 'items = [["(", ")", "extra"]]\n'),
        ("proofread/pairs.toml", 'items = [["(", ")"], ["(", ")"]]\n'),
        ("proofread/pairs.toml", 'items = [["(", ")"], ["[", ")"]]\n'),
        ("proofread/pairs.toml", 'items = [{ source = "(", close = ")" }]\n'),
        ("proofread/symbol_map.toml", "[entries]\nzero = [0]\n"),
        ("proofread/typos.toml", '[entries]\ncorrect = ["wrong", 7]\n'),
        ("proofread/sensitive_words.toml", "[entries]\nword = [7]\n"),
    ),
)
def test_deep_semantic_invalid_user_override_is_quarantined(
    tmp_path: Path,
    rel_path: str,
    invalid_text: str,
) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    user_path = user_dir / rel_path
    user_path.parent.mkdir(parents=True, exist_ok=True)
    invalid_bytes = invalid_text.encode("utf-8")
    user_path.write_bytes(invalid_bytes)

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)

    assert loader.config_state_trusted is True
    assert not user_path.exists()
    backups = list(user_path.parent.glob(f"{user_path.name}.bak_schema_failed_*"))
    assert len(backups) == 1
    assert backups[0].read_bytes() == invalid_bytes


@pytest.mark.parametrize(
    "runtime_overrides",
    (
        {"link": {"embedding": {"max_depth": 21}}},
        {"export": {"to_md_ocr_placement_mode": "sidecar"}},
        {"link": {"format": {"md_file_link_style": "markdown_embed"}}},
        {"software": {"default_priority": {"word_processors": [7]}}},
        {
            "numbering": {
                "cleanup": {
                    "rules": [{"id": "bad", "pattern": "^x", "level": "one"}],
                },
            },
        },
        {"proofread": {"pairs": {"items": [["only-one"]]}}},
    ),
)
def test_deep_semantic_invalid_runtime_override_fails_closed(
    tmp_path: Path,
    runtime_overrides: dict[str, object],
) -> None:
    from docwen_runtime.config.loader import ConfigLoader
    from docwen_runtime.config.validation import ConfigSemanticError

    with pytest.raises(ConfigSemanticError):
        ConfigLoader(
            base_dir=PROJECT_CONFIGS,
            user_dir=tmp_path / "user",
            runtime_overrides=runtime_overrides,
        )


def test_link_and_export_values_reach_consumers_only_after_validation(tmp_path: Path) -> None:
    from docwen_core.export_semantics import LinkRuntimeConfig, MarkdownExportSemantics
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    user_dir.mkdir()
    (user_dir / "link.toml").write_text(
        """\
[format]
image_link_style = "markdown_link"
md_file_link_style = "markdown_link"
[non_embed_links]
wiki_mode = "extract_text"
markdown_mode = "remove"
[embed_links]
wiki_image_mode = "keep"
markdown_image_mode = "extract_text"
md_file_mode = "remove"
[embedding]
max_depth = 20
[path_resolution]
search_dirs = []
[error_handling]
file_not_found = "ignore"
circular_reference = "keep"
max_depth_reached = "placeholder"
""",
        encoding="utf-8",
    )
    (user_dir / "export.toml").write_text(
        'to_md_image_extraction_mode = "base64"\nto_md_ocr_placement_mode = "image_md"\n',
        encoding="utf-8",
    )

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
    snapshot = loader.config.as_dict()
    link = LinkRuntimeConfig.from_config(snapshot["link"])
    export = MarkdownExportSemantics.from_config_snapshot(snapshot)

    assert snapshot["link"]["format"] == {
        "image_link_style": "markdown_link",
        "md_file_link_style": "markdown_link",
    }
    assert link.max_depth == 20
    assert link.search_dirs == ()
    assert link.non_embed_wiki_mode == "extract_text"
    assert link.non_embed_markdown_mode == "remove"
    assert link.embed_wiki_image_mode == "keep"
    assert link.embed_markdown_image_mode == "extract_text"
    assert link.embed_md_file_mode == "remove"
    assert link.file_not_found_mode == "ignore"
    assert link.circular_reference_mode == "keep"
    assert link.max_depth_reached_mode == "placeholder"
    assert export.image_link_style == "markdown_link"
    assert export.md_file_link_style == "markdown_link"
    assert export.image_extraction_mode == "base64"
    assert export.ocr_placement_mode == "image_md"
    assert not list(user_dir.glob("*.bak_*"))


@pytest.mark.parametrize(
    ("record", "expected"),
    (
        ('{ source = "(", target = ")" }', ["(", ")"]),
        ('{ open = "[", close = "]" }', ["[", "]"]),
    ),
)
def test_supported_proofread_pair_records_are_admitted_and_canonicalized(
    tmp_path: Path,
    record: str,
    expected: list[str],
) -> None:
    """Legacy GUI/runtime record forms remain active without rewriting user bytes."""
    from docwen_runtime.config import build_proofread_rules
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    user_file = user_dir / "proofread" / "pairs.toml"
    user_file.parent.mkdir(parents=True)
    original = f"items = [{record}]\n".encode()
    user_file.write_bytes(original)

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
    snapshot = loader.config.as_dict()

    assert loader.config_state_trusted is True
    assert snapshot["proofread"]["pairs"]["items"] == [expected]
    assert build_proofread_rules(snapshot).symbol_pairs == (tuple(expected),)
    assert user_file.read_bytes() == original
    assert not list(user_file.parent.glob("pairs.toml.bak_*"))


def test_identifier_lists_are_trimmed_before_runtime_consumption(tmp_path: Path) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    user_dir.mkdir()
    (user_dir / "software.toml").write_text(
        '[default_priority]\nword_processors = [" libreoffice "]\n',
        encoding="utf-8",
    )
    (user_dir / "optimize.toml").write_text(
        '[settings]\norder = [" gongwen "]\n',
        encoding="utf-8",
    )

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)

    assert loader.config.software.default_priority.word_processors == ["libreoffice"]
    assert loader.config.optimize.settings.order == ["gongwen"]


@pytest.mark.parametrize(
    "invalid_rule",
    (
        {"id": "bad-level", "pattern": "^x", "level": "one"},
        {"id": "bad-regex", "pattern": "(", "level": 1},
    ),
)
def test_invalid_cleanup_rule_cannot_reach_request_compilation(
    tmp_path: Path,
    invalid_rule: dict[str, object],
) -> None:
    import tomlkit

    from docwen_runtime.config import build_heading_cleanup_rules
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    user_file = user_dir / "numbering" / "cleanup.toml"
    user_file.parent.mkdir(parents=True)
    user_file.write_text(tomlkit.dumps({"rules": [invalid_rule]}), encoding="utf-8")

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
    compiled = build_heading_cleanup_rules(loader.config.as_dict())

    assert compiled
    assert all(rule_id not in {"bad-level", "bad-regex"} for rule_id, _pattern, _level in compiled)
    assert len(list(user_file.parent.glob("cleanup.toml.bak_schema_failed_*"))) == 1


def test_deep_semantic_raw_editor_failure_restores_exact_preimage(tmp_path: Path) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    user_dir.mkdir()
    user_file = user_dir / "link.toml"
    preimage = b"[embedding]\nmax_depth = 7\n"
    user_file.write_bytes(preimage)
    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)

    assert loader.save_file_text("link.toml", "[embedding]\nmax_depth = 0\n") is False
    assert user_file.read_bytes() == preimage
    assert loader.config.link.embedding.max_depth == 7
    assert loader.config_state_trusted is True
