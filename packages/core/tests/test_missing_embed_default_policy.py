"""User-facing missing-embed default keeps authored source spelling."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import process_markdown_links

pytestmark = pytest.mark.unit


def test_missing_embeds_default_to_keep_without_file_not_found_text(tmp_path: Path) -> None:
    source = tmp_path / "original.md"
    content = "![[missing.md]]\n\n![插图](missing.png)\n"
    source.write_text(content, encoding="utf-8")

    policy = LinkRuntimeConfig()
    assert policy.file_not_found_mode == "keep"
    assert LinkRuntimeConfig.from_config({}).file_not_found_mode == "keep"

    projected = process_markdown_links(
        content,
        str(source),
        link_config=policy,
        target_format="md",
    )
    assert "![[missing.md]]" in projected
    assert "![插图](missing.png)" in projected
    assert "File not found" not in projected


def test_user_can_explicitly_request_visible_placeholder(tmp_path: Path) -> None:
    source = tmp_path / "original.md"
    source.write_text("![[missing.md]]\n", encoding="utf-8")
    config = LinkRuntimeConfig.from_config({"error_handling": {"file_not_found": "placeholder"}})
    assert config.file_not_found_mode == "placeholder"
    projected = process_markdown_links("![[missing.md]]", str(source), link_config=config, target_format="md")
    assert "File not found" in projected


def test_shipped_link_config_and_runtime_default_agree() -> None:
    import tomllib

    root = Path(__file__).resolve().parents[3]
    data = tomllib.loads((root / "configs/link.toml").read_text(encoding="utf-8"))
    assert data["error_handling"]["file_not_found"] == "keep"
    assert LinkRuntimeConfig.from_config(data).file_not_found_mode == "keep"


def test_missing_embed_notification_reports_only_attempted_lookup(tmp_path: Path) -> None:
    source = tmp_path / "original.md"
    content = (
        "![[missing.md]]\n\n"
        "![photo](missing.png)\n\n"
        "`![literal](not-a-lookup.png)`\n\n"
        "```md\n![[also-not-a-lookup.md]]\n```\n"
    )
    source.write_text(content, encoding="utf-8")
    events: list[int] = []
    projected = process_markdown_links(
        content,
        str(source),
        link_config=LinkRuntimeConfig(),
        target_format="md",
        on_missing_embed=lambda: events.append(1),
    )
    assert len(events) == 2
    assert "File not found" not in projected
    assert "![photo](missing.png)" in projected
    assert "![[missing.md]]" in projected
