from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]


def _read(relative: str) -> str:
    return (ROOT / relative).read_text(encoding="utf-8")


def test_presentation_markdown_policy_is_request_owned_and_immutable() -> None:
    policy = _read("packages/plugins/presentation/src/docwen_plugin_presentation/pptx_md/request_policy.py")

    assert "@dataclass(frozen=True, slots=True)" in policy
    assert "class PresentationMarkdownRequestPolicy" in policy
    assert "MarkdownExportSemantics.from_config(" in policy
    assert 'getattr(context, "ocr_blockquote_title", None)' in policy
    assert "runtime_policy_transaction" not in policy
    assert "get_markdown_export_semantics" not in policy
    assert "get_ocr_blockquote_title" not in policy


def test_pptx_route_builds_one_policy_and_passes_every_export_consumer() -> None:
    converter = _read("packages/plugins/presentation/src/docwen_plugin_presentation/pptx_md/converter.py")
    parse = converter[
        converter.index("    def _parse_pptx(") : converter.index("    @staticmethod\n    def _is_hidden_slide")
    ]
    slide = converter[converter.index("    def _process_slide(") :]

    assert parse.count("build_presentation_markdown_request_policy(") == 1
    assert "request_policy=request_policy" in parse
    assert "policy.export.image_extraction_mode" in slide
    assert "policy.export.ocr_placement_mode" in slide
    assert "policy.export.image_link_style" in slide
    assert "policy.export.md_file_link_style" in slide
    assert "export_semantics=policy.export" in slide
    assert "policy.ocr_blockquote_title" in slide
    for legacy_getter in (
        "get_markdown_export_modes",
        "get_markdown_asset_link_semantics",
        "get_ocr_blockquote_title",
    ):
        assert legacy_getter not in converter


def test_legacy_ppt_captures_policy_before_hub_proxy_delegation() -> None:
    bridge = _read("packages/plugins/presentation/src/docwen_plugin_presentation/pptx_md/ppt_converter.py")
    policy_pos = bridge.index("request_policy = build_presentation_markdown_request_policy(")
    conversion_pos = bridge.index("result = convert_with_backend_priority(")
    proxy_pos = bridge.index("proxy_context = HubConversionContext(")

    assert policy_pos < conversion_pos < proxy_pos
    assert "request_policy=request_policy" in bridge[proxy_pos:]
    assert "config_snapshot=dict(context.request.config_snapshot)" in bridge
