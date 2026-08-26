"""Tests for application preconversion chain decisions."""

import pytest

pytestmark = pytest.mark.unit


class TestResolveChain:
    """Preconversion should only cover formats that need an external hub step."""

    def test_document_non_hub_formats_preconvert_to_docx(self) -> None:
        from docwen_application.preconversion.chain_resolver import resolve_chain

        assert resolve_chain("doc", "md") == ["docx", "md"]
        assert resolve_chain("odt", "md") == ["docx", "md"]
        assert resolve_chain("rtf", "markdown") == ["docx", "markdown"]

    @pytest.mark.parametrize("source", ["doc", "odt", "rtf", "wps"])
    def test_document_hub_target_uses_direct_plugin_route(self, source: str) -> None:
        from docwen_application.preconversion.chain_resolver import resolve_chain

        assert resolve_chain(source, "docx") == ["docx"]

    @pytest.mark.parametrize("source", ["doc", "odt", "rtf", "wps"])
    def test_document_hub_action_preconverts_before_runtime_dispatch(self, source: str) -> None:
        from docwen_application.preconversion.chain_resolver import resolve_chain

        assert resolve_chain(source, "docx", action_name="validate") == ["docx", "docx"]

    @pytest.mark.parametrize(
        ("source", "target"),
        [
            ("doc", "odt"),
            ("doc", "rtf"),
            ("doc", "wps"),
            ("odt", "doc"),
            ("rtf", "wps"),
            ("wps", "pdf"),
        ],
    )
    def test_document_non_markdown_targets_use_plugin_owned_routes(self, source: str, target: str) -> None:
        from docwen_application.preconversion.chain_resolver import resolve_chain

        assert resolve_chain(source, target) == [target]

    @pytest.mark.parametrize("source", ["csv", "tsv"])
    @pytest.mark.parametrize("target", ["md", "xlsx"])
    def test_delimited_spreadsheets_use_direct_plugin_routes(self, source: str, target: str) -> None:
        from docwen_application.preconversion.chain_resolver import resolve_chain

        assert resolve_chain(source, target) == [target]

    @pytest.mark.parametrize("source", ["xls", "ods", "et"])
    @pytest.mark.parametrize("target", ["md", "csv", "pdf", "xlsx"])
    def test_binary_spreadsheets_use_plugin_owned_routes(self, source: str, target: str) -> None:
        from docwen_application.preconversion.chain_resolver import resolve_chain

        assert resolve_chain(source, target) == [target]

    @pytest.mark.parametrize("source", ["", "unknown"])
    def test_non_concrete_source_is_rejected(self, source: str) -> None:
        from docwen_application.preconversion.chain_resolver import resolve_chain

        with pytest.raises(ValueError, match="concrete admitted format"):
            resolve_chain(source, "md")
