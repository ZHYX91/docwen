"""Focused tests split from test_convert_typed_options.py."""

from __future__ import annotations

from ._convert_typed_options_support import (
    _fake_convert_args,
    pytest,
)

pytestmark = pytest.mark.unit


class TestBuildConvertOptionsNumbering:
    """``build_execution_options()`` produces normalized numbering keys."""

    def test_convert_with_numbering(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "clean_numbering": "remove",
                "add_numbering": "gongwen_standard",
            }
        )
        opts = build_execution_options(args)
        assert opts["remove_numbering"] is True
        assert opts["add_numbering"] is True
        assert opts["numbering_scheme"] == "gongwen_standard"
        # Old flat keys must NOT leak.
        assert "clean_numbering" not in opts
        assert "add_numbering_scheme" not in opts

    def test_numbering_action_options(self) -> None:
        """A normalized Markdown-numbering action uses the shared option builder."""
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "action": "md_numbering",
                "clean_numbering": "keep",
                "add_numbering": "cn_article",
            }
        )
        opts = build_execution_options(args)
        assert opts["remove_numbering"] is False
        assert opts["add_numbering"] is True
        assert opts["numbering_scheme"] == "cn_article"

    def test_no_numbering_args_no_keys(self) -> None:
        """When no numbering args are given, no numbering keys appear."""
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args()
        opts = build_execution_options(args)
        assert "remove_numbering" not in opts
        assert "add_numbering" not in opts
        assert "numbering_scheme" not in opts
        assert "heading_numbering_render_mode" not in opts

    def test_render_mode_in_build_options(self) -> None:
        """``--heading-numbering-render-mode word_native`` reaches build_execution_options."""
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "to": "docx",
                "heading_numbering_render_mode": "word_native",
                "add_numbering": "gongwen_standard",
            }
        )
        opts = build_execution_options(args)
        assert opts["heading_numbering_render_mode"] == "word_native"
        assert opts["add_numbering"] is True

    def test_render_mode_preserved_for_document_target_when_no_numbering(self) -> None:
        """render_mode without add/clean is still carried for MD→document routes."""
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "to": "docx",
                "heading_numbering_render_mode": "text",
            }
        )
        opts = build_execution_options(args)
        assert opts["heading_numbering_render_mode"] == "text"

    def test_render_mode_applicability_is_owned_by_the_route(self) -> None:
        """The value builder normalizes; the canonical route decides support."""
        from docwen_cli.commands.execution_options import build_execution_options
        from docwen_cli.commands.execution_request import project_route_options

        args = _fake_convert_args(
            {
                "action": "md_numbering",
                "to": "md",
                "add_numbering": "gongwen_standard",
                "heading_numbering_render_mode": "word_native",
            }
        )
        opts = build_execution_options(args)
        assert opts["add_numbering"] is True
        assert opts["numbering_scheme"] == "gongwen_standard"
        assert opts["heading_numbering_render_mode"] == "word_native"
        with pytest.raises(ValueError, match="heading_numbering_render_mode"):
            project_route_options(
                opts,
                route_id="markdown-numbering",
                route_options=("remove_numbering", "add_numbering", "numbering_scheme"),
            )

    def test_heading_merge_mode_is_normalized_then_projected_by_route_contract(self) -> None:
        """H-D2-026: normalize the CLI value, then fail closed on unsupported routes."""
        from docwen_cli.commands.execution_options import build_execution_options
        from docwen_cli.commands.execution_request import project_route_options

        options = build_execution_options(
            _fake_convert_args(
                {
                    "to": "docx",
                    "heading_merge_mode": "\tAlWaYs\n",
                }
            )
        )

        assert options == {"heading_merge_mode": "always"}
        assert project_route_options(
            options,
            route_id="markdown-to-docx",
            route_options=("heading_merge_mode",),
        ) == {"heading_merge_mode": "always"}
        with pytest.raises(ValueError, match=r"markdown-to-xlsx.*heading_merge_mode"):
            project_route_options(
                options,
                route_id="markdown-to-xlsx",
                route_options=(),
            )


class TestPolicy02SpreadsheetPasswordOptions:
    def test_masked_prompt_builds_request_scoped_password_and_consent(self, monkeypatch: pytest.MonkeyPatch) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        prompts: list[str] = []

        def fake_getpass(prompt: str) -> str:
            prompts.append(prompt)
            return "test"

        monkeypatch.setattr("getpass.getpass", fake_getpass)
        args = _fake_convert_args(
            {
                "to": "ods",
                "files": ["/test/protected.xlsx"],
                "spreadsheet_password_prompt": True,
                "allow_spreadsheet_protection_loss": True,
            }
        )

        options = build_execution_options(
            args,
            route_options=("spreadsheet_password", "allow_spreadsheet_protection_loss"),
        )

        assert prompts
        assert options == {
            "spreadsheet_password": "test",
            "allow_spreadsheet_protection_loss": True,
        }

    def test_password_prompt_is_rejected_for_multi_input_reuse(self) -> None:
        from docwen_cli.commands.execution_options import validate_execution_options

        args = _fake_convert_args(
            {
                "to": "ods",
                "files": ["/test/one.xlsx", "/test/two.xlsx"],
                "spreadsheet_password_prompt": True,
            }
        )

        with pytest.raises(ValueError, match="exactly one input"):
            validate_execution_options(args)

    def test_xlsx_suffix_cannot_enable_password_prompt_for_non_xlsx_content(self) -> None:
        """A route without password support rejects before prompting."""
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "to": "ods",
                "files": ["/test/looks-like.xlsx"],
                "spreadsheet_password_prompt": True,
            }
        )

        with pytest.raises(ValueError, match="spreadsheet_password"):
            build_execution_options(args, route_options=())

    def test_wrong_xlsx_suffix_still_enables_xlsx_options_from_content(self, monkeypatch) -> None:
        """Concrete XLSX content is authoritative even when the filename says ODS."""
        from docwen_cli.commands.execution_options import build_execution_options, validate_execution_options

        prompts: list[str] = []
        monkeypatch.setattr("getpass.getpass", lambda prompt: prompts.append(prompt) or "secret")
        args = _fake_convert_args(
            {
                "to": "ods",
                "files": ["/test/actually-xlsx.ods"],
                "spreadsheet_password_prompt": True,
                "allow_spreadsheet_protection_loss": True,
            }
        )

        validate_execution_options(args)
        options = build_execution_options(
            args,
            route_options=("spreadsheet_password", "allow_spreadsheet_protection_loss"),
        )

        assert prompts
        assert options == {
            "spreadsheet_password": "secret",
            "allow_spreadsheet_protection_loss": True,
        }

    def test_template_accepts_detected_markdown_despite_docx_suffix(self) -> None:
        """Template eligibility follows admitted Markdown content, not ``.docx``."""
        from docwen_cli.commands.execution_options import validate_execution_options

        args = _fake_convert_args(
            {
                "to": "docx",
                "files": ["/test/actually-markdown.docx"],
                "template": f"template.docx.{'a' * 64}",
            }
        )

        validate_execution_options(args)

    def test_template_rejects_display_names_paths_and_filenames(self) -> None:
        from docwen_cli.commands.execution_options import validate_execution_options

        for identifier in ("official", "template.docx", "/templates/template.docx"):
            args = _fake_convert_args({"to": "docx", "template": identifier})
            with pytest.raises(ValueError, match="exact canonical ID"):
                validate_execution_options(args)

    def test_template_rejects_detected_docx_despite_markdown_suffix(self) -> None:
        """A route that does not declare templates rejects the explicit key."""
        from docwen_cli.commands.execution_options import build_execution_options
        from docwen_cli.commands.execution_request import project_route_options

        args = _fake_convert_args(
            {
                "to": "docx",
                "files": ["/test/actually-docx.md"],
                "template": f"template.docx.{'a' * 64}",
            }
        )

        options = build_execution_options(args)
        with pytest.raises(ValueError, match="template_name"):
            project_route_options(
                options,
                route_id="docx-to-docx",
                route_options=(),
            )

    def test_spreadsheet_flag_applicability_is_owned_by_the_route(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options
        from docwen_cli.commands.execution_request import project_route_options

        args = _fake_convert_args(
            {
                "to": "pdf",
                "spreadsheet_password_prompt": False,
                "allow_spreadsheet_protection_loss": True,
            }
        )

        options = build_execution_options(args)

        assert options == {"allow_spreadsheet_protection_loss": True}
        with pytest.raises(ValueError, match="allow_spreadsheet_protection_loss"):
            project_route_options(
                options,
                route_id="unrelated-pdf-route",
                route_options=(),
            )

    def test_dry_run_redaction_never_projects_password(self) -> None:
        from docwen_cli.commands.execution_request import redacted_options

        assert redacted_options(
            {
                "spreadsheet_password": "pw-SECRET-879",
                "allow_spreadsheet_protection_loss": True,
            }
        ) == {
            "spreadsheet_password": "<redacted>",
            "allow_spreadsheet_protection_loss": True,
        }


class TestBuildConvertOptionsProofread:
    """``build_execution_options()`` produces normalized proofread keys."""

    def test_check_all_expands(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args({"check": ["all"]})
        opts = build_execution_options(args)
        assert opts["enable_symbol_pairing"] is True
        assert opts["enable_symbol_correction"] is True
        assert opts["enable_typos_rule"] is True
        assert opts["enable_sensitive_word"] is True
        # Old flat key must NOT leak.
        assert "proofread_checks" not in opts

    def test_check_none_all_false(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args({"check": ["none"]})
        opts = build_execution_options(args)
        assert opts["enable_symbol_pairing"] is False
        assert opts["enable_symbol_correction"] is False
        assert opts["enable_typos_rule"] is False
        assert opts["enable_sensitive_word"] is False

    def test_check_punct_typo(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args({"check": ["punct", "typo"]})
        opts = build_execution_options(args)
        assert opts["enable_symbol_pairing"] is True
        assert opts["enable_symbol_correction"] is False
        assert opts["enable_typos_rule"] is True
        assert opts["enable_sensitive_word"] is False

    def test_no_checks_no_keys(self) -> None:
        """When no --check given, no enable_* keys appear (service defaults apply)."""
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args({"check": []})
        opts = build_execution_options(args)
        assert "enable_symbol_pairing" not in opts
        assert "proofread_checks" not in opts

    def test_validation_action_options(self) -> None:
        """A normalized validation action applies proofread normalization."""
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "action": "validate",
                "check": ["typo", "sensitive"],
            }
        )
        opts = build_execution_options(args)
        assert opts["enable_typos_rule"] is True
        assert opts["enable_sensitive_word"] is True
        assert opts["enable_symbol_pairing"] is False
        assert opts["enable_symbol_correction"] is False
