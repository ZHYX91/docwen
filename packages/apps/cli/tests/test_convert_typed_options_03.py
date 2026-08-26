"""Focused tests split from test_convert_typed_options.py."""

from __future__ import annotations

from ._convert_typed_options_support import (
    _fake_convert_args,
    pytest,
)

pytestmark = pytest.mark.unit


class TestBuildConvertOptionsPages:
    """``build_execution_options()`` parses pages to list[int]."""

    def test_split_pdf_pages_parsed(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "action": "split_pdf",
                "to": "",
                "pages": "1-3,5",
            }
        )
        opts = build_execution_options(args)
        assert opts["pages"] == [1, 2, 3, 5]
        assert "render_dpi" not in opts

    def test_public_split_pdf_projects_exact_runtime_route_options(self, tmp_path) -> None:
        """The public split command must reach its route without phantom render options."""
        from docwen_cli.commands.execution_options import build_execution_options
        from docwen_cli.commands.execution_request import project_route_options
        from docwen_cli.commands.execution_v3 import _prepare_args
        from docwen_cli.main import _build_parser
        from docwen_plugin_layout import LayoutPlugin

        args = _build_parser().parse_args(
            [
                "split",
                "pdf",
                str(tmp_path / "source.pdf"),
                "--pages",
                "1-3,5",
                "--output-dir",
                str(tmp_path / "out"),
            ]
        )
        _prepare_args(args)
        route = next(
            route
            for route in LayoutPlugin().manifest.routes
            if (route.source_format, route.target_format, route.action_name) == ("pdf", "pdf", "split_pdf")
        )
        route_options = tuple(route.options_schema["properties"])

        options = build_execution_options(args, route_options=route_options)
        projected = project_route_options(
            options,
            route_id="layout:pdf:pdf:split_pdf",
            route_options=route_options,
        )

        assert args.command_path == "split pdf"
        assert projected == {"pages": [1, 2, 3, 5]}

    def test_public_split_pdf_rejects_retired_unused_dpi_option(self, tmp_path) -> None:
        """PDF splitting is byte-preserving and therefore has no render DPI."""
        from docwen_cli.main import _build_parser
        from docwen_cli.parser import CliUsageError

        with pytest.raises(CliUsageError, match="unrecognized arguments: --dpi 300"):
            _build_parser().parse_args(
                [
                    "split",
                    "pdf",
                    str(tmp_path / "source.pdf"),
                    "--pages",
                    "1",
                    "--output-dir",
                    str(tmp_path / "out"),
                    "--dpi",
                    "300",
                ]
            )

    def test_standard_layout_render_dpi_is_not_an_action_option(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "to": "png",
                "dpi": 300,
            }
        )
        opts = build_execution_options(args)
        assert opts["render_dpi"] == 300

    def test_merge_tables_mode_uses_plugin_option_name(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "action": "merge_tables",
                "mode": "col",
            }
        )

        opts = build_execution_options(args)

        assert opts["merge_mode"] == "col"
        assert "mode" not in opts

    def test_merge_mode_is_rejected_by_routes_that_do_not_declare_it(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options
        from docwen_cli.commands.execution_request import project_route_options

        args = _fake_convert_args(
            {
                "action": "split_pdf",
                "to": "",
                "pages": "1-2",
                "mode": "col",
            }
        )

        opts = build_execution_options(args)

        assert opts["pages"] == [1, 2]
        assert opts["merge_mode"] == "col"
        with pytest.raises(ValueError, match="merge_mode"):
            project_route_options(
                opts,
                route_id="split-pdf",
                route_options=("pages",),
            )

    def test_merge_mode_is_not_silently_dropped_for_standard_conversion(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "to": "md",
                "mode": "col",
            }
        )

        opts = build_execution_options(args)

        assert opts["merge_mode"] == "col"

    def test_merge_images_to_tiff_keeps_public_keep_alpha_flag(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "action": "merge_images_to_tiff",
                "to": "",
                "keep_alpha": True,
            }
        )

        opts = build_execution_options(args)

        assert opts["keep_alpha"] is True

    def test_keep_alpha_is_not_silently_dropped_from_explicit_input(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        split_args = _fake_convert_args(
            {
                "action": "split_pdf",
                "to": "",
                "pages": "1",
                "keep_alpha": True,
            }
        )
        md_args = _fake_convert_args(
            {
                "to": "md",
                "keep_alpha": True,
            }
        )

        split_opts = build_execution_options(split_args)
        md_opts = build_execution_options(md_args)

        assert split_opts["pages"] == [1]
        assert split_opts["keep_alpha"] is True
        assert md_opts["keep_alpha"] is True

    def test_split_pdf_single_page(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "action": "split_pdf",
                "to": "",
                "pages": "7",
            }
        )
        opts = build_execution_options(args)
        assert opts["pages"] == [7]

    def test_no_pages_no_key(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args()
        opts = build_execution_options(args)
        assert "pages" not in opts

    def test_descending_range(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "action": "split_pdf",
                "to": "",
                "pages": "10-5",
            }
        )
        opts = build_execution_options(args)
        assert opts["pages"] == [5, 6, 7, 8, 9, 10]


class TestBuildConvertOptionsOutputPolicy:
    """``--output`` is carried by OutputPolicy, not plugin options."""

    def test_output_directory_does_not_leak_as_plugin_output_path(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args({"output": "exports"})

        opts = build_execution_options(args)

        assert "output_path" not in opts
        assert "output_dir" not in opts


class TestCliPluginContractAlignment:
    """End-to-end user-path tests: CLI options → plugin-consumable keys.

    These tests verify that ``build_execution_options()`` produces keys
    and value types that the downstream plugins actually consume.
    """

    # ── Numbering plugin contract ────────────────────────────────

    def test_numbering_keys_match_plugin_expectation(self) -> None:
        """Keys match numbering/converter.py:56-58 contract."""
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "clean_numbering": "remove",
                "add_numbering": "gongwen_standard",
            }
        )
        opts = build_execution_options(args)

        # Simulate what numbering/converter.py:56-58 does:
        remove_num = opts.get("remove_numbering", True)
        add_num = opts.get("add_numbering", False)
        scheme = opts.get("numbering_scheme", "gongwen_standard")

        assert isinstance(remove_num, bool)
        assert isinstance(add_num, bool)
        assert isinstance(scheme, str)
        assert remove_num is True
        assert add_num is True
        assert scheme == "gongwen_standard"

    # ── Proofread plugin contract ───────────────────────────────

    def test_proofread_keys_match_plugin_expectation(self) -> None:
        """Keys match md_validator.py:71-74 / docx_validator.py:64-67 contract."""
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args({"check": ["punct", "typo"]})
        opts = build_execution_options(args)

        # Simulate what proofread validators do:
        symbol_pairing = bool(opts.get("enable_symbol_pairing", True))
        symbol_correction = bool(opts.get("enable_symbol_correction", True))
        typos_rule = bool(opts.get("enable_typos_rule", True))
        sensitive_word = bool(opts.get("enable_sensitive_word", False))

        assert symbol_pairing is True
        assert symbol_correction is False
        assert typos_rule is True
        assert sensitive_word is False

    # ── PDF plugin contract ─────────────────────────────────────

    def test_pages_key_type_matches_plugin_expectation(self) -> None:
        """Pages is list[int], as converter.py:225-226 expects."""
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "action": "split_pdf",
                "to": "",
                "pages": "1-3,5",
            }
        )
        opts = build_execution_options(args)

        # Simulate what layout/operations/converter.py:225-226 does:
        pages: list[int] = [int(p) for p in opts.get("pages", []) if isinstance(p, (int, float))]
        assert pages == [1, 2, 3, 5]
        assert len(pages) == 4


class TestDryRunEffectiveOptions:
    """``--dry-run`` output must include normalized options."""

    def test_dry_run_options_are_normalized(self) -> None:
        """The effective_options in dry-run JSON uses normalized keys."""
        from docwen_cli.commands.execution_options import build_execution_options

        args = _fake_convert_args(
            {
                "check": ["all"],
                "clean_numbering": "remove",
                "add_numbering": "gongwen_standard",
            }
        )
        opts = build_execution_options(args)

        # Numbering keys are normalized
        assert "remove_numbering" in opts
        assert "add_numbering" in opts
        assert "numbering_scheme" in opts
        assert "clean_numbering" not in opts
        assert "add_numbering_scheme" not in opts

        # Proofread keys are normalized
        assert "enable_symbol_pairing" in opts
        assert "enable_symbol_correction" in opts
        assert "enable_typos_rule" in opts
        assert "enable_sensitive_word" in opts
        assert "proofread_checks" not in opts
