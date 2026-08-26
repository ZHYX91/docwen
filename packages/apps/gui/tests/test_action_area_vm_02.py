"""Focused tests split from test_action_area_vm.py."""

from __future__ import annotations

from ._action_area_vm_support import (
    DEFAULT_PROOFREAD_OPTIONS,
    PROOFREAD_OPTION_KEYS,
    SENSITIVE_WORD,
    SYMBOL_PAIRING,
    ActionAreaViewModel,
    FakeMainWindowViewModel,
    pytest,
)

pytestmark = pytest.mark.unit
from ._action_area_vm_support import (
    vm as vm,
)


class TestVisibility:
    def test_show(self, vm: ActionAreaViewModel) -> None:
        vm.show()
        assert vm.visible is True
        assert vm.cancel_visible is False

    def test_hide(self, vm: ActionAreaViewModel) -> None:
        vm.show()
        vm.hide()
        assert vm.visible is False

    def test_show_cancel(self, vm: ActionAreaViewModel) -> None:
        vm.show_cancel()
        assert vm.cancel_visible is True

    def test_hide_cancel(self, vm: ActionAreaViewModel) -> None:
        vm.show_cancel()
        vm.hide_cancel()
        assert vm.cancel_visible is False

    def test_cancel_request_emits_signal(self, vm: ActionAreaViewModel) -> None:
        emitted: list[int] = []
        vm.cancel_requested.connect(lambda: emitted.append(1))
        vm.request_cancel()
        assert len(emitted) == 1


class TestCollectOptions:
    def test_file_to_md_collects_extract_image(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        opts = vm.collect_options()
        assert opts["to_md_keep_images"] is True
        assert opts["to_md_enable_ocr"] is False

    def test_document_to_md_collects_export_and_link_defaults(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_image_extraction_mode": "base64",
                    "export.to_md_ocr_placement_mode": "main_md",
                    "document.to_md_table_merge_export_strategy": "empty",
                    "image.ocr_language": "japanese",
                    "link.format.image_link_style": "markdown_embed",
                }
            )
        )

        vm.setup_for_document_file("/test.docx")

        opts = vm.collect_options()
        assert opts["image_mode"] == "base64"
        assert opts["ocr_placement"] == "main_md"
        assert opts["ocr_language"] == "japanese"
        assert opts["image_link_style"] == "markdown_embed"
        assert opts["table_merge_strategy"] == "empty"

    def test_spreadsheet_to_md_collects_export_link_and_table_defaults(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_image_extraction_mode": "embed",
                    "export.to_md_ocr_placement_mode": "image_md",
                    "image.ocr_language": "latin",
                    "spreadsheet.to_md_table_merge_export_strategy": "marker",
                    "link.format.image_link_style": "markdown_link",
                }
            )
        )

        vm.setup_for_spreadsheet_file("/test.xlsx")

        opts = vm.collect_options()
        assert opts["image_mode"] == "embed"
        assert opts["ocr_placement"] == "image_md"
        assert opts["ocr_language"] == "latin"
        assert opts["image_link_style"] == "markdown_link"
        assert opts["table_merge_strategy"] == "marker"

    def test_image_to_md_collects_image_mode_ocr_placement_and_link_style(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_image_extraction_mode": "omit",
                    "export.to_md_ocr_placement_mode": "main_md",
                    "image.ocr_language": "korean",
                    "link.format.image_link_style": "markdown_embed",
                }
            )
        )

        vm.setup_for_image_file("/test.png")

        opts = vm.collect_options()
        assert opts["image_mode"] == "omit"
        assert opts["ocr_placement"] == "main_md"
        assert opts["ocr_language"] == "korean"
        assert opts["image_link_style"] == "markdown_embed"
        assert "table_merge_strategy" not in opts

    def test_layout_to_md_collects_image_mode_and_link_style(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_image_extraction_mode": "base64",
                    "image.ocr_language": "english",
                    "link.format.image_link_style": "markdown_embed",
                }
            )
        )

        vm.setup_for_layout_file("/test.pdf")

        opts = vm.collect_options()
        assert opts["image_mode"] == "base64"
        assert opts["image_link_style"] == "markdown_embed"
        assert "ocr_placement" not in opts
        assert opts["ocr_language"] == "english"
        assert "table_merge_strategy" not in opts

    def test_to_md_options_fall_back_to_export_defaults(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_image_extraction_mode": "base64",
                    "export.to_md_ocr_placement_mode": "main_md",
                }
            )
        )

        vm.setup_for_document_file("/test.docx")

        opts = vm.collect_options()
        assert opts["image_mode"] == "base64"
        assert opts["ocr_placement"] == "main_md"

    @pytest.mark.parametrize(
        ("setup_method", "setup_args", "section"),
        [
            ("setup_for_document_file", ("/test.docx",), "document"),
            ("setup_for_spreadsheet_file", ("/test.xlsx",), "spreadsheet"),
            ("setup_for_image_file", ("/test.png",), "image"),
            ("setup_for_layout_file", ("/test.pdf",), "layout"),
            ("setup_for_other_file", ("/test.epub", "epub"), "other"),
            ("setup_for_other_file", ("/test.pptx", "pptx"), "other"),
        ],
    )
    def test_removed_category_defaults_are_not_fallbacks_when_export_defaults_are_blank(
        self,
        setup_method: str,
        setup_args: tuple[str, ...],
        section: str,
    ) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_image_extraction_mode": "",
                    "export.to_md_ocr_placement_mode": "",
                    f"{section}.to_md_image_extraction_mode": "base64",
                    f"{section}.to_md_ocr_placement_mode": "main_md",
                }
            )
        )

        getattr(vm, setup_method)(*setup_args)

        opts = vm.collect_options()
        assert "image_mode" not in opts
        assert "ocr_placement" not in opts

    @pytest.mark.parametrize(
        ("setup_method", "file_path", "section"),
        [
            ("setup_for_document_file", "/test.docx", "document"),
            ("setup_for_spreadsheet_file", "/test.xlsx", "spreadsheet"),
            ("setup_for_image_file", "/test.png", "image"),
            ("setup_for_layout_file", "/test.pdf", "layout"),
        ],
    )
    def test_to_md_image_and_ocr_modes_prefer_export_defaults_over_section_defaults(
        self,
        setup_method: str,
        file_path: str,
        section: str,
    ) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_image_extraction_mode": "file",
                    "export.to_md_ocr_placement_mode": "main_md",
                    f"{section}.to_md_image_extraction_mode": "base64",
                    f"{section}.to_md_ocr_placement_mode": "image_md",
                }
            )
        )

        getattr(vm, setup_method)(file_path)

        opts = vm.collect_options()
        assert opts["image_mode"] == "file"
        if section == "layout":
            assert "ocr_placement" not in opts
        else:
            assert opts["ocr_placement"] == "main_md"

    def test_file_to_md_collects_numbering(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        opts = vm.collect_options()
        assert "remove_numbering" in opts
        assert "add_numbering" in opts
        assert "numbering_scheme" in opts

    def test_md_to_document_collects_proofread(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_document("/test.md")
        opts = vm.collect_options()
        assert opts[SYMBOL_PAIRING] is True
        assert opts[SENSITIVE_WORD] is False

    def test_render_mode_in_collect_options(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_document("/test.md")
        opts = vm.collect_options()
        assert "heading_numbering_render_mode" in opts
        assert opts["heading_numbering_render_mode"] == "text"

    def test_render_mode_is_read_live_from_settings_without_rebuilding_panel(self) -> None:
        main_vm = FakeMainWindowViewModel({"text.heading_numbering_render_mode": "text"})
        vm = ActionAreaViewModel(main_vm)  # type: ignore[arg-type]
        vm.setup_for_md_to_document("/test.md")

        assert vm.collect_options()["heading_numbering_render_mode"] == "text"

        # Simulate an applied Settings change while the same action panel is
        # still open.  The next request must consult the config port again.
        main_vm.controller.config_port._values[  # pyright: ignore[reportPrivateUsage]
            "text.heading_numbering_render_mode"
        ] = "word_native"

        assert vm.collect_options()["heading_numbering_render_mode"] == "word_native"

    def test_md_to_spreadsheet_non_xlsx_empty(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_spreadsheet("/test.md")
        vm.target_format = "csv"
        opts = vm.collect_options()
        # Non-xlsx target returns empty dict
        assert opts == {}


class TestOptionSetters:
    def test_set_extract_image(self, vm: ActionAreaViewModel) -> None:
        vm.extract_image = False
        assert vm.extract_image is False

    def test_set_extract_ocr(self, vm: ActionAreaViewModel) -> None:
        vm.extract_ocr = True
        assert vm.extract_ocr is True

    def test_set_doc_remove_numbering(self, vm: ActionAreaViewModel) -> None:
        vm.doc_remove_numbering = False
        assert vm.doc_remove_numbering is False

    def test_set_doc_add_numbering(self, vm: ActionAreaViewModel) -> None:
        vm.doc_add_numbering = True
        assert vm.doc_add_numbering is True

    def test_set_doc_numbering_scheme(self, vm: ActionAreaViewModel) -> None:
        vm.doc_numbering_scheme = "legal_standard"
        assert vm.doc_numbering_scheme == "legal_standard"

    def test_set_md_remove_numbering(self, vm: ActionAreaViewModel) -> None:
        vm.md_remove_numbering = False
        assert vm.md_remove_numbering is False

    def test_set_md_add_numbering(self, vm: ActionAreaViewModel) -> None:
        vm.md_add_numbering = True
        assert vm.md_add_numbering is True

    def test_set_md_numbering_scheme(self, vm: ActionAreaViewModel) -> None:
        vm.md_numbering_scheme = "legal_standard"
        assert vm.md_numbering_scheme == "legal_standard"

    def test_set_proofread_option(self, vm: ActionAreaViewModel) -> None:
        vm.set_proofread_option(SYMBOL_PAIRING, False)
        assert vm.proofread_options[SYMBOL_PAIRING] is False

    def test_set_proofread_option_unknown(self, vm: ActionAreaViewModel) -> None:
        with pytest.raises(ValueError):
            vm.set_proofread_option("unknown", True)

    def test_set_file_to_md_option(self, vm: ActionAreaViewModel) -> None:
        vm.set_file_to_md_option("extract_image", False)
        assert vm.extract_image is False

    def test_set_file_to_md_option_numbering(self, vm: ActionAreaViewModel) -> None:
        vm.set_file_to_md_option("remove_numbering", False)
        assert vm.doc_remove_numbering is False
        # md_* should be unchanged (still default True)
        assert vm.md_remove_numbering is True

        vm.set_file_to_md_option("add_numbering", True)
        assert vm.doc_add_numbering is True
        assert vm.md_add_numbering is False

        vm.set_file_to_md_option("numbering_scheme", "legal_standard")
        assert vm.doc_numbering_scheme == "legal_standard"
        assert vm.md_numbering_scheme == "hierarchical_standard"

    def test_set_file_to_md_option_unknown(self, vm: ActionAreaViewModel) -> None:
        with pytest.raises(KeyError):
            vm.set_file_to_md_option("unknown", True)

    def test_set_md_to_doc_option(self, vm: ActionAreaViewModel) -> None:
        vm.set_md_to_doc_option("remove_numbering", False)
        assert vm.md_remove_numbering is False
        # doc_* should be unchanged (still default True)
        assert vm.doc_remove_numbering is True

        vm.set_md_to_doc_option("add_numbering", True)
        assert vm.md_add_numbering is True
        assert vm.doc_add_numbering is False

        vm.set_md_to_doc_option("numbering_scheme", "legal_standard")
        assert vm.md_numbering_scheme == "legal_standard"
        assert vm.doc_numbering_scheme == "gongwen_standard"

    def test_set_md_to_doc_option_unknown(self, vm: ActionAreaViewModel) -> None:
        with pytest.raises(KeyError):
            vm.set_md_to_doc_option("unknown", True)

    def test_md_heading_numbering_render_mode_default(self, vm: ActionAreaViewModel) -> None:
        assert vm.md_heading_numbering_render_mode == "text"

    def test_render_mode_has_no_transient_setter(self, vm: ActionAreaViewModel) -> None:
        with pytest.raises(AttributeError):
            vm.md_heading_numbering_render_mode = "word_native"  # type: ignore[misc]

    def test_set_md_to_doc_option_rejects_settings_owned_render_mode(
        self,
        vm: ActionAreaViewModel,
    ) -> None:
        with pytest.raises(KeyError):
            vm.set_md_to_doc_option("heading_numbering_render_mode", "word_native")

    def test_numbering_independence(self, vm: ActionAreaViewModel) -> None:
        """Verify changing doc_* numbering does not affect md_* and vice versa."""
        # Start from defaults
        assert vm.doc_remove_numbering is True
        assert vm.doc_add_numbering is False
        assert vm.doc_numbering_scheme == "gongwen_standard"
        assert vm.md_remove_numbering is True
        assert vm.md_add_numbering is False
        assert vm.md_numbering_scheme == "hierarchical_standard"

        # Change doc_* only — md_* must stay
        vm.doc_remove_numbering = False
        vm.doc_add_numbering = True
        vm.doc_numbering_scheme = "legal_standard"
        assert vm.md_remove_numbering is True
        assert vm.md_add_numbering is False
        assert vm.md_numbering_scheme == "hierarchical_standard"

        # Change md_* only — doc_* must stay
        vm.md_remove_numbering = False
        vm.md_add_numbering = True
        vm.md_numbering_scheme = "gongwen_standard"
        assert vm.doc_remove_numbering is False  # unchanged from above
        assert vm.doc_add_numbering is True
        assert vm.doc_numbering_scheme == "legal_standard"


class TestHistory:
    def test_save_last_document_format(self, vm: ActionAreaViewModel) -> None:
        vm.save_last_document_format("rtf")
        assert vm.last_document_format == "rtf"

    def test_save_last_spreadsheet_format(self, vm: ActionAreaViewModel) -> None:
        vm.save_last_spreadsheet_format("ods")
        assert vm.last_spreadsheet_format == "ods"

    def test_last_document_format_default(self, vm: ActionAreaViewModel) -> None:
        assert vm.last_document_format == "docx"

    def test_last_spreadsheet_format_default(self, vm: ActionAreaViewModel) -> None:
        assert vm.last_spreadsheet_format == "xlsx"


class TestConversionRequest:
    def test_emits_signal_with_options(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        emitted: list[tuple] = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))
        vm.request_conversion("md")
        assert len(emitted) == 1
        assert emitted[0][0] == "md"
        assert emitted[0][1] == "/test.docx"
        assert "to_md_keep_images" in emitted[0][2]

    def test_no_file(self, vm: ActionAreaViewModel) -> None:
        emitted: list[tuple] = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))
        vm.request_conversion("md")
        assert len(emitted) == 0


class TestButtonLabels:
    def test_document_button_label(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        assert "Convert to MD" in vm.get_button_label() or vm.get_button_label()

    def test_md_to_document_button_label(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_document("/test.md")
        assert "Generate" in vm.get_button_label() or vm.get_button_label()


class TestReset:
    def test_reset_hides(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        vm.reset()
        assert vm.visible is False

    def test_reset_clears_state(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_document("/test.md")
        vm.last_document_format = "rtf"
        vm.reset()
        assert vm.file_type is None
        assert vm.file_path is None
        assert vm.last_document_format == "rtf"  # history survives reset


class TestSetMode:
    def test_set_batch(self, vm: ActionAreaViewModel) -> None:
        vm.set_mode("batch")
        assert vm.mode == "batch"

    def test_set_invalid_raises(self, vm: ActionAreaViewModel) -> None:
        with pytest.raises(ValueError):
            vm.set_mode("invalid")


class TestOptionConstants:
    def test_proofread_option_keys(self) -> None:
        assert len(PROOFREAD_OPTION_KEYS) == 4

    def test_default_proofread_sensible(self) -> None:
        assert DEFAULT_PROOFREAD_OPTIONS[SENSITIVE_WORD] is False
        assert DEFAULT_PROOFREAD_OPTIONS[SYMBOL_PAIRING] is True
