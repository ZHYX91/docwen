"""Focused tests split from test_action_area_vm.py."""

from __future__ import annotations

from ._action_area_vm_support import (
    MODE_DOCUMENT,
    MODE_IMAGE,
    MODE_LAYOUT,
    MODE_MD_TO_DOCUMENT,
    MODE_MD_TO_SPREADSHEET,
    MODE_SPREADSHEET,
    SENSITIVE_WORD,
    SYMBOL_CORRECTION,
    SYMBOL_PAIRING,
    TYPOS_RULE,
    ActionAreaViewModel,
    FakeMainWindowViewModel,
    _t,
    pytest,
)

pytestmark = pytest.mark.unit
from ._action_area_vm_support import (
    vm as vm,
)


class TestInitialState:
    def test_not_visible(self, vm: ActionAreaViewModel) -> None:
        assert vm.visible is False

    def test_cancel_not_visible(self, vm: ActionAreaViewModel) -> None:
        assert vm.cancel_visible is False

    def test_file_type_none(self, vm: ActionAreaViewModel) -> None:
        assert vm.file_type is None

    def test_file_path_none(self, vm: ActionAreaViewModel) -> None:
        assert vm.file_path is None

    def test_mode_single(self, vm: ActionAreaViewModel) -> None:
        assert vm.mode == "single"

    def test_extract_image_default(self, vm: ActionAreaViewModel) -> None:
        assert vm.extract_image is True

    def test_extract_ocr_default(self, vm: ActionAreaViewModel) -> None:
        assert vm.extract_ocr is False

    def test_proofread_not_shown(self, vm: ActionAreaViewModel) -> None:
        assert vm.show_proofread is False


class TestSetupDocumentFile:
    def test_sets_file_type(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        assert vm.file_type == MODE_DOCUMENT

    def test_sets_file_path(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        assert vm.file_path == "/test.docx"

    def test_makes_visible(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        assert vm.visible is True

    def test_shows_numbering(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        assert vm.show_numbering is True

    def test_reads_numbering_defaults_from_config_port(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "document.to_md_remove_numbering": False,
                    "document.to_md_add_numbering": True,
                    "document.to_md_default_scheme": "legal_standard",
                }
            )
        )

        vm.setup_for_document_file("/test.docx")

        assert vm.doc_remove_numbering is False
        assert vm.doc_add_numbering is True
        assert vm.doc_numbering_scheme == "legal_standard"

    def test_reads_document_extract_and_optimization_defaults_from_config_port(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "document.to_md_keep_images": False,
                    "document.to_md_enable_ocr": True,
                    "document.to_md_enable_optimization": True,
                    "document.to_md_optimization_type": "gongwen",
                }
            )
        )

        vm.setup_for_document_file("/test.docx")

        assert vm.extract_image is False
        assert vm.extract_ocr is True
        assert vm.action_name == "gongwen"


class TestSetupMdToDocumentDefaults:
    def test_reads_defaults_from_config_port(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "text.remove_numbering": False,
                    "text.add_numbering": True,
                    "text.numbering_scheme": "legal_standard",
                    "text.heading_numbering_render_mode": "word_native",
                }
            )
        )

        vm.setup_for_md_to_document("/test.md")

        assert vm.md_remove_numbering is False
        assert vm.md_add_numbering is True
        assert vm.md_numbering_scheme == "legal_standard"
        assert vm.md_heading_numbering_render_mode == "word_native"

    def test_exposes_non_blocking_status_when_discovery_is_unavailable(self) -> None:
        vm = ActionAreaViewModel()
        vm.setup_for_document_file("/test.docx")
        assert vm.show_optimize is True
        assert vm.optimization_choices_result.status == "failed"
        assert vm.optimization_choices == ()
        assert vm.action_name == ""

    def test_ocr_disabled(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_document_file("/test.docx")
        assert vm.extract_ocr is False

    def test_emits_state_changed(self, vm: ActionAreaViewModel) -> None:
        signals: list[int] = []
        vm.state_changed.connect(lambda: signals.append(1))
        vm.setup_for_document_file("/test.docx")
        assert len(signals) == 1


class TestSetupSpreadsheetFile:
    def test_sets_file_type(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_spreadsheet_file("/test.xlsx")
        assert vm.file_type == MODE_SPREADSHEET

    def test_no_numbering(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_spreadsheet_file("/test.xlsx")
        assert vm.show_numbering is False

    def test_hides_optimize_without_a_matching_runtime_resource(self) -> None:
        vm = ActionAreaViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]
        vm.setup_for_spreadsheet_file("/test.xlsx")
        assert vm.show_optimize is False

    def test_ocr_disabled(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_spreadsheet_file("/test.xlsx")
        assert vm.extract_ocr is False

    def test_reads_spreadsheet_defaults_from_config_port(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "spreadsheet.to_md_keep_images": False,
                    "spreadsheet.to_md_enable_ocr": True,
                }
            )
        )

        vm.setup_for_spreadsheet_file("/test.xlsx")

        assert vm.extract_image is False
        assert vm.extract_ocr is True


class TestSetupImageFile:
    def test_sets_file_type(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_image_file("/test.png")
        assert vm.file_type == MODE_IMAGE

    def test_ocr_enabled(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_image_file("/test.png")
        assert vm.extract_ocr is True

    def test_reads_image_defaults_from_config_port(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "image.to_md_keep_images": False,
                    "image.to_md_enable_ocr": False,
                }
            )
        )

        vm.setup_for_image_file("/test.png")

        assert vm.extract_image is False
        assert vm.extract_ocr is False


class TestSetupLayoutFile:
    def test_sets_file_type(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_layout_file("/test.pdf")
        assert vm.file_type == MODE_LAYOUT

    def test_ocr_disabled(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_layout_file("/test.pdf")
        assert vm.extract_ocr is False

    def test_reads_layout_defaults_and_optimization_from_config_port(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "layout.to_md_keep_images": False,
                    "layout.to_md_enable_ocr": True,
                    "layout.to_md_enable_optimization": True,
                    "layout.to_md_optimization_type": "invoice_cn",
                }
            )
        )

        vm.setup_for_layout_file("/test.pdf")

        assert vm.extract_image is False
        assert vm.extract_ocr is True
        assert vm.action_name == "invoice_cn"


class TestSetupOtherFile:
    def test_sets_file_type_to_actual_format(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_other_file("/test.xyz", "xyz")
        assert vm.file_type == "xyz"

    def test_no_numbering(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_other_file("/test.xyz", "xyz")
        assert vm.show_numbering is False

    def test_ready_empty_catalog_hides_optimize(self) -> None:
        vm = ActionAreaViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]
        vm.setup_for_other_file("/test.xyz", "xyz")
        assert vm.show_optimize is False
        assert vm.optimization_choices_result.status == "ready"

    def test_visible(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_other_file("/test.xyz", "xyz")
        assert vm.visible is True

    def test_extract_image_true(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_other_file("/test.xyz", "xyz")
        assert vm.extract_image is True

    def test_reads_other_extract_defaults_from_config_port(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "other.to_md_keep_images": False,
                    "other.to_md_enable_ocr": True,
                    "other.to_md_enable_optimization": True,
                    "other.to_md_optimization_type": "gongwen",
                }
            )
        )

        vm.setup_for_other_file("/test.epub", "epub")

        assert vm.extract_image is False
        assert vm.extract_ocr is True
        assert vm.action_name == ""
        assert vm.collect_options() == {
            "to_md_keep_images": False,
            "to_md_enable_ocr": True,
        }

    def test_uses_file_to_markdown_label_for_concrete_other_format(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_other_file("/test.epub", "epub")
        assert vm.get_button_label() == _t("action_area.document.export_markdown", "Convert to MD")

    def test_collects_file_to_markdown_options_for_concrete_other_format(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_other_file("/test.pptx", "pptx")
        vm.set_file_to_md_option("extract_image", False)
        vm.set_file_to_md_option("extract_ocr", True)

        assert vm.collect_options() == {
            "to_md_keep_images": False,
            "to_md_enable_ocr": True,
        }

    def test_presentation_other_format_collects_image_link_style(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {"link.format.image_link_style": "markdown_link"}
            )
        )

        vm.setup_for_other_file("/test.pptx", "pptx")

        assert vm.collect_options() == {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "image_link_style": "markdown_link",
        }

    def test_presentation_other_format_collects_export_image_mode(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_image_extraction_mode": "base64",
                    "link.format.image_link_style": "markdown_link",
                }
            )
        )

        vm.setup_for_other_file("/test.pptx", "pptx")

        assert vm.collect_options() == {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "image_mode": "base64",
            "image_link_style": "markdown_link",
        }

    def test_presentation_other_format_prefers_export_image_mode_over_other_default(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_image_extraction_mode": "omit",
                    "other.to_md_image_extraction_mode": "base64",
                    "link.format.image_link_style": "markdown_link",
                }
            )
        )

        vm.setup_for_other_file("/test.ppt", "ppt")

        assert vm.collect_options() == {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "image_mode": "omit",
            "image_link_style": "markdown_link",
        }

    def test_markup_other_format_leaves_image_link_style_to_runtime_semantics(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {"link.format.image_link_style": "markdown_link"}
            )
        )

        vm.setup_for_other_file("/test.epub", "epub")

        assert vm.collect_options() == {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
        }

    def test_markup_other_format_collects_export_image_mode(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_image_extraction_mode": "base64",
                    "link.format.image_link_style": "markdown_link",
                }
            )
        )

        vm.setup_for_other_file("/test.epub", "epub")

        assert vm.collect_options() == {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "image_mode": "base64",
        }

    def test_markup_other_format_prefers_export_image_mode_over_other_default(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_image_extraction_mode": "omit",
                    "other.to_md_image_extraction_mode": "base64",
                }
            )
        )

        vm.setup_for_other_file("/test.mhtml", "mhtml")

        assert vm.collect_options() == {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "image_mode": "omit",
        }

    def test_markup_other_format_collects_export_ocr_placement(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_ocr_placement_mode": "main_md",
                    "image.ocr_language": "japanese",
                }
            )
        )

        vm.setup_for_other_file("/test.mhtml", "mhtml")

        assert vm.collect_options() == {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "ocr_placement": "main_md",
            "ocr_language": "japanese",
        }

    def test_markup_other_format_prefers_export_ocr_placement_over_other_default(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_ocr_placement_mode": "image_md",
                    "other.to_md_ocr_placement_mode": "main_md",
                    "image.ocr_language": "japanese",
                }
            )
        )

        vm.setup_for_other_file("/test.mhtml", "mhtml")

        assert vm.collect_options() == {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "ocr_placement": "image_md",
            "ocr_language": "japanese",
        }

    def test_presentation_other_format_collects_export_ocr_placement(self) -> None:
        vm = ActionAreaViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "export.to_md_ocr_placement_mode": "main_md",
                    "image.ocr_language": "japanese",
                    "link.format.image_link_style": "markdown_link",
                }
            )
        )

        vm.setup_for_other_file("/test.pptx", "pptx")

        assert vm.collect_options() == {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "ocr_placement": "main_md",
            "ocr_language": "japanese",
            "image_link_style": "markdown_link",
        }


class TestSetupMdToDocument:
    def test_sets_file_type(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_document("/test.md")
        assert vm.file_type == MODE_MD_TO_DOCUMENT

    def test_shows_proofread(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_document("/test.md")
        assert vm.show_proofread is True

    def test_proofread_defaults(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_document("/test.md")
        opts = vm.proofread_options
        assert opts[SYMBOL_PAIRING] is True
        assert opts[SYMBOL_CORRECTION] is True
        assert opts[TYPOS_RULE] is True
        assert opts[SENSITIVE_WORD] is False

    def test_target_format_default(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_document("/test.md")
        assert vm.target_format == "docx"

    def test_available_formats(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_document("/test.md")
        assert "DOCX" in vm.available_target_formats
        assert "RTF" in vm.available_target_formats
        assert "WPS" in vm.available_target_formats
        assert "PDF" in vm.available_target_formats


class TestSetupMdToSpreadsheet:
    def test_sets_file_type(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_spreadsheet("/test.md")
        assert vm.file_type == MODE_MD_TO_SPREADSHEET

    def test_no_proofread(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_spreadsheet("/test.md")
        assert vm.show_proofread is False

    def test_target_format_default(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_spreadsheet("/test.md")
        assert vm.target_format == "xlsx"

    def test_available_formats(self, vm: ActionAreaViewModel) -> None:
        vm.setup_for_md_to_spreadsheet("/test.md")
        assert "XLSX" in vm.available_target_formats
        assert "CSV" in vm.available_target_formats


class TestSetupModeIsolation:
    def test_md_setup_clears_previous_optimization_action(self, vm: ActionAreaViewModel) -> None:
        vm.optimize_for_type = "invoice_cn"

        vm.setup_for_md_to_document("/test.md")

        assert vm.action_name == ""

    @pytest.mark.parametrize(
        ("method_name", "args"),
        [
            ("setup_for_document_file", ("/test.docx",)),
            ("setup_for_spreadsheet_file", ("/test.xlsx",)),
            ("setup_for_image_file", ("/test.png",)),
            ("setup_for_layout_file", ("/test.pdf",)),
            ("setup_for_other_file", ("/test.epub", "epub")),
            ("setup_for_md_to_document", ("/test.md",)),
            ("setup_for_md_to_spreadsheet", ("/test.md",)),
        ],
    )
    def test_standard_setup_clears_previous_aggregate_state(
        self,
        vm: ActionAreaViewModel,
        method_name: str,
        args: tuple[str, ...],
    ) -> None:
        vm.setup_for_aggregate("merge_pdfs", ["/a.pdf", "/b.pdf"])

        getattr(vm, method_name)(*args)

        assert vm.collect_aggregate_request_context() is None
        assert vm._aggregate_action_name == ""  # pyright: ignore[reportPrivateUsage]
        assert vm._aggregate_file_list == []  # pyright: ignore[reportPrivateUsage]

    def test_aggregate_setup_clears_previous_optimization_action(self, vm: ActionAreaViewModel) -> None:
        vm.optimize_for_type = "gongwen"

        vm.setup_for_aggregate("merge_pdfs", ["/a.pdf", "/b.pdf"])

        assert vm.action_name == ""
