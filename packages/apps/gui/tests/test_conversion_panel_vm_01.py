"""Focused tests split from test_conversion_panel_vm.py."""

from __future__ import annotations

from ._conversion_panel_vm_support import (
    SENSITIVE_WORD,
    SYMBOL_CORRECTION,
    SYMBOL_PAIRING,
    TYPOS_RULE,
    VALIDATION_OPTION_KEYS,
    ConversionPanelViewModel,
    FakeMainWindowViewModel,
    pytest,
)

pytestmark = pytest.mark.unit
from ._conversion_panel_vm_support import (
    vm as vm,
)


class TestInitialState:
    def test_file_category_is_none(self, vm: ConversionPanelViewModel) -> None:
        assert vm.file_category is None

    def test_current_format_is_empty(self, vm: ConversionPanelViewModel) -> None:
        assert vm.current_format == ""

    def test_current_file_path_is_none(self, vm: ConversionPanelViewModel) -> None:
        assert vm.current_file_path is None

    def test_file_list_is_empty(self, vm: ConversionPanelViewModel) -> None:
        assert vm.file_list == []

    def test_ui_mode_defaults_to_single(self, vm: ConversionPanelViewModel) -> None:
        assert vm.ui_mode == "single"

    def test_has_files_is_false(self, vm: ConversionPanelViewModel) -> None:
        assert vm.has_files is False


class TestImageDefaults:
    def test_compress_mode_default(self, vm: ConversionPanelViewModel) -> None:
        assert vm.compress_mode == "lossless"

    def test_size_limit_default(self, vm: ConversionPanelViewModel) -> None:
        assert vm.size_limit == 200

    def test_size_unit_default(self, vm: ConversionPanelViewModel) -> None:
        assert vm.size_unit == "KB"

    def test_tiff_mode_default(self, vm: ConversionPanelViewModel) -> None:
        assert vm.tiff_mode == "smart"

    def test_pdf_quality_default(self, vm: ConversionPanelViewModel) -> None:
        assert vm.pdf_quality == "original"

    def test_reads_image_defaults_from_config_port(self) -> None:
        vm = ConversionPanelViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "image.compress_mode": "limit_size",
                    "image.size_limit": 512,
                    "image.size_unit": "MB",
                    "image.tiff_mode": "rgb",
                    "image.pdf_quality": "fit_a4",
                }
            )
        )

        assert vm.compress_mode == "limit_size"
        assert vm.size_limit == 512
        assert vm.size_unit == "MB"
        assert vm.tiff_mode == "RGB"
        assert vm.pdf_quality == "a4"

    def test_invalid_image_defaults_fall_back_safely(self) -> None:
        vm = ConversionPanelViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "image.compress_mode": "bad",
                    "image.size_limit": "not-int",
                    "image.size_unit": "GB",
                    "image.tiff_mode": "smart",
                    "image.pdf_quality": "original",
                }
            )
        )

        assert vm.compress_mode == "lossless"
        assert vm.size_limit == 200
        assert vm.size_unit == "KB"

    def test_numeric_string_image_size_limit_falls_back_safely(self) -> None:
        vm = ConversionPanelViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {"image.size_limit": "512"}
            )
        )

        assert vm.size_limit == 200

    @pytest.mark.parametrize("invalid_value", [True, 512.5, [512], {"value": 512}])
    def test_non_integer_image_size_limit_falls_back_safely(self, invalid_value: object) -> None:
        vm = ConversionPanelViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {"image.size_limit": invalid_value}
            )
        )

        assert vm.size_limit == 200


class TestSpreadsheetDefaults:
    def test_merge_mode_default(self, vm: ConversionPanelViewModel) -> None:
        assert vm.merge_mode == 3

    def test_reference_table_name_default(self, vm: ConversionPanelViewModel) -> None:
        assert vm.reference_table_name == ""

    def test_reads_merge_mode_default_from_config_port(self) -> None:
        vm = ConversionPanelViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {"spreadsheet.merge_mode": 2}
            )
        )

        assert vm.merge_mode == 2

    def test_invalid_merge_mode_default_falls_back_safely(self) -> None:
        vm = ConversionPanelViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {"spreadsheet.merge_mode": 7}
            )
        )

        assert vm.merge_mode == 3

    def test_boolean_merge_mode_default_falls_back_safely(self) -> None:
        vm = ConversionPanelViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {"spreadsheet.merge_mode": True}
            )
        )

        assert vm.merge_mode == 3


class TestLayoutDefaults:
    def test_render_format_default(self, vm: ConversionPanelViewModel) -> None:
        assert vm.render_format == "TIF"

    def test_render_dpi_default(self, vm: ConversionPanelViewModel) -> None:
        assert vm.render_dpi == 300

    def test_reads_render_dpi_default_from_config_port(self) -> None:
        vm = ConversionPanelViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {"layout.render_dpi": 600}
            )
        )

        assert vm.render_dpi == 600

    def test_invalid_render_dpi_default_falls_back_safely(self) -> None:
        vm = ConversionPanelViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {"layout.render_dpi": 72}
            )
        )

        assert vm.render_dpi == 300

    def test_float_render_dpi_default_falls_back_safely(self) -> None:
        vm = ConversionPanelViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {"layout.render_dpi": 600.0}
            )
        )

        assert vm.render_dpi == 300

    def test_page_input_default(self, vm: ConversionPanelViewModel) -> None:
        assert vm.page_input == ""

    def test_pdf_total_pages_default(self, vm: ConversionPanelViewModel) -> None:
        assert vm.pdf_total_pages == 0

    def test_pdf_file_name_default(self, vm: ConversionPanelViewModel) -> None:
        assert vm.pdf_file_name == ""


class TestValidationOptions:
    def test_default_validation_options(self, vm: ConversionPanelViewModel) -> None:
        opts = vm.validation_options
        assert opts[SYMBOL_PAIRING] is True
        assert opts[SYMBOL_CORRECTION] is True
        assert opts[TYPOS_RULE] is True
        assert opts[SENSITIVE_WORD] is False

    def test_is_any_validation_option_checked(self, vm: ConversionPanelViewModel) -> None:
        assert vm.is_any_validation_option_checked is True

    def test_set_validation_option_emits_signal(self, vm: ConversionPanelViewModel) -> None:
        signals: list[int] = []
        vm.state_changed.connect(lambda: signals.append(1))
        vm.set_validation_option(SYMBOL_PAIRING, False)
        assert len(signals) == 1
        assert vm.validation_options[SYMBOL_PAIRING] is False

    def test_set_validation_option_noop_does_not_emit(self, vm: ConversionPanelViewModel) -> None:
        signals: list[int] = []
        vm.state_changed.connect(lambda: signals.append(1))
        vm.set_validation_option(SYMBOL_PAIRING, True)  # already True
        assert len(signals) == 0

    def test_set_validation_option_unknown_key_raises(self, vm: ConversionPanelViewModel) -> None:
        with pytest.raises(ValueError):
            vm.set_validation_option("unknown_key", True)

    def test_all_unchecked(self, vm: ConversionPanelViewModel) -> None:
        for key in VALIDATION_OPTION_KEYS:
            vm.set_validation_option(key, False)
        assert vm.is_any_validation_option_checked is False


class TestSetFileInfo:
    def test_set_file_info_updates_state(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")
        assert vm.file_category == "document"
        assert vm.current_format == "docx"
        assert vm.current_file_path == "/test.docx"

    def test_set_file_info_emits_state_changed(self, vm: ConversionPanelViewModel) -> None:
        signals: list[int] = []
        vm.state_changed.connect(lambda: signals.append(1))
        vm.set_file_info("spreadsheet", "xlsx")
        assert len(signals) == 1

    def test_set_file_info_normalizes_format(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "DOCX")
        assert vm.current_format == "docx"

    def test_set_file_info_stores_file_list(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("spreadsheet", "xlsx", file_list=["/a.xlsx", "/b.xlsx"])
        assert vm.file_list == ["/a.xlsx", "/b.xlsx"]

    def test_set_file_info_ui_mode(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx", ui_mode="batch")
        assert vm.ui_mode == "batch"

    def test_set_pdf_info_updates_pages_and_file_name(self, vm: ConversionPanelViewModel) -> None:
        signals: list[int] = []
        vm.state_changed.connect(lambda: signals.append(1))

        vm.set_pdf_info(12, "annual-report.pdf")

        assert vm.pdf_total_pages == 12
        assert vm.pdf_file_name == "annual-report.pdf"
        assert len(signals) == 1

    def test_set_pdf_info_noop_does_not_emit(self, vm: ConversionPanelViewModel) -> None:
        vm.set_pdf_info(12, "annual-report.pdf")
        signals: list[int] = []
        vm.state_changed.connect(lambda: signals.append(1))

        vm.set_pdf_info(12, "annual-report.pdf")

        assert len(signals) == 0

    def test_file_path_change_clears_pdf_specific_state(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("layout", "pdf", file_path="/readable-a.pdf")
        vm.page_input = "1-3"
        vm.set_pdf_info(12, "readable-a.pdf")
        signals: list[int] = []
        vm.state_changed.connect(lambda: signals.append(1))

        vm.set_file_info("layout", "pdf", file_path="/unreadable-b.pdf")

        assert vm.page_input == ""
        assert vm.pdf_total_pages == 0
        assert vm.pdf_file_name == ""
        assert len(signals) == 1


class TestConversionFormatLookups:
    def test_document_formats(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx")
        fmts = vm.get_conversion_formats()
        assert "DOCX" not in fmts
        assert "DOC" in fmts
        assert "WPS" in fmts

    def test_wps_source_filters_same_format_but_keeps_other_document_targets(
        self, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("document", "wps")
        fmts = vm.get_conversion_formats()
        assert "WPS" not in fmts
        assert "DOCX" in fmts
        assert "RTF" in fmts

    def test_spreadsheet_formats(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("spreadsheet", "xlsx")
        fmts = vm.get_conversion_formats()
        assert "XLSX" not in fmts
        assert "XLS" in fmts
        assert "TSV" in fmts
        assert "ET" in fmts

    def test_tsv_source_only_exposes_reachable_spreadsheet_target(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("spreadsheet", "tsv")
        assert vm.get_conversion_formats() == ["XLSX"]

    def test_csv_source_filters_unreachable_spreadsheet_targets(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("spreadsheet", "csv")
        fmts = vm.get_conversion_formats()
        assert fmts == ["XLSX", "XLS", "ODS"]
        assert "CSV" not in fmts
        assert "TSV" not in fmts
        assert "ET" not in fmts

    def test_et_source_exposes_manifest_backed_spreadsheet_targets(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("spreadsheet", "et")
        assert vm.get_conversion_formats() == ["XLSX", "XLS", "ODS", "CSV"]

    def test_image_formats(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("image", "png")
        fmts = vm.get_conversion_formats()
        assert "PNG" not in fmts
        assert "JPG" in fmts

    def test_image_formats_filter_same_format_aliases(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("image", "jpeg")
        fmts = vm.get_conversion_formats()
        assert "JPG" not in fmts
        assert "PNG" in fmts

    def test_image_limit_size_allows_same_compressible_format(self, vm: ConversionPanelViewModel) -> None:
        vm.compress_mode = "limit_size"
        vm.set_file_info("image", "jpeg")
        fmts = vm.get_conversion_formats()
        assert "JPG" in fmts
        assert "PNG" in fmts

    def test_image_limit_size_keeps_same_non_compressible_format_hidden(self, vm: ConversionPanelViewModel) -> None:
        vm.compress_mode = "limit_size"
        vm.set_file_info("image", "png")
        fmts = vm.get_conversion_formats()
        assert "PNG" not in fmts
        assert "JPG" in fmts

    def test_layout_formats(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("layout", "pdf")
        fmts = vm.get_conversion_formats()
        assert fmts == []

    def test_ofd_layout_exposes_pdf_conversion(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("layout", "ofd")
        assert vm.get_conversion_formats() == ["PDF"]

    def test_layout_export_formats_include_manifest_backed_document_targets(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("layout", "pdf")
        assert vm.get_layout_export_formats() == ["DOCX", "DOC", "ODT", "RTF"]

    def test_layout_render_formats_include_manifest_backed_image_targets(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("layout", "pdf")
        assert vm.get_layout_render_formats() == ["PNG", "JPG", "TIF"]

    def test_layout_pdf_operations_only_for_pdf_source(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("layout", "pdf")
        assert vm.supports_layout_pdf_operations is True
        vm.set_file_info("layout", "ofd")
        assert vm.supports_layout_pdf_operations is False

    def test_none_category(self, vm: ConversionPanelViewModel) -> None:
        assert vm.get_conversion_formats() == []

    def test_saveas_formats(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx")
        assert vm.get_saveas_formats() == ["PDF"]

    def test_spreadsheet_saveas_pdf_uses_route_backed_sources(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("spreadsheet", "csv")
        assert vm.get_saveas_formats() == ["PDF"]

    def test_tsv_source_hides_unreachable_pdf_saveas(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("spreadsheet", "tsv")
        assert vm.get_saveas_formats() == []

    def test_saveas_empty_for_layout(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("layout", "pdf")
        assert vm.get_saveas_formats() == []


class TestFormatNormalization:
    def test_jpg(self) -> None:
        assert ConversionPanelViewModel.normalize_format("jpg") == "jpeg"
        assert ConversionPanelViewModel.normalize_format("jpeg") == "jpeg"

    def test_tif(self) -> None:
        assert ConversionPanelViewModel.normalize_format("tif") == "tiff"
        assert ConversionPanelViewModel.normalize_format("tiff") == "tiff"

    def test_heif(self) -> None:
        assert ConversionPanelViewModel.normalize_format("heif") == "heic"

    def test_docx_unchanged(self) -> None:
        assert ConversionPanelViewModel.normalize_format("docx") == "docx"

    def test_case_insensitive(self) -> None:
        assert ConversionPanelViewModel.normalize_format("JPG") == "jpeg"


class TestSizeValidation:
    def test_valid_kb(self) -> None:
        assert ConversionPanelViewModel.validate_size_input("200", "KB") is True

    def test_valid_mb(self) -> None:
        assert ConversionPanelViewModel.validate_size_input("50", "MB") is True

    def test_invalid_kb_too_large(self) -> None:
        assert ConversionPanelViewModel.validate_size_input("20000", "KB") is False

    def test_invalid_kb_zero(self) -> None:
        assert ConversionPanelViewModel.validate_size_input("0", "KB") is False

    def test_invalid_non_numeric(self) -> None:
        assert ConversionPanelViewModel.validate_size_input("abc", "KB") is False

    def test_boundary_min(self) -> None:
        assert ConversionPanelViewModel.validate_size_input("1", "KB") is True

    def test_boundary_max_kb(self) -> None:
        assert ConversionPanelViewModel.validate_size_input("10240", "KB") is True

    def test_boundary_max_mb(self) -> None:
        assert ConversionPanelViewModel.validate_size_input("100", "MB") is True


class TestPageInputValidation:
    def test_wildcard_star(self) -> None:
        assert ConversionPanelViewModel.validate_page_input("*", 5) is True

    def test_wildcard_star_single_page(self) -> None:
        assert ConversionPanelViewModel.validate_page_input("*", 1) is False

    def test_wildcard_hash(self) -> None:
        assert ConversionPanelViewModel.validate_page_input("#", 5) is True

    def test_custom_range(self) -> None:
        assert ConversionPanelViewModel.validate_page_input("1-5", 10) is True

    def test_custom_comma(self) -> None:
        assert ConversionPanelViewModel.validate_page_input("1,3,5", 10) is True

    def test_all_pages_rejected(self) -> None:
        assert ConversionPanelViewModel.validate_page_input("1-5", 5) is False

    def test_empty_input(self) -> None:
        assert ConversionPanelViewModel.validate_page_input("", 5) is False

    def test_invalid_chars(self) -> None:
        assert ConversionPanelViewModel.validate_page_input("abc", 5) is False

    def test_no_total_pages(self) -> None:
        assert ConversionPanelViewModel.validate_page_input("1-5", 0) is True


class TestSplitInputParsing:
    def test_star(self) -> None:
        mode, pages = ConversionPanelViewModel.parse_split_input("*")
        assert mode == "every_page"
        assert pages is None

    def test_hash(self) -> None:
        mode, pages = ConversionPanelViewModel.parse_split_input("#")
        assert mode == "odd_even"
        assert pages is None

    def test_custom_single(self) -> None:
        mode, pages = ConversionPanelViewModel.parse_split_input("1,3,5")
        assert mode == "custom"
        assert pages == [1, 3, 5]

    def test_custom_range(self) -> None:
        mode, pages = ConversionPanelViewModel.parse_split_input("1-5")
        assert mode == "custom"
        assert pages == [1, 2, 3, 4, 5]

    def test_custom_mixed(self) -> None:
        mode, pages = ConversionPanelViewModel.parse_split_input("1-3,7,9-11")
        assert mode == "custom"
        assert pages == [1, 2, 3, 7, 9, 10, 11]


class TestPageRangeParsing:
    def test_single(self) -> None:
        assert ConversionPanelViewModel.parse_page_ranges("5") == [5]

    def test_comma(self) -> None:
        assert ConversionPanelViewModel.parse_page_ranges("1,3,5") == [1, 3, 5]

    def test_range(self) -> None:
        assert ConversionPanelViewModel.parse_page_ranges("1-3") == [1, 2, 3]

    def test_mixed(self) -> None:
        assert ConversionPanelViewModel.parse_page_ranges("1-3,7,9-11") == [1, 2, 3, 7, 9, 10, 11]

    def test_reversed_range(self) -> None:
        assert ConversionPanelViewModel.parse_page_ranges("5-3") == [3, 4, 5]

    def test_chinese_separators(self) -> None:
        result = ConversionPanelViewModel.parse_page_ranges("1，3，5")
        assert result == [1, 3, 5]

    def test_chinese_range(self) -> None:
        result = ConversionPanelViewModel.parse_page_ranges("1至5")
        assert result == [1, 2, 3, 4, 5]

    def test_sorted_result(self) -> None:
        assert ConversionPanelViewModel.parse_page_ranges("3,1,2") == [1, 2, 3]

    def test_page_zero_excluded(self) -> None:
        result = ConversionPanelViewModel.parse_page_ranges("1,0,3")
        assert 0 not in result
        assert result == [1, 3]

    def test_empty(self) -> None:
        assert ConversionPanelViewModel.parse_page_ranges("") == []

    def test_invalid_raises(self) -> None:
        with pytest.raises(ValueError):
            ConversionPanelViewModel.parse_page_ranges("abc-def")
