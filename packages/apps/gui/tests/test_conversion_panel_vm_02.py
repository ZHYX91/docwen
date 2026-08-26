"""Focused tests split from test_conversion_panel_vm.py."""

from __future__ import annotations

from ._conversion_panel_vm_support import (
    BUTTON_COLORS,
    COMPRESSIBLE_FORMATS,
    VALIDATION_OPTION_KEYS,
    ConversionPanelViewModel,
    FakeMainWindowViewModel,
    pytest,
)

pytestmark = pytest.mark.unit
from ._conversion_panel_vm_support import (
    vm as vm,
)


class TestProperties:
    def test_compress_mode_setter(self, vm: ConversionPanelViewModel) -> None:
        vm.compress_mode = "limit_size"
        assert vm.compress_mode == "limit_size"

    def test_compress_mode_invalid(self, vm: ConversionPanelViewModel) -> None:
        with pytest.raises(ValueError):
            vm.compress_mode = "invalid"

    def test_tiff_mode_setter(self, vm: ConversionPanelViewModel) -> None:
        vm.tiff_mode = "RGB"
        assert vm.tiff_mode == "RGB"

    def test_tiff_mode_setter_accepts_settings_alias(self, vm: ConversionPanelViewModel) -> None:
        vm.tiff_mode = "rgb"
        assert vm.tiff_mode == "RGB"

    def test_tiff_mode_invalid(self, vm: ConversionPanelViewModel) -> None:
        with pytest.raises(ValueError):
            vm.tiff_mode = "invalid"

    def test_size_unit_setter(self, vm: ConversionPanelViewModel) -> None:
        vm.size_unit = "MB"
        assert vm.size_unit == "MB"

    def test_size_unit_invalid(self, vm: ConversionPanelViewModel) -> None:
        with pytest.raises(ValueError):
            vm.size_unit = "GB"

    def test_merge_mode_setter(self, vm: ConversionPanelViewModel) -> None:
        vm.merge_mode = 1
        assert vm.merge_mode == 1

    def test_merge_mode_invalid(self, vm: ConversionPanelViewModel) -> None:
        with pytest.raises(ValueError):
            vm.merge_mode = 0

    def test_render_format_setter(self, vm: ConversionPanelViewModel) -> None:
        vm.render_format = "JPG"
        assert vm.render_format == "JPG"

    def test_render_format_setter_accepts_png(self, vm: ConversionPanelViewModel) -> None:
        vm.render_format = "PNG"
        assert vm.render_format == "PNG"

    def test_render_format_invalid(self, vm: ConversionPanelViewModel) -> None:
        with pytest.raises(ValueError):
            vm.render_format = "WEBP"

    def test_render_dpi_setter(self, vm: ConversionPanelViewModel) -> None:
        vm.render_dpi = 600
        assert vm.render_dpi == 600

    def test_render_dpi_invalid(self, vm: ConversionPanelViewModel) -> None:
        with pytest.raises(ValueError):
            vm.render_dpi = 200

    def test_pdf_quality_setter_accepts_settings_alias(self, vm: ConversionPanelViewModel) -> None:
        vm.pdf_quality = "fit_a3"
        assert vm.pdf_quality == "a3"


class TestReset:
    def test_reset_clears_all(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")
        vm.reset()
        assert vm.file_category is None
        assert vm.current_format == ""
        assert vm.current_file_path is None

    def test_reset_emits_signal(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx")
        signals: list[int] = []
        vm.state_changed.connect(lambda: signals.append(1))
        vm.reset()
        assert len(signals) == 1

    def test_reset_restores_configured_image_defaults(self) -> None:
        vm = ConversionPanelViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "image.compress_mode": "limit_size",
                    "image.size_limit": 7,
                    "image.size_unit": "KB",
                    "image.tiff_mode": "rgb",
                    "image.pdf_quality": "fit_a3",
                }
            )
        )
        vm.compress_mode = "lossless"
        vm.size_limit = 123
        vm.size_unit = "MB"
        vm.tiff_mode = "smart"
        vm.pdf_quality = "original"

        vm.reset()

        assert vm.compress_mode == "limit_size"
        assert vm.size_limit == 7
        assert vm.size_unit == "KB"
        assert vm.tiff_mode == "RGB"
        assert vm.pdf_quality == "a3"

    def test_reset_restores_configured_spreadsheet_and_layout_defaults(self) -> None:
        vm = ConversionPanelViewModel(
            FakeMainWindowViewModel(  # type: ignore[arg-type]
                {
                    "spreadsheet.merge_mode": 1,
                    "layout.render_dpi": 600,
                }
            )
        )
        vm.merge_mode = 3
        vm.render_dpi = 150

        vm.reset()

        assert vm.merge_mode == 1
        assert vm.render_dpi == 600


class TestConversionRequest:
    def test_emits_signal(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")
        emitted: list[tuple] = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))
        vm.request_conversion("pdf")
        assert len(emitted) == 1
        assert emitted[0][0] == "pdf"
        assert emitted[0][1] == "/test.docx"

    def test_no_file(self, vm: ConversionPanelViewModel) -> None:
        emitted: list[tuple] = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))
        vm.request_conversion("pdf")
        assert len(emitted) == 0

    def test_with_options(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("image", "png", file_path="/test.png")
        emitted: list[tuple] = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))
        vm.request_conversion("jpeg", options={"compress_mode": "lossless"})
        assert emitted[0][2] == {"compress_mode": "lossless"}


class TestNamedAction:
    def test_emits_signal(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")
        emitted: list[tuple] = []
        vm.named_action_requested.connect(lambda n, fp, o: emitted.append((n, fp, o)))
        vm.request_named_action("validate", options={"symbol_pairing": True})
        assert len(emitted) == 1
        assert emitted[0][0] == "validate"
        assert emitted[0][1] == "/test.docx"
        assert emitted[0][2] == {"symbol_pairing": True}

    def test_no_file(self, vm: ConversionPanelViewModel) -> None:
        emitted: list[tuple] = []
        vm.named_action_requested.connect(lambda n, fp, o: emitted.append((n, fp, o)))
        vm.request_named_action("merge_pdfs")
        assert len(emitted) == 0


class TestFormatConstants:
    def test_button_colors(self) -> None:
        assert BUTTON_COLORS["DOCX"] == "primary"
        assert BUTTON_COLORS["WPS"] == "info"
        assert BUTTON_COLORS["ET"] == "info"
        assert BUTTON_COLORS["TSV"] == "warning"
        assert BUTTON_COLORS["PDF"] == "danger"
        assert BUTTON_COLORS["OFD"] == "success"

    def test_compressible_formats(self) -> None:
        assert set(COMPRESSIBLE_FORMATS) == {"JPG", "JPEG", "WEBP"}

    def test_validation_option_keys(self) -> None:
        assert set(VALIDATION_OPTION_KEYS) == {
            "symbol_pairing",
            "symbol_correction",
            "typos_rule",
            "sensitive_word",
        }
