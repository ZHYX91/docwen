"""Focused tests split from test_main_window_projection_binding.py."""

from __future__ import annotations

from ._main_window_projection_binding_support import (
    QWidget,
    SimpleNamespace,
    _bind_admitted_ref,
    _file_ref,
    _make_window_with_config,
    _write_format_fixture,
    pytest,
)
from ._main_window_projection_binding_support import (
    left_frame as left_frame,
)

pytestmark = pytest.mark.gui
from ._main_window_projection_binding_support import (
    right_frame as right_frame,
)
from ._main_window_projection_binding_support import (
    right_stack as right_stack,
)
from ._main_window_projection_binding_support import (
    window as window,
)


class TestExecutionThreadDispatch:
    def test_batch_thread_uses_controller_execute_batch(self, qapp) -> None:
        from docwen_gui.main_window import _ExecutionThread

        calls: list[object] = []

        class _Controller:
            def execute_batch(self, request):
                calls.append(request)
                return ["batch-result"]

            def execute_single(self, request):
                raise AssertionError("batch execution must not call execute_single")

        request = SimpleNamespace(request_id="request-1")
        context = {"request_id": "request-1", "batch": True}
        thread = _ExecutionThread(
            controller=_Controller(),  # type: ignore[arg-type]
            request=request,  # type: ignore[arg-type]
            context=context,
            batch_execution=True,
        )
        received: list[tuple[object, dict]] = []
        thread.result_signal.connect(lambda result, ctx: received.append((result, ctx)))

        thread.run()

        assert calls == [request]
        assert received == [(["batch-result"], context)]

    def test_aggregate_thread_uses_admitted_controller_boundary(self, qapp) -> None:
        from docwen_gui.main_window import _ExecutionThread

        calls: list[str] = []

        class _Controller:
            def execute_aggregate(self, request, action_name: str):
                calls.append(action_name)
                return {"request": request}

            def execute_single(self, request):
                raise AssertionError("aggregate execution must not call execute_single")

        request = SimpleNamespace(request_id="request-1")
        context = {"request_id": "request-1"}
        thread = _ExecutionThread(
            controller=_Controller(),  # type: ignore[arg-type]
            request=request,  # type: ignore[arg-type]
            context=context,
            aggregate_action_name="merge_pdfs",
        )
        received: list[tuple[object, dict]] = []
        thread.result_signal.connect(lambda result, ctx: received.append((result, ctx)))

        thread.run()

        assert calls == ["merge_pdfs"]
        assert received == [({"request": request}, context)]

    @pytest.mark.parametrize("target_format", ["wps", "pdf"])
    def test_txt_document_context_builds_markdown_runtime_request_for_office_bridge_targets(
        self, window, tmp_path, target_format: str
    ) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "note.txt"
        source.write_text("# Title\n\ncontent", encoding="utf-8")
        window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}

        request, _context = window._build_request(
            file_path=str(source),
            target_format=target_format,
            action_name="",
            options={},
        )

        input_ref = request.input_refs[0]
        assert request.target_format == target_format
        assert input_ref.format == "markdown"
        assert input_ref.category == "markdown"

    @pytest.mark.parametrize(
        ("filename", "fmt", "category"),
        [
            ("page.html", "html", "markup"),
            ("book.epub", "epub", "markup"),
            ("slides.pptx", "pptx", "presentation"),
        ],
    )
    def test_action_only_plugin_formats_build_source_specific_markdown_request(
        self,
        window,
        right_frame,
        tmp_path,
        filename: str,
        fmt: str,
        category: str,
    ) -> None:
        source = tmp_path / filename
        _write_format_fixture(source, fmt)
        _bind_admitted_ref(window, source, category, fmt)

        window._view_model.set_selected_file(_file_ref(str(source), category, fmt))

        assert right_frame.isHidden() or not right_frame.isVisible()
        assert window._action_area_vm.visible is True
        assert window._action_area_vm.file_type == fmt

        request, _context = window._build_request(
            file_path=str(source),
            target_format="md",
            action_name=window._action_area_vm.action_name,
            options=window._action_area_vm.collect_options(),
        )

        input_ref = request.input_refs[0]
        assert request.action_name == ""
        assert request.target_format == "md"
        assert request.options["to_md_keep_images"] is True
        assert request.options["to_md_enable_ocr"] is False
        assert input_ref.format == fmt
        assert input_ref.category == category

    def test_presentation_action_only_request_consumes_link_style_default(self, qapp, tmp_path) -> None:
        window = _make_window_with_config(
            qapp,
            {
                "export.to_md_ocr_placement_mode": "image_md",
                "link.format.image_link_style": "markdown_link",
            },
        )
        try:
            right_frame = window.findChild(QWidget, "rightPanelFrame")
            source = tmp_path / "slides.pptx"
            _write_format_fixture(source, "pptx")
            _bind_admitted_ref(window, source, "presentation", "pptx")

            window._view_model.set_selected_file(_file_ref(str(source), "presentation", "pptx"))

            assert right_frame is not None
            assert right_frame.isHidden() or not right_frame.isVisible()
            assert window._action_area_vm.visible is True
            assert window._action_area_vm.file_type == "pptx"

            request, _context = window._build_request(
                file_path=str(source),
                target_format="md",
                action_name=window._action_area_vm.action_name,
                options=window._action_area_vm.collect_options(),
            )

            assert request.options["to_md_keep_images"] is True
            assert request.options["to_md_enable_ocr"] is False
            assert request.options["ocr_placement"] == "image_md"
            assert request.options["image_link_style"] == "markdown_link"
            input_ref = request.input_refs[0]
            assert input_ref.format == "pptx"
            assert input_ref.category == "presentation"
        finally:
            window.close()

    def test_validate_docx_builds_proofread_category_request(self, window, tmp_path) -> None:
        source = tmp_path / "sample.docx"
        _write_format_fixture(source, "docx")
        _bind_admitted_ref(window, source, "document", "docx")

        request, _context = window._build_request(
            file_path=str(source),
            target_format="docx",
            action_name="validate",
            options={
                "symbol_pairing": False,
                "symbol_correction": True,
                "typos_rule": False,
                "sensitive_word": True,
            },
        )

        input_ref = request.input_refs[0]
        assert request.action_name == "validate"
        assert request.target_format == "docx"
        assert input_ref.format == "docx"
        assert input_ref.category == "document"
        assert request.options == {
            "enable_symbol_pairing": False,
            "enable_symbol_correction": True,
            "enable_typos_rule": False,
            "enable_sensitive_word": True,
        }

    def test_validate_markdown_text_context_builds_proofread_category_request(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "note.md"
        source.write_text("# Title", encoding="utf-8")
        window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}

        request, _context = window._build_request(
            file_path=str(source),
            target_format="markdown",
            action_name="validate",
            options={
                "symbol_pairing": True,
                "symbol_correction": False,
                "typos_rule": True,
                "sensitive_word": False,
            },
            route_options=(
                "enable_symbol_pairing",
                "enable_symbol_correction",
                "enable_typos_rule",
                "enable_sensitive_word",
            ),
        )

        input_ref = request.input_refs[0]
        assert request.action_name == "validate"
        assert request.target_format == "markdown"
        assert input_ref.format == "markdown"
        assert input_ref.category == "markdown"
        assert request.options == {
            "enable_symbol_pairing": True,
            "enable_symbol_correction": False,
            "enable_typos_rule": True,
            "enable_sensitive_word": False,
        }
        assert "locale" not in request.options
        assert "yaml_key_labels" not in request.options

    def test_merge_tables_builds_spreadsheet_category_request(self, window, tmp_path) -> None:
        source = tmp_path / "table.xlsx"
        _write_format_fixture(source, "xlsx")
        _bind_admitted_ref(window, source, "spreadsheet", "xlsx")

        request, _context = window._build_request(
            file_path=str(source),
            target_format="xlsx",
            action_name="merge_tables",
            options={"merge_mode": "col"},
        )

        input_ref = request.input_refs[0]
        assert request.action_name == "merge_tables"
        assert request.target_format == "xlsx"
        assert request.options == {"merge_mode": "col"}
        assert input_ref.format == "xlsx"
        assert input_ref.category == "spreadsheet"

    @pytest.mark.parametrize(
        ("category", "fmt", "target", "filename"),
        [
            ("document", "docx", "doc", "sample.docx"),
            ("document", "docx", "wps", "sample.docx"),
            ("document", "wps", "docx", "sample.wps"),
            ("spreadsheet", "xlsx", "xls", "table.xlsx"),
            ("spreadsheet", "xlsx", "tsv", "table.xlsx"),
            ("spreadsheet", "xlsx", "et", "table.xlsx"),
            ("spreadsheet", "tsv", "xlsx", "table.tsv"),
            ("layout", "pdf", "png", "layout.pdf"),
            ("layout", "pdf", "odt", "layout.pdf"),
        ],
    )
    def test_conversion_panel_standard_targets_keep_reachable_source_route(
        self,
        window,
        tmp_path,
        category: str,
        fmt: str,
        target: str,
        filename: str,
    ) -> None:
        source = tmp_path / filename
        _write_format_fixture(source, fmt)
        _bind_admitted_ref(window, source, category, fmt)

        request, _context = window._build_request(
            file_path=str(source),
            target_format=target,
            action_name="",
            options={},
        )

        input_ref = request.input_refs[0]
        assert request.action_name == ""
        assert request.target_format == target
        assert input_ref.format == fmt
        assert input_ref.category == category

    def test_layout_render_preserves_render_dpi_option(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "layout.pdf"
        source.write_bytes(b"%PDF-1.4\n")
        window._file_contexts = {_normalize_path(str(source)): ("pdf", "layout")}

        request, _context = window._build_request(
            file_path=str(source),
            target_format="jpg",
            action_name="",
            options={"render_dpi": 600},
        )

        input_ref = request.input_refs[0]
        assert request.action_name == ""
        assert request.target_format == "jpg"
        assert request.options == {"render_dpi": 600}
        assert input_ref.format == "pdf"
        assert input_ref.category == "layout"

    def test_layout_to_markdown_preserves_render_dpi_option(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "layout.pdf"
        source.write_bytes(b"%PDF-1.4\n")
        window._file_contexts = {_normalize_path(str(source)): ("pdf", "layout")}

        request, _context = window._build_request(
            file_path=str(source),
            target_format="md",
            action_name="",
            options={"render_dpi": 600},
        )

        assert request.input_refs[0].category == "layout"
        assert request.target_format == "md"
        assert request.options["render_dpi"] == 600

    def test_layout_pdf_normalize_drops_render_dpi_option(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "layout.pdf"
        source.write_bytes(b"%PDF-1.4\n")
        window._file_contexts = {_normalize_path(str(source)): ("pdf", "layout")}

        request, _context = window._build_request(
            file_path=str(source),
            target_format="pdf",
            action_name="",
            options={"render_dpi": 600},
            route_options=(),
        )

        assert request.input_refs[0].category == "layout"
        assert request.target_format == "pdf"
        assert "render_dpi" not in request.options

    def test_document_to_markdown_drops_layout_render_dpi_option(self, window, tmp_path) -> None:
        source = tmp_path / "report.docx"
        _write_format_fixture(source, "docx")
        _bind_admitted_ref(window, source, "document", "docx")

        request, _context = window._build_request(
            file_path=str(source),
            target_format="md",
            action_name="",
            options={"render_dpi": 600},
            route_options=("locale", "yaml_key_labels"),
        )

        assert request.input_refs[0].category == "document"
        assert request.target_format == "md"
        assert "render_dpi" not in request.options

    def test_merge_images_to_tiff_builds_image_category_request(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "image.png"
        source.write_bytes(b"\x89PNG\r\n\x1a\n")
        window._file_contexts = {_normalize_path(str(source)): ("png", "image")}

        request, _context = window._build_request(
            file_path=str(source),
            target_format="tif",
            action_name="merge_images_to_tiff",
            options={},
        )

        input_ref = request.input_refs[0]
        assert request.action_name == "merge_images_to_tiff"
        assert request.target_format == "tif"
        assert input_ref.format == "png"
        assert input_ref.category == "image"

    def test_standard_image_conversion_builds_image_category_request(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "image.png"
        source.write_bytes(b"\x89PNG\r\n\x1a\n")
        window._file_contexts = {_normalize_path(str(source)): ("png", "image")}

        request, _context = window._build_request(
            file_path=str(source),
            target_format="webp",
            action_name="",
            options={"compress_mode": "lossless"},
        )

        input_ref = request.input_refs[0]
        assert request.action_name == ""
        assert request.target_format == "webp"
        assert input_ref.format == "png"
        assert input_ref.category == "image"

    def test_standard_image_conversion_keeps_only_image_format_options(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "image.png"
        source.write_bytes(b"\x89PNG\r\n\x1a\n")
        window._file_contexts = {_normalize_path(str(source)): ("png", "image")}

        request, _context = window._build_request(
            file_path=str(source),
            target_format="webp",
            action_name="",
            options={
                "compress_mode": "limit_size",
                "size_limit": 512,
                "size_unit": "KB",
                "quality_mode": "a4",
            },
            route_options=("compress_mode", "size_limit", "size_unit"),
        )

        assert request.options == {
            "compress_mode": "limit_size",
            "size_limit": 512,
            "size_unit": "KB",
        }

    def test_image_to_markdown_drops_image_format_and_pdf_options(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "image.png"
        source.write_bytes(b"\x89PNG\r\n\x1a\n")
        window._file_contexts = {_normalize_path(str(source)): ("png", "image")}

        request, context = window._build_request(
            file_path=str(source),
            target_format="md",
            action_name="",
            options={
                "compress_mode": "limit_size",
                "size_limit": 512,
                "size_unit": "KB",
                "quality_mode": "a4",
                "to_md_keep_images": True,
            },
            route_options=("to_md_keep_images", "locale", "yaml_key_labels"),
        )

        assert request.input_refs[0].format == "png"
        assert request.input_refs[0].category == "image"
        assert request.options["to_md_keep_images"] is True
        assert "locale" in request.options
        assert "yaml_key_labels" in request.options
        for key in ("compress_mode", "size_limit", "size_unit", "quality_mode"):
            assert key not in request.options
            assert key not in context["options"]

    def test_image_invoice_builds_image_category_request(self, window, tmp_path) -> None:
        source = tmp_path / "invoice.jpg"
        _write_format_fixture(source, "jpeg")
        _bind_admitted_ref(window, source, "image", "jpeg")

        request, _context = window._build_request(
            file_path=str(source),
            target_format="md",
            action_name="invoice_cn",
            options={"to_md_enable_ocr": True},
        )

        input_ref = request.input_refs[0]
        assert request.action_name == "invoice_cn"
        assert request.target_format == "md"
        assert input_ref.format == "jpeg"
        assert input_ref.category == "image"
