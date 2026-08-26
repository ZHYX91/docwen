"""Focused tests split from test_main_window_projection_binding.py."""

from __future__ import annotations

from ._main_window_projection_binding_support import (
    _DOCX_TEMPLATE_ID,
    _bind_admitted_ref,
    _file_ref,
    _load_request_templates,
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


class TestRuntimeRequestBinding:
    def test_gui_confirmation_fails_closed_without_frozen_inspection(self, window, tmp_path) -> None:
        from docwen_core.models import ConversionRequest, FileRef

        source = tmp_path / "unadmitted.md"
        source.write_text("# Content\n", encoding="utf-8")
        request = ConversionRequest(
            request_id="unadmitted",
            input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
            target_format="docx",
        )

        assert window._confirm_request_admission(request) is False

    def test_request_builder_admits_programmatic_path_through_core(self, window, tmp_path) -> None:
        from docwen_core.models import FILE_INSPECTION_METADATA_KEY
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "not-added.md"
        source.write_text("# Content\n", encoding="utf-8")
        window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}

        request, _context = window._build_request(
            file_path=str(source),
            target_format="docx",
            action_name="",
            options={},
        )

        assert request.input_refs[0].format == "markdown"
        assert request.input_refs[0].category == "markdown"
        assert FILE_INSPECTION_METADATA_KEY in request.input_refs[0].metadata

    def test_cross_family_request_requires_and_records_explicit_acceptance(self, window, tmp_path, monkeypatch) -> None:
        from docwen_core.models import (
            FILE_ADMISSION_ACCEPTANCE_METADATA_KEY,
            FILE_INSPECTION_METADATA_KEY,
            FileInspection,
            admission_is_satisfied,
        )

        source = tmp_path / "renamed.docx"
        source.write_bytes(b"%PDF-1.4\n% deterministic probe\n")
        outcome = window._view_model.add_files([str(source)])
        assert len(outcome.added) == 1
        request, _context = window._build_request(
            file_path=str(source),
            target_format="md",
            action_name="",
            options={},
        )

        monkeypatch.setattr("docwen_gui.dialogs.feedback.confirm", lambda *_args, **_kwargs: False)
        assert window._confirm_request_admission(request) is False
        assert FILE_ADMISSION_ACCEPTANCE_METADATA_KEY not in request.input_refs[0].metadata

        monkeypatch.setattr("docwen_gui.dialogs.feedback.confirm", lambda *_args, **_kwargs: True)
        assert window._confirm_request_admission(request) is True
        raw = request.input_refs[0].metadata[FILE_INSPECTION_METADATA_KEY]
        inspection = FileInspection.from_dict(raw)
        assert admission_is_satisfied(inspection, request.input_refs[0].metadata)
        assert FILE_ADMISSION_ACCEPTANCE_METADATA_KEY in outcome.added[0].metadata
        entry = window._batch_list_vm.get_file_entry(str(source))
        assert entry is not None
        assert FILE_ADMISSION_ACCEPTANCE_METADATA_KEY in entry.metadata

        monkeypatch.setattr(
            "docwen_gui.dialogs.feedback.confirm",
            lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("must not ask twice")),
        )
        second_request, _context = window._build_request(
            file_path=str(source),
            target_format="md",
            action_name="",
            options={},
        )
        assert window._confirm_request_admission(second_request) is True

    def test_confirmed_file_replacement_requires_refresh_before_execution(self, window, tmp_path, monkeypatch) -> None:
        source = tmp_path / "renamed.docx"
        source.write_bytes(b"%PDF-1.4\nfirst version\n")
        window._view_model.add_files([str(source)])
        first_request, _context = window._build_request(
            file_path=str(source),
            target_format="md",
            action_name="",
            options={},
        )
        monkeypatch.setattr("docwen_gui.dialogs.feedback.confirm", lambda *_args, **_kwargs: True)
        assert window._confirm_request_admission(first_request) is True

        source.write_bytes(b"%PDF-1.4\nreplacement with a different size\n")
        second_request, _context = window._build_request(
            file_path=str(source),
            target_format="md",
            action_name="",
            options={},
        )
        monkeypatch.setattr(
            "docwen_gui.dialogs.feedback.confirm",
            lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("stale fact must stop before dialog")),
        )

        assert window._confirm_request_admission(second_request) is False
        from docwen_gui.i18n import t as _t

        guidance = _t("main_window.file_admission_changed")
        assert "重新添加" in guidance or "add it again" in guidance
        assert any(row.message == guidance for row in window._info_area_vm.history_rows)

    def test_request_keeps_localized_ingress_warning_and_inspection_after_text_route_normalization(
        self, window, tmp_path
    ) -> None:
        from docwen_core.models import FILE_INSPECTION_METADATA_KEY, FileInspection
        from docwen_gui.file_admission_i18n import render_file_inspection_message

        source = tmp_path / "note.md"
        source.write_text("plain text", encoding="utf-8")
        outcome = window._view_model.add_files([str(source)])
        assert len(outcome.added) == 1

        request, _context = window._build_request(
            file_path=str(source),
            target_format="docx",
            action_name="",
            options={},
        )

        input_ref = request.input_refs[0]
        assert input_ref.format == "txt"
        assert input_ref.category == "markdown"
        inspection = FileInspection.from_dict(input_ref.metadata[FILE_INSPECTION_METADATA_KEY])
        assert inspection.warning_code == "FILE_FORMAT_COMPATIBLE_TEXT"
        assert input_ref.warning_message == render_file_inspection_message(inspection)
        assert input_ref.warning_message != inspection.warning_message

    def test_docx_template_selection_is_added_to_document_request(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "note.md"
        source.write_text("# Title", encoding="utf-8")
        window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}
        _load_request_templates(window)
        selector = window._template_selector.get_selector("docx")
        assert selector is not None
        selector.select_template("Corporate Report", selection_source="user")

        request, context = window._build_request(
            file_path=str(source),
            target_format="docx",
            action_name="",
            options={"remove_numbering": True},
        )

        assert request.options["template_name"] == _DOCX_TEMPLATE_ID
        assert request.options["remove_numbering"] is True
        assert context["options"]["template_name"] == _DOCX_TEMPLATE_ID

    def test_document_to_markdown_request_carries_locale_yaml_labels(self, window, tmp_path) -> None:
        from docwen_gui.i18n import get_locale, set_locale

        previous_locale = get_locale()
        set_locale("de_DE")
        try:
            source = tmp_path / "report.docx"
            _write_format_fixture(source, "docx")
            _bind_admitted_ref(window, source, "document", "docx")

            request, context = window._build_request(
                file_path=str(source),
                target_format="md",
                action_name="",
                options={"to_md_keep_images": True},
            )

            assert request.options["locale"] == "de_DE"
            assert request.options["yaml_key_labels"] == {"title": "Titel", "subtitle": "Untertitel"}
            assert context["options"]["yaml_key_labels"]["title"] == "Titel"
        finally:
            set_locale(previous_locale)

    def test_spreadsheet_to_markdown_request_carries_locale_yaml_labels(self, window, tmp_path) -> None:
        from docwen_gui.i18n import get_locale, set_locale

        previous_locale = get_locale()
        set_locale("de_DE")
        try:
            source = tmp_path / "book.xlsx"
            _write_format_fixture(source, "xlsx")
            _bind_admitted_ref(window, source, "spreadsheet", "xlsx")

            request, context = window._build_request(
                file_path=str(source),
                target_format="md",
                action_name="",
                options={"to_md_keep_images": True},
            )

            assert request.options["locale"] == "de_DE"
            assert request.options["yaml_key_labels"]["title"] == "Titel"
            assert context["options"]["yaml_key_labels"]["subtitle"] == "Untertitel"
        finally:
            set_locale(previous_locale)

    def test_policy02_password_stays_in_request_but_is_redacted_from_gui_context(self, window, tmp_path) -> None:
        source = tmp_path / "protected.xlsx"
        _write_format_fixture(source, "xlsx")
        _bind_admitted_ref(window, source, "spreadsheet", "xlsx")

        request, context = window._build_request(
            file_path=str(source),
            target_format="ods",
            action_name="",
            options={
                "spreadsheet_password": "pw-SECRET-879",
                "allow_spreadsheet_protection_loss": True,
            },
        )

        assert request.options["spreadsheet_password"] == "pw-SECRET-879"
        assert request.options["allow_spreadsheet_protection_loss"] is True
        assert context["options"] == {
            "spreadsheet_password": "<redacted>",
            "allow_spreadsheet_protection_loss": True,
        }
        assert "pw-SECRET-879" not in str(context)

    def test_image_to_markdown_request_carries_locale_yaml_labels(self, window, tmp_path) -> None:
        from docwen_gui.i18n import get_locale, set_locale
        from docwen_gui.main_window import _normalize_path

        previous_locale = get_locale()
        set_locale("ja_JP")
        try:
            source = tmp_path / "sample.png"
            source.write_bytes(b"\x89PNG\r\n\x1a\n")
            window._file_contexts = {_normalize_path(str(source)): ("png", "image")}

            request, context = window._build_request(
                file_path=str(source),
                target_format="md",
                action_name="",
                options={"to_md_keep_images": True},
            )

            assert request.options["locale"] == "ja_JP"
            assert request.options["yaml_key_labels"] == {"title": "タイトル", "subtitle": "サブタイトル"}
            assert context["options"]["locale"] == "ja_JP"
        finally:
            set_locale(previous_locale)

    def test_image_to_markdown_request_consumes_link_style_default(self, qapp, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        window = _make_window_with_config(
            qapp,
            {
                "link.format.image_link_style": "markdown_embed",
            },
        )
        try:
            source = tmp_path / "sample.png"
            source.write_bytes(b"\x89PNG\r\n\x1a\n")
            window._file_contexts = {_normalize_path(str(source)): ("png", "image")}

            window._view_model.set_selected_file(_file_ref(str(source), "image", "png"))

            request, context = window._build_request(
                file_path=str(source),
                target_format="md",
                action_name=window._action_area_vm.action_name,
                options=window._action_area_vm.collect_options(),
            )

            assert request.options["image_link_style"] == "markdown_embed"
            assert context["options"]["image_link_style"] == "markdown_embed"
            assert request.input_refs[0].format == "png"
            assert request.input_refs[0].category == "image"
        finally:
            window.close()

    def test_markup_to_markdown_request_carries_locale_yaml_labels(self, window, tmp_path) -> None:
        from docwen_gui.i18n import get_locale, set_locale
        from docwen_gui.main_window import _normalize_path

        previous_locale = get_locale()
        set_locale("de_DE")
        try:
            source = tmp_path / "page.html"
            source.write_text(
                "<html><head><title>Probe</title></head><body><h1>Probe</h1></body></html>",
                encoding="utf-8",
            )
            window._file_contexts = {_normalize_path(str(source)): ("html", "markup")}

            request, context = window._build_request(
                file_path=str(source),
                target_format="md",
                action_name="",
                options={"to_md_keep_images": True},
            )

            assert request.input_refs[0].format == "html"
            assert request.input_refs[0].category == "markup"
            assert request.options["locale"] == "de_DE"
            assert request.options["yaml_key_labels"]["title"] == "Titel"
            assert context["options"]["yaml_key_labels"]["subtitle"] == "Untertitel"
        finally:
            set_locale(previous_locale)

    def test_layout_to_markdown_request_carries_locale_yaml_labels(self, window, tmp_path) -> None:
        from docwen_gui.i18n import get_locale, set_locale
        from docwen_gui.main_window import _normalize_path

        previous_locale = get_locale()
        set_locale("de_DE")
        try:
            source = tmp_path / "layout.pdf"
            source.write_bytes(b"%PDF-1.4\n")
            window._file_contexts = {_normalize_path(str(source)): ("pdf", "layout")}

            request, context = window._build_request(
                file_path=str(source),
                target_format="md",
                action_name="",
                options={"to_md_keep_images": False},
            )

            assert request.input_refs[0].format == "pdf"
            assert request.input_refs[0].category == "layout"
            assert request.options["locale"] == "de_DE"
            assert request.options["yaml_key_labels"]["title"] == "Titel"
            assert context["options"]["locale"] == "de_DE"
        finally:
            set_locale(previous_locale)

    def test_invoice_cn_to_markdown_request_carries_locale_yaml_labels(self, window, tmp_path) -> None:
        from docwen_gui.i18n import get_locale, set_locale
        from docwen_gui.main_window import _normalize_path

        previous_locale = get_locale()
        set_locale("de_DE")
        try:
            source = tmp_path / "invoice.pdf"
            source.write_bytes(b"%PDF-1.4\n%\x00\x00\x00\x00\n")
            window._file_contexts = {_normalize_path(str(source)): ("pdf", "layout")}

            request, context = window._build_request(
                file_path=str(source),
                target_format="md",
                action_name="invoice_cn",
                options={"to_md_enable_ocr": False},
                route_options=("locale", "yaml_key_labels"),
            )

            assert request.input_refs[0].format == "pdf"
            assert request.input_refs[0].category == "layout"
            assert request.action_name == "invoice_cn"
            assert request.target_format == "md"
            assert request.options["locale"] == "de_DE"
            assert request.options["yaml_key_labels"] == {"title": "Titel", "subtitle": "Untertitel"}
            assert context["options"]["yaml_key_labels"]["subtitle"] == "Untertitel"
        finally:
            set_locale(previous_locale)

    def test_gongwen_to_markdown_request_carries_locale_without_yaml_labels(self, window, tmp_path) -> None:
        from docwen_gui.i18n import get_locale, set_locale

        previous_locale = get_locale()
        set_locale("de_DE")
        try:
            source = tmp_path / "gongwen.docx"
            _write_format_fixture(source, "docx")
            _bind_admitted_ref(window, source, "document", "docx")

            request, context = window._build_request(
                file_path=str(source),
                target_format="md",
                action_name="gongwen",
                options={"to_md_enable_ocr": False},
                route_options=("locale",),
            )

            assert request.input_refs[0].format == "docx"
            assert request.input_refs[0].category == "document"
            assert request.action_name == "gongwen"
            assert request.options["locale"] == "de_DE"
            assert "yaml_key_labels" not in request.options
            assert context["options"]["locale"] == "de_DE"
        finally:
            set_locale(previous_locale)

    def test_md_numbering_request_does_not_carry_markdown_export_metadata(self, window, tmp_path) -> None:
        from docwen_gui.i18n import get_locale, set_locale
        from docwen_gui.main_window import _normalize_path

        previous_locale = get_locale()
        set_locale("de_DE")
        try:
            source = tmp_path / "note.md"
            source.write_text("# Title\n", encoding="utf-8")
            window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}

            request, context = window._build_request(
                file_path=str(source),
                target_format="md",
                action_name="process_md_numbering",
                options={"remove_numbering": True},
            )

            assert request.input_refs[0].format == "markdown"
            assert request.input_refs[0].category == "markdown"
            assert request.action_name == "process_md_numbering"
            assert request.options == {"remove_numbering": True}
            assert "locale" not in context["options"]
            assert "yaml_key_labels" not in context["options"]
        finally:
            set_locale(previous_locale)

    def test_docx_template_selection_is_added_to_wps_request(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "note.md"
        source.write_text("# Title", encoding="utf-8")
        window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}
        _load_request_templates(window, xlsx=False)
        selector = window._template_selector.get_selector("docx")
        assert selector is not None
        selector.select_template("Corporate Report", selection_source="user")

        request, _context = window._build_request(
            file_path=str(source),
            target_format="wps",
            action_name="",
            options={},
        )

        assert request.options["template_name"] == _DOCX_TEMPLATE_ID

    def test_docx_template_selection_is_added_to_pdf_request(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "note.md"
        source.write_text("# Title", encoding="utf-8")
        window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}
        _load_request_templates(window, xlsx=False)
        selector = window._template_selector.get_selector("docx")
        assert selector is not None
        selector.select_template("Corporate Report", selection_source="user")

        request, _context = window._build_request(
            file_path=str(source),
            target_format="pdf",
            action_name="",
            options={},
        )

        assert request.options["template_name"] == _DOCX_TEMPLATE_ID

    @pytest.mark.parametrize("suffix", [".txt", ".md"])
    def test_actual_txt_in_markdown_workflow_keeps_selected_template(
        self,
        window,
        tmp_path,
        suffix: str,
    ) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / f"plain{suffix}"
        source.write_text("plain UTF-8 text without Markdown syntax\n", encoding="utf-8")
        window._file_contexts = {_normalize_path(str(source)): ("txt", "markdown")}
        _load_request_templates(window, xlsx=False)
        selector = window._template_selector.get_selector("docx")
        assert selector is not None
        selector.select_template("Corporate Report", selection_source="user")

        request, _context = window._build_request(
            file_path=str(source),
            target_format="docx",
            action_name="",
            options={},
        )

        assert request.options["template_name"] == _DOCX_TEMPLATE_ID
