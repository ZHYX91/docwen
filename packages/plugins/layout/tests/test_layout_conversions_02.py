"""Focused tests split from test_layout_conversions.py."""

from __future__ import annotations

from ._layout_conversions_support import (
    Any,
    Path,
    _build_fake_context,
    _create_text_pdf,
    pytest,
    sys,
    tempfile,
)

pytestmark = pytest.mark.contract


class TestLayoutToDocumentOfficeBridge:
    def test_pdf_to_docx_uses_admitted_pdf_format_despite_ofd_suffix(
        self, sample_pdf_path: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """The PDF-specific chain is selected from FileRef.format, not the suffix."""
        import docwen_plugin_layout.to_document.converter as converter
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_layout import LayoutPlugin

        misleading_path = tmp_path / "admitted-pdf.ofd"
        misleading_path.write_bytes(sample_pdf_path.read_bytes())

        def fake_pdf_chain(context, input_path: Path, output_path: Path) -> BridgeResult:
            del context
            assert input_path == misleading_path
            output_path.write_bytes(b"docx-from-admitted-pdf")
            return BridgeResult(True, output_path=str(output_path), backend="admitted-pdf-backend")

        def fail_generic_bridge(*args, **kwargs):
            raise AssertionError("generic layout bridge must not replace the admitted PDF-specific chain")

        monkeypatch.setattr(converter, "_convert_pdf_with_configured_office_priority", fake_pdf_chain)
        monkeypatch.setattr(converter, "convert_with_backend_priority", fail_generic_bridge)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(misleading_path),
                staging,
                "docx",
                source_format="pdf",
            )
            result = LayoutPlugin().convert(context)

        assert result.success is True, f"unexpected error: {result.error}"
        assert result.metrics.extra["backend"] == "admitted-pdf-backend"

    def test_pdf_to_docx_consumes_configured_office_priority_before_pdf2docx(
        self, sample_pdf_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """PDF→DOCX should honor the user-configured external backend order first."""
        import docwen_plugin_layout.to_document.converter as converter
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_layout import LayoutPlugin

        observed_priority: list[str] = []
        observed_candidates: set[str] = set()

        def fake_convert_with_backend_priority(
            input_path,
            output_path,
            *,
            source_format,
            backend_priority,
            com_candidates,
            libreoffice_format,
            cancel=None,
            failure_subject,
        ):
            del input_path, cancel, failure_subject
            assert source_format == "pdf"
            observed_priority.extend(backend_priority)
            observed_candidates.update(com_candidates)
            assert libreoffice_format == "docx"
            Path(output_path).write_bytes(b"docx-by-libreoffice")
            return BridgeResult(True, output_path=str(output_path), backend="LibreOffice")

        def fake_pdf2docx(
            input_path: Path,
            output_path: Path,
            *,
            cancellation: Any = None,
        ) -> BridgeResult:
            raise AssertionError("pdf2docx should only run after configured external backends fail")

        monkeypatch.setattr(converter, "convert_with_backend_priority", fake_convert_with_backend_priority)
        monkeypatch.setattr(converter, "_convert_pdf_with_pdf2docx", fake_pdf2docx)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "docx",
                config_values={
                    "software": {"special_conversions": {"pdf_to_office": ["libreoffice", "msoffice_word"]}}
                },
            )
            result = LayoutPlugin().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert result.artifacts[0].suggested_name == "sample.docx"
            assert (
                result.artifacts[0].media_type
                == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            assert Path(result.artifacts[0].staging_path).read_bytes() == b"docx-by-libreoffice"
            assert result.metrics.extra["engine"] == "office_bridge"
            assert result.metrics.extra["backend"] == "LibreOffice"
            assert observed_priority == ["libreoffice", "msoffice_word"]
            assert observed_candidates == {"msoffice_word"}

    def test_pdf_to_docx_uses_pdf2docx_after_configured_office_backends_fail(
        self, sample_pdf_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """pdf2docx remains the final fallback after Office/LibreOffice fail."""
        import docwen_plugin_layout.to_document.converter as converter
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_layout import LayoutPlugin

        observed_priority: list[str] = []

        def fake_convert_with_backend_priority(
            input_path,
            output_path,
            *,
            source_format,
            backend_priority,
            com_candidates,
            libreoffice_format,
            cancel=None,
            failure_subject,
        ):
            del input_path, output_path, cancel, failure_subject
            assert source_format == "pdf"
            observed_priority.extend(backend_priority)
            assert set(com_candidates) == {"msoffice_word"}
            assert libreoffice_format == "docx"
            return BridgeResult(False, message="backend unavailable")

        def fake_pdf2docx(
            input_path: Path,
            output_path: Path,
            *,
            cancellation: Any = None,
        ) -> BridgeResult:
            assert cancellation is not None
            output_path.write_bytes(b"docx-by-pdf2docx")
            return BridgeResult(True, output_path=str(output_path), backend="pdf2docx")

        monkeypatch.setattr(converter, "convert_with_backend_priority", fake_convert_with_backend_priority)
        monkeypatch.setattr(converter, "_convert_pdf_with_pdf2docx", fake_pdf2docx)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_pdf_path), staging, "docx")
            result = LayoutPlugin().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert Path(result.artifacts[0].staging_path).read_bytes() == b"docx-by-pdf2docx"
            assert result.metrics.extra["engine"] == "pdf2docx"
            assert result.metrics.extra["backend"] == "pdf2docx"
            assert observed_priority == ["msoffice_word", "libreoffice"]

    def test_pdf_to_docx_ignores_wps_writer_even_if_configured(
        self, sample_pdf_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """PDF→DOCX should keep WPS out of the local backend chain."""
        import docwen_plugin_layout.to_document.converter as converter
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_layout import LayoutPlugin

        observed_priority: list[str] = []
        observed_candidates: set[str] = set()

        def fake_convert_with_backend_priority(
            input_path,
            output_path,
            *,
            source_format,
            backend_priority,
            com_candidates,
            libreoffice_format,
            cancel=None,
            failure_subject,
        ):
            del input_path, output_path, cancel, failure_subject
            assert source_format == "pdf"
            observed_priority.extend(backend_priority)
            observed_candidates.update(com_candidates)
            assert libreoffice_format == "docx"
            return BridgeResult(False, message="backend unavailable")

        def fake_pdf2docx(
            input_path: Path,
            output_path: Path,
            *,
            cancellation: Any = None,
        ) -> BridgeResult:
            assert cancellation is not None
            output_path.write_bytes(b"docx-by-pdf2docx")
            return BridgeResult(True, output_path=str(output_path), backend="pdf2docx")

        monkeypatch.setattr(converter, "convert_with_backend_priority", fake_convert_with_backend_priority)
        monkeypatch.setattr(converter, "_convert_pdf_with_pdf2docx", fake_pdf2docx)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "docx",
                config_values={
                    "software": {
                        "special_conversions": {"pdf_to_office": ["wps_writer", "msoffice_word", "libreoffice"]}
                    }
                },
            )
            result = LayoutPlugin().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert result.metrics.extra["engine"] == "pdf2docx"
            assert observed_priority == ["wps_writer", "msoffice_word", "libreoffice"]
            assert observed_candidates == {"msoffice_word"}

    def test_pdf_to_docx_returns_structured_error_when_all_backends_fail(
        self, sample_pdf_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """PDF→DOCX should report a structured error only after all backends fail."""
        import docwen_plugin_layout.to_document.converter as converter
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_layout import LayoutPlugin

        def fake_convert_with_backend_priority(
            input_path,
            output_path,
            *,
            source_format,
            backend_priority,
            com_candidates,
            libreoffice_format,
            cancel=None,
            failure_subject,
        ):
            del input_path, output_path, backend_priority, com_candidates, libreoffice_format, cancel, failure_subject
            assert source_format == "pdf"
            return BridgeResult(False, message="Office and LibreOffice unavailable")

        def fake_pdf2docx(
            input_path: Path,
            output_path: Path,
            *,
            cancellation: Any = None,
        ) -> BridgeResult:
            assert cancellation is not None
            return BridgeResult(False, message="pdf2docx unavailable")

        monkeypatch.setattr(converter, "convert_with_backend_priority", fake_convert_with_backend_priority)
        monkeypatch.setattr(converter, "_convert_pdf_with_pdf2docx", fake_pdf2docx)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_pdf_path), staging, "docx")
            result = LayoutPlugin().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "conversion_failed"
            assert result.error.diagnostic_code == "LAYOUT-PDF2DOCX-FAILED"
            assert "Office and LibreOffice unavailable" in result.error.message
            assert "pdf2docx unavailable" in result.error.message

    def test_pdf_to_docx_missing_pdf2docx_message_describes_final_fallback(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """A missing pdf2docx dependency should be described as the final fallback failure."""
        from docwen_plugin_layout.to_document.converter import _convert_pdf_with_pdf2docx

        monkeypatch.setitem(sys.modules, "pdf2docx", None)

        result = _convert_pdf_with_pdf2docx(tmp_path / "input.pdf", tmp_path / "output.docx")

        assert result.success is False
        assert result.message is not None
        assert "pdf2docx fallback" in result.message
        assert "layout plugin dependencies" in result.message

    def test_pdf2docx_fallback_reads_pdf_content_despite_txt_suffix(self, tmp_path: Path) -> None:
        """The pure-Python fallback must use PDF bytes rather than the filename."""
        from docwen_plugin_layout.to_document.converter import _convert_pdf_with_pdf2docx

        misleading_path = tmp_path / "actual-pdf-content.txt"
        _create_text_pdf(misleading_path, ["PDF2DOCX CONTENT FIRST"])
        output_path = tmp_path / "content-first.docx"

        class CancellationProbe:
            def __init__(self) -> None:
                self.check_count = 0
                self.is_cancelled = False

            def check(self) -> None:
                self.check_count += 1

        cancellation = CancellationProbe()
        result = _convert_pdf_with_pdf2docx(
            misleading_path,
            output_path,
            cancellation=cancellation,
        )

        assert result.success is True, result.message
        assert result.backend == "pdf2docx"
        assert output_path.read_bytes().startswith(b"PK")
        assert cancellation.check_count == 3

    @pytest.mark.parametrize("target", ["doc", "rtf"])
    def test_layout_non_docx_word_targets_honor_configured_word_priority(
        self,
        sample_pdf_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        target: str,
    ) -> None:
        """DOC/RTF targets must consume the authoritative word backend order."""
        import docwen_plugin_layout.to_document.converter as converter
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_layout import LayoutPlugin

        observed_priority: list[str] = []
        observed_candidates: set[str] = set()

        def fake_convert_with_backend_priority(
            input_path,
            output_path,
            *,
            source_format,
            backend_priority,
            com_candidates,
            libreoffice_format,
            cancel=None,
            failure_subject,
        ):
            del input_path, cancel, failure_subject
            assert source_format == "pdf"
            observed_priority.extend(backend_priority)
            observed_candidates.update(com_candidates)
            assert libreoffice_format == target
            Path(output_path).write_bytes(f"{target}-by-office".encode("ascii"))
            return BridgeResult(True, output_path=str(output_path), backend="test Word")

        monkeypatch.setattr(converter, "convert_with_backend_priority", fake_convert_with_backend_priority)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                target,
                config_values={
                    "software": {
                        "default_priority": {"word_processors": ["msoffice_word", "libreoffice", "wps_writer"]}
                    }
                },
            )
            result = LayoutPlugin().convert(context)

        assert result.success is True, f"unexpected error: {result.error}"
        assert observed_priority == ["msoffice_word", "libreoffice", "wps_writer"]
        assert observed_candidates == {"wps_writer", "msoffice_word"}

    def test_layout_odt_target_honors_configured_odt_priority_and_excludes_wps(
        self,
        sample_pdf_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """ODT targets must consume the ODT order while keeping WPS illegal."""
        import docwen_plugin_layout.to_document.converter as converter
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_layout import LayoutPlugin

        observed_priority: list[str] = []
        observed_candidates: set[str] = set()

        def fake_convert_with_backend_priority(
            input_path,
            output_path,
            *,
            source_format,
            backend_priority,
            com_candidates,
            libreoffice_format,
            cancel=None,
            failure_subject,
        ):
            del input_path, cancel, failure_subject
            assert source_format == "pdf"
            observed_priority.extend(backend_priority)
            observed_candidates.update(com_candidates)
            assert libreoffice_format == "odt"
            Path(output_path).write_bytes(b"odt-by-office")
            return BridgeResult(True, output_path=str(output_path), backend="test Word")

        monkeypatch.setattr(converter, "convert_with_backend_priority", fake_convert_with_backend_priority)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "odt",
                config_values={
                    "software": {"special_conversions": {"odt": ["libreoffice", "msoffice_word", "wps_writer"]}}
                },
            )
            result = LayoutPlugin().convert(context)

        assert result.success is True, f"unexpected error: {result.error}"
        assert observed_priority == ["libreoffice", "msoffice_word", "wps_writer"]
        assert observed_candidates == {"msoffice_word"}

    @pytest.mark.parametrize(
        ("target", "media_type"),
        [
            ("doc", "application/msword"),
            ("odt", "application/vnd.oasis.opendocument.text"),
            ("rtf", "application/rtf"),
        ],
    )
    def test_layout_to_other_document_formats_use_office_bridge(
        self, sample_pdf_path: Path, monkeypatch: pytest.MonkeyPatch, target: str, media_type: str
    ) -> None:
        """PDF→DOC/ODT/RTF should use the shared Office bridge path."""
        import docwen_plugin_layout.to_document.converter as converter
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_layout import LayoutPlugin

        def fake_convert_with_backend_priority(
            input_path,
            output_path,
            *,
            source_format,
            backend_priority,
            com_candidates,
            libreoffice_format,
            cancel=None,
            failure_subject,
        ):
            del input_path, backend_priority, cancel, failure_subject
            assert source_format == "pdf"
            Path(output_path).write_bytes(f"{libreoffice_format}-by-office".encode("ascii"))
            return BridgeResult(
                success=True, output_path=str(output_path), backend=next(iter(com_candidates.values())).name
            )

        monkeypatch.setattr(converter, "convert_with_backend_priority", fake_convert_with_backend_priority)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_pdf_path), staging, target)
            result = LayoutPlugin().convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert result.artifacts[0].suggested_name == f"sample.{target}"
            assert result.artifacts[0].media_type == media_type
            assert result.metrics.extra["engine"] == "office_bridge"
