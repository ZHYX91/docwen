"""Focused tests split from test_office_bridge.py."""

from __future__ import annotations

from ._office_bridge_support import (
    Path,
    os,
    pytest,
)

pytestmark = pytest.mark.unit


def test_convert_with_fallback_uses_libreoffice_after_com_candidates_fail(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """LibreOffice remains the shared fallback after Windows COM candidates fail."""
    from docwen_core import office_bridge

    input_path = tmp_path / "input.docx"
    output_path = tmp_path / "output.pdf"
    input_path.write_bytes(b"docx")
    calls: list[tuple[str, str, str]] = []

    def fake_com_conversion(*args: object, **kwargs: object) -> None:
        return None

    def fake_libreoffice_conversion(
        input_file: str,
        output_file: str,
        *,
        convert_to: str,
        cancel: object | None = None,
        timeout_s: float = office_bridge.LIBREOFFICE_CONVERSION_TIMEOUT_S,
    ) -> str:
        assert cancel is None
        assert timeout_s == 42.0
        calls.append((input_file, output_file, convert_to))
        Path(output_file).write_bytes(b"%PDF-1.4")
        return str(Path(output_file).resolve())

    monkeypatch.setattr(office_bridge, "_try_com_conversion_bounded", fake_com_conversion)
    monkeypatch.setattr(office_bridge, "_try_libreoffice_conversion", fake_libreoffice_conversion)

    result = office_bridge._convert_with_fallback(
        str(input_path),
        str(output_path),
        com_candidates=[
            office_bridge.BridgeCandidate(
                name="word",
                prog_id="Word.Application",
                save_format=17,
                app_type="word",
            )
        ],
        libreoffice_format="pdf:writer_pdf_Export",
        libreoffice_timeout_s=42.0,
    )

    assert result.success is True
    assert result.backend == "LibreOffice"
    assert result.output_path == str(output_path.resolve())
    assert calls == [(str(input_path), str(output_path), "pdf:writer_pdf_Export")]
    assert output_path.read_bytes() == b"%PDF-1.4"


def test_convert_with_fallback_shares_one_timeout_across_com_candidates(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A route with several Office candidates must have one total COM budget."""
    from docwen_core import office_bridge

    observed_timeouts: list[float] = []
    monotonic_values = iter([100.0, 101.0, 125.0])

    def fake_com_conversion(*args: object, timeout_s: float, **kwargs: object) -> None:
        observed_timeouts.append(timeout_s)
        return None

    monkeypatch.setattr(office_bridge.time, "monotonic", lambda: next(monotonic_values))
    monkeypatch.setattr(office_bridge, "_try_com_conversion_bounded", fake_com_conversion)

    result = office_bridge._convert_with_fallback(
        str(tmp_path / "input.doc"),
        str(tmp_path / "output.docx"),
        com_candidates=[
            office_bridge.BridgeCandidate("word", "Word.Application", 16, "word"),
            office_bridge.BridgeCandidate("wps", "Kwps.Application", 12, "word"),
        ],
        com_timeout_s=30,
    )

    assert result.success is False
    assert observed_timeouts == [pytest.approx(29.0), pytest.approx(5.0)]


def test_convert_with_fallback_accepts_runtime_cancellation_view(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The shared bridge should understand runtime cancellation views, not only threading.Event."""
    from docwen_core import office_bridge

    input_path = tmp_path / "input.docx"
    output_path = tmp_path / "output.pdf"
    input_path.write_bytes(b"docx")

    class _CancellationView:
        is_cancelled = True

    def fake_com_conversion(*args: object, **kwargs: object) -> None:
        raise AssertionError("cancelled bridge must not enter COM conversion")

    monkeypatch.setattr(office_bridge, "_try_com_conversion", fake_com_conversion)

    result = office_bridge._convert_with_fallback(
        str(input_path),
        str(output_path),
        com_candidates=[
            office_bridge.BridgeCandidate(
                name="word",
                prog_id="Word.Application",
                save_format=17,
                app_type="word",
            )
        ],
        libreoffice_format="pdf:writer_pdf_Export",
        cancel=_CancellationView(),
    )

    assert result.success is False
    assert result.message == "cancelled"
    assert not output_path.exists()


def test_convert_with_backend_priority_preserves_libreoffice_position(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Configured LibreOffice must run where it appears, not always after COM."""
    from docwen_core import office_bridge

    calls: list[tuple[list[str], str | None]] = []
    word = office_bridge.BridgeCandidate("Microsoft Word", "Word.Application", 12, "word")
    source = tmp_path / "input.doc"
    source.write_bytes(b"legacy document")

    def fake_fallback(
        input_path: str,
        output_path: str,
        *,
        com_candidates: list[office_bridge.BridgeCandidate],
        libreoffice_format: str | None,
        cancel: object | None = None,
        com_timeout_s: float = office_bridge.COM_CONVERSION_TIMEOUT_S,
        libreoffice_timeout_s: float = office_bridge.LIBREOFFICE_CONVERSION_TIMEOUT_S,
    ) -> office_bridge.BridgeResult:
        del input_path, cancel, com_timeout_s, libreoffice_timeout_s
        calls.append(([candidate.prog_id for candidate in com_candidates], libreoffice_format))
        if libreoffice_format is not None:
            return office_bridge.BridgeResult(False, message="LibreOffice failed")
        Path(output_path).write_bytes(b"docx")
        return office_bridge.BridgeResult(True, output_path=output_path, backend="Microsoft Word")

    monkeypatch.setattr(office_bridge, "_convert_with_fallback", fake_fallback)

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(tmp_path / "output.docx"),
        source_format="doc",
        backend_priority=["libreoffice", "msoffice_word"],
        com_candidates={"msoffice_word": word},
        libreoffice_format="docx",
    )

    assert result.success is True
    assert result.backend == "Microsoft Word"
    assert calls == [([], "docx"), (["Word.Application"], None)]


def test_convert_with_backend_priority_reports_unknown_and_empty_ids(tmp_path: Path) -> None:
    """Unknown or empty configured chains fail explicitly without implicit fallbacks."""
    from docwen_core import office_bridge

    source = tmp_path / "input.doc"
    source.write_bytes(b"legacy document")
    unknown = office_bridge.convert_with_backend_priority(
        str(source),
        str(tmp_path / "output.docx"),
        source_format="doc",
        backend_priority=["unknown-office"],
        com_candidates={},
        libreoffice_format="docx",
        failure_subject="Configured DOC backends",
    )
    empty = office_bridge.convert_with_backend_priority(
        str(source),
        str(tmp_path / "output.docx"),
        source_format="doc",
        backend_priority=[],
        com_candidates={},
        libreoffice_format="docx",
        failure_subject="Configured DOC backends",
    )

    assert unknown.success is False
    assert unknown.message == "Configured DOC backends did not succeed. unsupported backend id: unknown-office"
    assert empty.success is False
    assert empty.message == "Configured DOC backends did not succeed."


def test_convert_with_backend_priority_reports_available_backend_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Callers can distinguish unavailable dependencies from conversion failure."""
    from docwen_core import office_bridge

    word = office_bridge.BridgeCandidate("Microsoft Word", "Word.Application", 12, "word")
    source = tmp_path / "input.doc"
    source.write_bytes(b"legacy document")
    monkeypatch.setattr(office_bridge, "_is_com_candidate_available", lambda _candidate: True)
    monkeypatch.setattr(
        office_bridge,
        "_convert_with_fallback",
        lambda *args, **kwargs: office_bridge.BridgeResult(False, message="SaveAs failed"),
    )

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(tmp_path / "output.docx"),
        source_format="doc",
        backend_priority=["msoffice_word"],
        com_candidates={"msoffice_word": word},
        libreoffice_format="docx",
    )

    assert result.success is False
    assert result.attempted_backend_ids == ("msoffice_word",)
    assert result.available_backend_ids == ("msoffice_word",)


def test_backend_priority_materializes_misleading_suffix_with_admitted_format(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """External Office must see the content-derived format, not the user's suffix."""
    from docwen_core import office_bridge

    source = tmp_path / "report.docx"
    source.write_bytes(b"%PDF-1.7\ncontent")
    output = tmp_path / "result.pdf"
    observed_inputs: list[Path] = []
    candidate = office_bridge.BridgeCandidate("Microsoft Word", "Word.Application", 17, "word")

    def fake_fallback(input_path: str, output_path: str, **_kwargs: object) -> office_bridge.BridgeResult:
        canonical_input = Path(input_path)
        observed_inputs.append(canonical_input)
        assert canonical_input.suffix == ".pdf"
        assert canonical_input.read_bytes() == source.read_bytes()
        canonical_input.write_bytes(b"office repaired its private copy")
        Path(output_path).write_bytes(b"converted")
        return office_bridge.BridgeResult(True, output_path=output_path, backend="Microsoft Word")

    monkeypatch.setattr(office_bridge, "_convert_with_fallback", fake_fallback)

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(output),
        source_format="pdf",
        backend_priority=["msoffice_word"],
        com_candidates={"msoffice_word": candidate},
    )

    assert result.success is True
    assert observed_inputs and observed_inputs[0] != source
    assert not observed_inputs[0].exists()
    assert source.read_bytes() == b"%PDF-1.7\ncontent"
    assert output.read_bytes() == b"converted"


def test_backend_priority_snapshots_even_a_correctly_named_input(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A correct suffix does not grant an external process access to the source."""
    from docwen_core import office_bridge

    source = tmp_path / "report.doc"
    source.write_bytes(b"legacy document")
    output = tmp_path / "result.docx"
    observed_inputs: list[Path] = []
    candidate = office_bridge.BridgeCandidate("Microsoft Word", "Word.Application", 16, "word")

    def fake_fallback(input_path: str, output_path: str, **_kwargs: object) -> office_bridge.BridgeResult:
        private_input = Path(input_path)
        observed_inputs.append(private_input)
        private_input.write_bytes(b"backend repaired private input")
        Path(output_path).write_bytes(b"converted")
        return office_bridge.BridgeResult(True, output_path=output_path, backend="Microsoft Word")

    monkeypatch.setattr(office_bridge, "_convert_with_fallback", fake_fallback)

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(output),
        source_format="doc",
        backend_priority=["msoffice_word"],
        com_candidates={"msoffice_word": candidate},
    )

    assert result.success is True
    assert len(observed_inputs) == 1
    assert observed_inputs[0] != source.resolve()
    assert observed_inputs[0].suffix == ".doc"
    assert not observed_inputs[0].exists()
    assert source.read_bytes() == b"legacy document"


def test_backend_priority_rejects_input_output_identity(tmp_path: Path) -> None:
    from docwen_core import office_bridge

    source = tmp_path / "same.doc"
    source.write_bytes(b"legacy document")

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(source),
        source_format="doc",
        backend_priority=[],
        com_candidates={},
    )

    assert result.success is False
    assert result.error_code == "OFFICE_INPUT_OUTPUT_CONFLICT"
    assert result.cleanup_message == "No private Office workspace was created."
    assert source.read_bytes() == b"legacy document"
    assert not list(tmp_path.glob(f"{office_bridge._OFFICE_WORKSPACE_PREFIX}*"))


def test_backend_priority_cancels_while_copying_private_input(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    source = tmp_path / "large.doc"
    source.write_bytes(b"0123456789abcdef")

    class _CancellationView:
        calls = 0

        @property
        def is_cancelled(self) -> bool:
            self.calls += 1
            return self.calls >= 4

    cancel = _CancellationView()
    monkeypatch.setattr(office_bridge, "_OFFICE_SNAPSHOT_CHUNK_SIZE", 4)

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(tmp_path / "result.docx"),
        source_format="doc",
        backend_priority=[],
        com_candidates={},
        cancel=cancel,
    )

    assert result.success is False
    assert result.cancelled is True
    assert result.error_code == "OFFICE_CONVERSION_CANCELLED"
    assert result.cleanup_message == "Private Office workspace cleaned up."
    assert not list(tmp_path.glob(f"{office_bridge._OFFICE_WORKSPACE_PREFIX}*"))


def test_backend_priority_rejects_source_identity_change_during_copy(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    source = tmp_path / "changing.doc"
    source.write_bytes(b"legacy document")
    real_identity = office_bridge._source_identity
    observations = 0

    def changing_identity(file_stat: os.stat_result) -> office_bridge._SourceIdentity:
        nonlocal observations
        observations += 1
        identity = real_identity(file_stat)
        if observations == 3:
            return office_bridge._SourceIdentity(
                device=identity.device,
                inode=identity.inode,
                size=identity.size,
                modified_ns=identity.modified_ns + 1,
            )
        return identity

    monkeypatch.setattr(office_bridge, "_source_identity", changing_identity)

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(tmp_path / "result.docx"),
        source_format="doc",
        backend_priority=[],
        com_candidates={},
    )

    assert result.success is False
    assert result.error_code == "OFFICE_SOURCE_CHANGED"
    assert result.cleanup_message == "Private Office workspace cleaned up."
    assert not list(tmp_path.glob(f"{office_bridge._OFFICE_WORKSPACE_PREFIX}*"))


def test_backend_priority_reports_copy_exception_and_successful_cleanup(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    source = tmp_path / "input.doc"
    source.write_bytes(b"legacy document")

    def fail_copy(*_args: object, **_kwargs: object) -> None:
        raise PermissionError("copy denied")

    monkeypatch.setattr(office_bridge, "_copy_private_office_input", fail_copy)

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(tmp_path / "result.docx"),
        source_format="doc",
        backend_priority=[],
        com_candidates={},
    )

    assert result.success is False
    assert result.error_code == "OFFICE_SNAPSHOT_PREPARATION_FAILED"
    assert "copy denied" in result.message
    assert result.cleanup_message == "Private Office workspace cleaned up."
    assert not list(tmp_path.glob(f"{office_bridge._OFFICE_WORKSPACE_PREFIX}*"))


def test_backend_priority_reports_workspace_creation_exception_without_fake_cleanup(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    source = tmp_path / "input.doc"
    source.write_bytes(b"legacy document")
    monkeypatch.setattr(
        office_bridge.tempfile,
        "mkdtemp",
        lambda **_kwargs: (_ for _ in ()).throw(PermissionError("workspace denied")),
    )

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(tmp_path / "result.docx"),
        source_format="doc",
        backend_priority=[],
        com_candidates={},
    )

    assert result.success is False
    assert result.error_code == "OFFICE_SNAPSHOT_PREPARATION_FAILED"
    assert "workspace denied" in result.message
    assert result.cleanup_message == "No private Office workspace was created."


def test_backend_exception_retains_error_when_workspace_cleanup_reports_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    source = tmp_path / "input.doc"
    source.write_bytes(b"legacy document")

    def explode(*_args: object, **_kwargs: object) -> office_bridge.BridgeResult:
        raise RuntimeError("backend exploded")

    def cleanup_but_report_failure(path: Path, **_kwargs: object) -> bool:
        office_bridge.shutil.rmtree(path)
        return False

    monkeypatch.setattr(office_bridge, "_convert_with_backend_priority_canonical_input", explode)
    monkeypatch.setattr(office_bridge, "_remove_temp_tree", cleanup_but_report_failure)

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(tmp_path / "result.docx"),
        source_format="doc",
        backend_priority=["msoffice_word"],
        com_candidates={},
    )

    assert result.success is False
    assert result.error_code == "OFFICE_BACKEND_EXCEPTION"
    assert "backend exploded" in result.message
    assert "workspace cleanup failed" in result.cleanup_message
    assert result.cleanup_failed is True
