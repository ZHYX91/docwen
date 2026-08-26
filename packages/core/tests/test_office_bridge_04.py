"""Focused tests split from test_office_bridge.py."""

from __future__ import annotations

from typing import ClassVar

from ._office_bridge_support import (
    Path,
    pytest,
)

pytestmark = pytest.mark.unit


def test_cancelled_backend_remains_typed_when_workspace_cleanup_reports_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    source = tmp_path / "input.doc"
    source.write_bytes(b"legacy document")

    def cancelled(*_args: object, **_kwargs: object) -> office_bridge.BridgeResult:
        return office_bridge.BridgeResult(
            False,
            message="backend stopped",
            cancelled=True,
            error_code="OFFICE_CONVERSION_CANCELLED",
        )

    def cleanup_but_report_failure(path: Path, **_kwargs: object) -> bool:
        office_bridge.shutil.rmtree(path)
        return False

    monkeypatch.setattr(office_bridge, "_convert_with_backend_priority_canonical_input", cancelled)
    monkeypatch.setattr(office_bridge, "_remove_temp_tree", cleanup_but_report_failure)

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(tmp_path / "result.docx"),
        source_format="doc",
        backend_priority=["msoffice_word"],
        com_candidates={},
    )

    assert result.success is False
    assert result.cancelled is True
    assert result.error_code == "OFFICE_CONVERSION_CANCELLED"
    assert result.message == "backend stopped"
    assert "workspace cleanup failed" in result.cleanup_message
    assert result.cleanup_failed is True


def test_cancellation_after_backend_output_prevents_publication(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    source = tmp_path / "input.doc"
    source.write_bytes(b"legacy document")
    output = tmp_path / "result.docx"

    class _CancellationView:
        is_cancelled = False

    cancel = _CancellationView()

    def finish_then_cancel(
        _input_path: str,
        output_path: str,
        **_kwargs: object,
    ) -> office_bridge.BridgeResult:
        Path(output_path).write_bytes(b"private result")
        cancel.is_cancelled = True
        return office_bridge.BridgeResult(
            True,
            output_path=output_path,
            backend="Microsoft Word",
            attempted_backend_ids=("msoffice_word",),
            available_backend_ids=("msoffice_word",),
        )

    monkeypatch.setattr(
        office_bridge,
        "_convert_with_backend_priority_canonical_input",
        finish_then_cancel,
    )

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(output),
        source_format="doc",
        backend_priority=["msoffice_word"],
        com_candidates={},
        cancel=cancel,
    )

    assert result.success is False
    assert result.cancelled is True
    assert result.error_code == "OFFICE_CONVERSION_CANCELLED"
    assert result.backend == "Microsoft Word"
    assert result.attempted_backend_ids == ("msoffice_word",)
    assert not output.exists()
    assert result.cleanup_message == "Private Office workspace cleaned up."


def test_backend_priority_fails_closed_for_unknown_admitted_format(tmp_path: Path) -> None:
    """The bridge cannot silently recover a missing concrete source contract from a suffix."""
    from docwen_core import office_bridge

    source = tmp_path / "report.doc"
    source.write_bytes(b"legacy document")
    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(tmp_path / "result.docx"),
        source_format="unknown",
        backend_priority=[],
        com_candidates={},
    )

    assert result.success is False
    assert result.message == "Unsupported admitted source format for external Office: unknown."
    assert result.error_code == "OFFICE_SOURCE_FORMAT_UNSUPPORTED"


def test_backend_priority_fails_when_canonical_alias_cannot_be_cleaned(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A leaked or still-open external input invalidates an otherwise successful conversion."""
    from docwen_core import office_bridge

    source = tmp_path / "report.docx"
    source.write_bytes(b"%PDF-1.7\ncontent")
    output = tmp_path / "result.pdf"
    output.write_bytes(b"pre-existing output")
    candidate = office_bridge.BridgeCandidate("Microsoft Word", "Word.Application", 17, "word")

    def fake_fallback(input_path: str, output_path: str, **_kwargs: object) -> office_bridge.BridgeResult:
        assert Path(input_path).suffix == ".pdf"
        Path(output_path).write_bytes(b"converted")
        return office_bridge.BridgeResult(True, output_path=output_path, backend="Microsoft Word")

    monkeypatch.setattr(office_bridge, "_convert_with_fallback", fake_fallback)

    def cleanup_but_report_failure(path: Path, **_kwargs: object) -> bool:
        office_bridge.shutil.rmtree(path)
        return False

    monkeypatch.setattr(office_bridge, "_remove_temp_tree", cleanup_but_report_failure)

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(output),
        source_format="pdf",
        backend_priority=["msoffice_word"],
        com_candidates={"msoffice_word": candidate},
    )

    assert result.success is False
    assert result.message == "Private external Office workspace could not be cleaned up."
    assert result.error_code == "OFFICE_SNAPSHOT_CLEANUP_FAILED"
    assert result.backend == "Microsoft Word"
    assert result.attempted_backend_ids == ("msoffice_word",)
    assert result.available_backend_ids == ("msoffice_word",)
    assert "workspace cleanup failed" in result.cleanup_message
    assert result.cleanup_failed is True
    assert output.read_bytes() == b"pre-existing output"


def test_convert_with_backend_priority_stops_after_cancellation(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A cancelled configured backend chain must not attempt later candidates."""
    from docwen_core import office_bridge

    calls: list[str] = []
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
        del (
            input_path,
            output_path,
            com_candidates,
            libreoffice_format,
            cancel,
            com_timeout_s,
            libreoffice_timeout_s,
        )
        calls.append("attempt")
        return office_bridge.BridgeResult(
            False,
            message="cancelled",
            cancelled=True,
            error_code="OFFICE_CONVERSION_CANCELLED",
        )

    monkeypatch.setattr(office_bridge, "_convert_with_fallback", fake_fallback)

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(tmp_path / "output.docx"),
        source_format="doc",
        backend_priority=["msoffice_word", "libreoffice"],
        com_candidates={"msoffice_word": word},
        libreoffice_format="docx",
    )

    assert result.success is False
    assert result.message == "cancelled"
    assert result.cancelled is True
    assert result.error_code == "OFFICE_CONVERSION_CANCELLED"
    assert calls == ["attempt"]


def test_convert_with_backend_priority_reports_pre_cancel_even_without_usable_candidates(tmp_path: Path) -> None:
    """Pre-cancelled empty/unknown chains must not be misreported as backend failures."""
    from docwen_core import office_bridge

    class _CancellationView:
        is_cancelled = True

    for backend_priority in ([], ["unknown-office"]):
        result = office_bridge.convert_with_backend_priority(
            str(tmp_path / "input.doc"),
            str(tmp_path / "output.docx"),
            source_format="doc",
            backend_priority=backend_priority,
            com_candidates={},
            libreoffice_format="docx",
            cancel=_CancellationView(),
        )

        assert result.success is False
        assert result.message == "cancelled"


def test_convert_with_backend_priority_shares_one_com_timeout_budget(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Configured COM candidates share one budget while preserving their order."""
    from docwen_core import office_bridge

    observed_timeouts: list[float] = []
    monotonic_values = iter([100.0, 110.0, 110.0, 115.0, 115.0])
    source = tmp_path / "input.doc"
    source.write_bytes(b"legacy document")

    def fake_fallback(
        input_path: str,
        output_path: str,
        *,
        com_candidates: list[office_bridge.BridgeCandidate],
        libreoffice_format: str | None,
        cancel: object | None = None,
        com_timeout_s: float,
        libreoffice_timeout_s: float = office_bridge.LIBREOFFICE_CONVERSION_TIMEOUT_S,
    ) -> office_bridge.BridgeResult:
        del input_path, output_path, com_candidates, libreoffice_format, cancel, libreoffice_timeout_s
        observed_timeouts.append(com_timeout_s)
        return office_bridge.BridgeResult(False, message="COM failed")

    monkeypatch.setattr(office_bridge, "_convert_with_fallback", fake_fallback)
    monkeypatch.setattr(office_bridge.time, "monotonic", lambda: next(monotonic_values))

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(tmp_path / "output.docx"),
        source_format="doc",
        backend_priority=["wps_writer", "msoffice_word"],
        com_candidates={
            "wps_writer": office_bridge.BridgeCandidate("WPS Writer", "Kwps.Application", 12, "word"),
            "msoffice_word": office_bridge.BridgeCandidate("Microsoft Word", "Word.Application", 12, "word"),
        },
        libreoffice_format="docx",
        com_timeout_s=30.0,
    )

    assert result.success is False
    assert observed_timeouts == [pytest.approx(30.0), pytest.approx(20.0)]


def test_libreoffice_wrong_suffix_uses_canonical_private_name_and_publishes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    source = tmp_path / "ledger.ods"
    source.write_bytes(b"xlsx package bytes")
    output = tmp_path / "converted.pdf"
    observed_input: list[Path] = []
    observed_output_dirs: list[Path] = []

    def fake_run(args: list[str], **_kwargs: object) -> bool:
        private_input = Path(args[-1])
        output_dir = Path(args[args.index("--outdir") + 1])
        observed_input.append(private_input)
        observed_output_dirs.append(output_dir)
        assert private_input.suffix == ".xlsx"
        assert private_input != source.resolve()
        (output_dir / f"{private_input.stem}.pdf").write_bytes(b"converted pdf")
        return True

    monkeypatch.setattr(office_bridge, "find_soffice_path", lambda: "soffice")
    monkeypatch.setattr(office_bridge, "_run_libreoffice_process", fake_run)

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(output),
        source_format="xlsx",
        backend_priority=["libreoffice"],
        com_candidates={},
        libreoffice_format="pdf:calc_pdf_Export",
    )

    assert result.success is True
    assert result.backend == "LibreOffice"
    assert result.output_path == str(output.resolve())
    assert output.read_bytes() == b"converted pdf"
    assert source.read_bytes() == b"xlsx package bytes"
    assert observed_input and not observed_input[0].exists()
    assert observed_output_dirs and not observed_output_dirs[0].exists()


def test_libreoffice_failure_removes_private_partial_without_deleting_existing_output(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    source = tmp_path / "ledger.ods"
    source.write_bytes(b"xlsx package bytes")
    output = tmp_path / "converted.pdf"
    output.write_bytes(b"existing output")
    observed_output_dirs: list[Path] = []

    def fake_run(args: list[str], **_kwargs: object) -> bool:
        private_input = Path(args[-1])
        output_dir = Path(args[args.index("--outdir") + 1])
        observed_output_dirs.append(output_dir)
        assert private_input.suffix == ".xlsx"
        (output_dir / f"{private_input.stem}.pdf").write_bytes(b"partial pdf")
        return False

    monkeypatch.setattr(office_bridge, "find_soffice_path", lambda: "soffice")
    monkeypatch.setattr(office_bridge, "_run_libreoffice_process", fake_run)

    result = office_bridge.convert_with_backend_priority(
        str(source),
        str(output),
        source_format="xlsx",
        backend_priority=["libreoffice"],
        com_candidates={},
        libreoffice_format="pdf:calc_pdf_Export",
    )

    assert result.success is False
    assert result.cancelled is False
    assert result.error_code == "OFFICE_BACKEND_FAILED"
    assert output.read_bytes() == b"existing output"
    assert observed_output_dirs and not observed_output_dirs[0].exists()
    assert not list(tmp_path.glob(f"{office_bridge._OFFICE_WORKSPACE_PREFIX}*"))
    assert not list(tmp_path.glob(f"{office_bridge._OFFICE_PUBLICATION_PREFIX}*"))


def test_libreoffice_timeout_removes_partial_without_deleting_existing_output(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    input_path = tmp_path / "input.docx"
    output_path = tmp_path / "requested.pdf"
    input_path.write_bytes(b"docx")
    output_path.write_bytes(b"existing output")
    output_dirs: list[Path] = []

    class _FakeProcess:
        returncode: int | None = None

        def __init__(self, args: list[str], **_kwargs: object) -> None:
            output_dir = Path(args[args.index("--outdir") + 1])
            output_dirs.append(output_dir)
            (output_dir / f"{Path(args[-1]).stem}.pdf").write_bytes(b"partial")

        def poll(self) -> int | None:
            return self.returncode

        def kill(self) -> None:
            self.returncode = -9

        def wait(self, timeout: float | None = None) -> int:
            del timeout
            return int(self.returncode or 0)

        def communicate(self) -> tuple[str, str]:
            return "", ""

    monotonic_values = iter([0.0, 2.0])
    monkeypatch.setattr(office_bridge, "find_soffice_path", lambda: "soffice")
    monkeypatch.setattr(office_bridge.subprocess, "Popen", _FakeProcess)
    monkeypatch.setattr(office_bridge.time, "monotonic", lambda: next(monotonic_values, 2.0))

    result = office_bridge._try_libreoffice_conversion(
        str(input_path),
        str(output_path),
        convert_to="pdf:writer_pdf_Export",
        timeout_s=1.0,
    )

    assert result is None
    assert output_path.read_bytes() == b"existing output"
    assert output_dirs and not output_dirs[0].exists()
    assert not list(tmp_path.glob(f"{office_bridge._LIBREOFFICE_RESULT_PREFIX}*"))


def test_libreoffice_conversion_terminates_process_when_cancelled(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """LibreOffice subprocesses should not wait for the hard timeout after cancellation."""
    from docwen_core import office_bridge

    input_path = tmp_path / "input.docx"
    output_path = tmp_path / "output.pdf"
    input_path.write_bytes(b"docx")

    class _CancellationView:
        is_cancelled = False

    cancel = _CancellationView()

    class _FakeProcess:
        instances: ClassVar[list[_FakeProcess]] = []

        def __init__(self, *args: object, **kwargs: object) -> None:
            self.returncode: int | None = None
            self.terminated = False
            self.killed = False
            _FakeProcess.instances.append(self)

        def poll(self) -> int | None:
            return self.returncode

        def terminate(self) -> None:
            self.terminated = True
            self.returncode = -15

        def kill(self) -> None:
            self.killed = True
            self.returncode = -9

        def wait(self, timeout: float | None = None) -> int:
            if self.returncode is None:
                self.returncode = 0
            return self.returncode

        def communicate(self) -> tuple[str, str]:
            return "", ""

    def fake_sleep(_seconds: float) -> None:
        cancel.is_cancelled = True

    monkeypatch.setattr(office_bridge, "find_soffice_path", lambda: "soffice")
    monkeypatch.setattr(office_bridge.subprocess, "Popen", _FakeProcess)
    monkeypatch.setattr(office_bridge.time, "sleep", fake_sleep)

    result = office_bridge._try_libreoffice_conversion(
        str(input_path),
        str(output_path),
        convert_to="pdf",
        cancel=cancel,
    )

    assert result is None
    assert len(_FakeProcess.instances) == 1
    assert _FakeProcess.instances[0].terminated is True
    assert _FakeProcess.instances[0].killed is False
    assert not output_path.exists()


def test_libreoffice_profile_cleanup_retries_a_transient_windows_lock(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    profile = tmp_path / "docwen-lo-profile-transient"
    profile.mkdir()
    (profile / "extensions.pmap").write_bytes(b"locked")
    real_rmtree = office_bridge.shutil.rmtree
    attempts = 0

    def flaky_rmtree(path: Path) -> None:
        nonlocal attempts
        attempts += 1
        if attempts == 1:
            raise PermissionError(32, "file is in use", str(path / "extensions.pmap"))
        real_rmtree(path)

    monkeypatch.setattr(office_bridge.shutil, "rmtree", flaky_rmtree)

    assert office_bridge._remove_libreoffice_profile(
        profile,
        timeout_s=1.0,
        retry_interval_s=0.0,
    )
    assert attempts == 2
    assert not profile.exists()
