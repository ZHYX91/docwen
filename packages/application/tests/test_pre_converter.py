"""Configured backend-priority contracts for application pre-conversion."""

from pathlib import Path

import pytest

from docwen_core.cancellation import CancellationToken

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    ("source_format", "protected_name"),
    [
        ("doc", "input.doc"),
        ("wps", "input.doc"),
        ("rtf", "input.rtf"),
        ("odt", "input.odt"),
    ],
)
def test_pre_convert_bridges_a_content_typed_protective_copy_and_preserves_source_stem(
    tmp_path,
    monkeypatch,
    source_format: str,
    protected_name: str,
) -> None:
    from docwen_application.preconversion import pre_converter
    from docwen_core.office_bridge import BridgeResult

    source_dir = tmp_path / "source"
    staging_dir = tmp_path / "staging"
    source_dir.mkdir()
    staging_dir.mkdir()
    source = source_dir / "quarterly.disguised"
    source.write_bytes(b"ORIGINAL")
    calls: list[tuple[Path, Path]] = []

    def fake_convert(input_path, output_path, **_kwargs):
        protected = Path(input_path)
        output = Path(output_path)
        calls.append((protected, output))
        assert protected != source
        assert protected.parent == staging_dir
        assert protected.name == protected_name
        assert protected.read_bytes() == b"ORIGINAL"

        # An untrusted external backend may modify its input.  That mutation
        # must remain contained to the application-owned snapshot.
        protected.write_bytes(b"BACKEND MUTATION")
        assert source.read_bytes() == b"ORIGINAL"
        output.write_bytes(b"converted")
        return BridgeResult(True, output_path=str(output), backend="Fake Office")

    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", fake_convert)

    result = pre_converter.pre_convert(
        str(source),
        source_format,
        staging_dir=str(staging_dir),
    )

    assert isinstance(result, pre_converter.PreConversionResult)
    assert Path(result.pre_converted_path).name == "quarterly.docx"
    assert calls == [(staging_dir / protected_name, staging_dir / "quarterly.docx")]
    assert source.read_bytes() == b"ORIGINAL"


def test_pre_convert_cancelled_before_copy_skips_copy_and_bridge(tmp_path, monkeypatch) -> None:
    from docwen_application.preconversion import pre_converter
    from docwen_core.office_bridge import BridgeResult

    source = tmp_path / "legacy.doc"
    source.write_bytes(b"legacy")
    staging_dir = tmp_path / "staging"
    staging_dir.mkdir()
    cancel_owner = CancellationToken()
    cancel = cancel_owner.view()
    cancel_owner.cancel()
    copy_calls: list[tuple[object, ...]] = []
    bridge_calls: list[tuple[object, ...]] = []

    def fake_convert(input_path, output_path, **_kwargs):
        bridge_calls.append((input_path, output_path))
        Path(output_path).write_bytes(b"converted")
        return BridgeResult(True, output_path=output_path, backend="Fake Office")

    monkeypatch.setattr(
        pre_converter,
        "_copy_snapshot_stream",
        lambda *args, **_kwargs: copy_calls.append(args),
    )
    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", fake_convert)

    outcome = pre_converter.pre_convert(
        str(source),
        "doc",
        staging_dir=str(staging_dir),
        cancel=cancel,
    )

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.cancelled is True
    assert outcome.error_type == "cancelled"
    assert copy_calls == []
    assert bridge_calls == []


def test_pre_convert_cancelled_after_copy_skips_bridge(tmp_path, monkeypatch) -> None:
    from docwen_application.preconversion import pre_converter
    from docwen_core.office_bridge import BridgeResult

    source = tmp_path / "legacy.doc"
    source.write_bytes(b"legacy")
    staging_dir = tmp_path / "staging"
    staging_dir.mkdir()
    cancel_owner = CancellationToken()
    cancel = cancel_owner.view()
    real_copy = pre_converter._copy_snapshot_stream
    bridge_calls: list[tuple[object, ...]] = []

    def copy_then_cancel(source_stream, destination_stream, token):
        copied = real_copy(source_stream, destination_stream, token)
        cancel_owner.cancel()
        return copied

    def fake_convert(input_path, output_path, **_kwargs):
        bridge_calls.append((input_path, output_path))
        Path(output_path).write_bytes(b"converted")
        return BridgeResult(True, output_path=output_path, backend="Fake Office")

    monkeypatch.setattr(pre_converter, "_copy_snapshot_stream", copy_then_cancel)
    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", fake_convert)

    outcome = pre_converter.pre_convert(
        str(source),
        "doc",
        staging_dir=str(staging_dir),
        cancel=cancel,
    )

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.cancelled is True
    assert outcome.error_type == "cancelled"
    assert not (staging_dir / "input.doc").exists()
    assert bridge_calls == []


def test_pre_convert_copy_failure_is_structured_and_skips_bridge(tmp_path, monkeypatch) -> None:
    from docwen_application.preconversion import pre_converter
    from docwen_core.office_bridge import BridgeResult

    source = tmp_path / "legacy.doc"
    source.write_bytes(b"legacy")
    staging_dir = tmp_path / "staging"
    staging_dir.mkdir()
    bridge_calls: list[tuple[object, ...]] = []

    def fail_copy(*_args, **_kwargs):
        raise PermissionError("source is locked")

    def fake_convert(input_path, output_path, **_kwargs):
        bridge_calls.append((input_path, output_path))
        Path(output_path).write_bytes(b"converted")
        return BridgeResult(True, output_path=output_path, backend="Fake Office")

    monkeypatch.setattr(pre_converter, "_copy_snapshot_stream", fail_copy)
    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", fake_convert)

    outcome = pre_converter.pre_convert(str(source), "doc", staging_dir=str(staging_dir))

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.cancelled is False
    assert outcome.error_type == "conversion_failed"
    assert outcome.diagnostic_code == "PRECONVERSION_INPUT_COPY_FAILED"
    assert "protective input copy" in outcome.message
    assert "source is locked" in outcome.message
    assert bridge_calls == []


def test_pre_convert_concurrent_cancel_wins_over_copy_failure(tmp_path, monkeypatch) -> None:
    from docwen_application.preconversion import pre_converter

    source = tmp_path / "legacy.doc"
    source.write_bytes(b"legacy")
    staging_dir = tmp_path / "staging"
    staging_dir.mkdir()
    cancel_owner = CancellationToken()
    cancel = cancel_owner.view()
    bridge_calls: list[tuple[object, ...]] = []

    def cancel_and_fail(*_args, **_kwargs):
        cancel_owner.cancel()
        raise PermissionError("source became unavailable")

    monkeypatch.setattr(pre_converter, "_copy_snapshot_stream", cancel_and_fail)
    monkeypatch.setattr(
        pre_converter,
        "convert_with_backend_priority",
        lambda *args, **_kwargs: bridge_calls.append(args),
    )

    outcome = pre_converter.pre_convert(
        str(source),
        "doc",
        staging_dir=str(staging_dir),
        cancel=cancel,
    )

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.cancelled is True
    assert outcome.error_type == "cancelled"
    assert bridge_calls == []


def test_pre_convert_rejects_hub_source_without_copy_or_bridge(tmp_path, monkeypatch) -> None:
    from docwen_application.preconversion import pre_converter

    source = tmp_path / "already.docx"
    source.write_bytes(b"docx")
    copy_calls: list[tuple[object, ...]] = []
    bridge_calls: list[tuple[object, ...]] = []

    monkeypatch.setattr(
        pre_converter,
        "_copy_snapshot_stream",
        lambda *args, **_kwargs: copy_calls.append(args),
    )
    monkeypatch.setattr(
        pre_converter,
        "convert_with_backend_priority",
        lambda *args, **_kwargs: bridge_calls.append(args),
    )

    outcome = pre_converter.pre_convert(str(source), "docx", staging_dir=str(tmp_path))

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.cancelled is False
    assert outcome.error_type == "dependency_missing"
    assert outcome.message == "Unsupported pre-conversion source format: docx"
    assert copy_calls == []
    assert bridge_calls == []


def test_pre_convert_honors_configured_word_priority(tmp_path, monkeypatch) -> None:
    from docwen_application.preconversion import pre_converter
    from docwen_core.office_bridge import BridgeResult

    source = tmp_path / "legacy.doc"
    source.write_bytes(b"legacy")
    calls: list[dict[str, object]] = []

    def fake_convert(input_path, output_path, **kwargs):
        calls.append(kwargs)
        Path(output_path).write_bytes(b"converted")
        return BridgeResult(True, output_path=output_path, backend="Microsoft Word")

    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", fake_convert)

    result = pre_converter.pre_convert(
        str(source),
        "doc",
        staging_dir=str(tmp_path),
        backend_priority=["msoffice_word", "libreoffice", "wps_writer"],
    )

    assert isinstance(result, pre_converter.PreConversionResult)
    assert Path(result.pre_converted_path).name == "legacy.docx"
    assert calls[0]["backend_priority"] == ["msoffice_word", "libreoffice", "wps_writer"]
    candidates = calls[0]["com_candidates"]
    assert isinstance(candidates, dict)
    assert set(candidates) == {"wps_writer", "msoffice_word"}


def test_pre_convert_odt_excludes_wps_even_if_configured(tmp_path, monkeypatch) -> None:
    from docwen_application.preconversion import pre_converter
    from docwen_core.office_bridge import BridgeResult

    source = tmp_path / "legacy.odt"
    source.write_bytes(b"legacy")
    calls: list[dict[str, object]] = []

    def fake_convert(input_path, output_path, **kwargs):
        calls.append(kwargs)
        Path(output_path).write_bytes(b"converted")
        return BridgeResult(True, output_path=output_path, backend="LibreOffice")

    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", fake_convert)

    result = pre_converter.pre_convert(
        str(source),
        "odt",
        staging_dir=str(tmp_path),
        backend_priority=["wps_writer", "libreoffice", "msoffice_word"],
    )

    assert isinstance(result, pre_converter.PreConversionResult)
    assert calls[0]["backend_priority"] == ["wps_writer", "libreoffice", "msoffice_word"]
    candidates = calls[0]["com_candidates"]
    assert isinstance(candidates, dict)
    assert set(candidates) == {"msoffice_word"}


def test_pre_convert_preserves_bridge_cancellation_outcome(tmp_path, monkeypatch) -> None:
    from docwen_application.preconversion import pre_converter
    from docwen_core.office_bridge import BridgeResult

    source = tmp_path / "legacy.doc"
    source.write_bytes(b"legacy")
    monkeypatch.setattr(
        pre_converter,
        "convert_with_backend_priority",
        lambda *_args, **_kwargs: BridgeResult(
            False,
            message="bridge stopped",
            cancelled=True,
            error_code="OFFICE_CONVERSION_CANCELLED",
            cleanup_message="Private Office workspace cleanup failed: test workspace",
            cleanup_failed=True,
        ),
    )

    outcome = pre_converter.pre_convert(str(source), "doc", staging_dir=str(tmp_path))

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.cancelled is True
    assert outcome.error_type == "cancelled"
    assert outcome.message == "bridge stopped"
    assert outcome.diagnostic_code == "OFFICE_CONVERSION_CANCELLED"
    assert outcome.cleanup_message == "Private Office workspace cleanup failed: test workspace"
    assert outcome.cleanup_failed is True


def test_pre_convert_preserves_bridge_failure_message(tmp_path, monkeypatch) -> None:
    from docwen_application.preconversion import pre_converter
    from docwen_core.office_bridge import BridgeResult

    source = tmp_path / "legacy.doc"
    source.write_bytes(b"legacy")
    monkeypatch.setattr(
        pre_converter,
        "convert_with_backend_priority",
        lambda *_args, **_kwargs: BridgeResult(
            False,
            message="Word failed; LibreOffice unavailable",
            attempted_backend_ids=("msoffice_word", "libreoffice"),
            error_code="OFFICE_BACKEND_FAILED",
        ),
    )

    outcome = pre_converter.pre_convert(str(source), "doc", staging_dir=str(tmp_path))

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.cancelled is False
    assert outcome.error_type == "dependency_missing"
    assert outcome.message == "Word failed; LibreOffice unavailable"


def test_pre_convert_distinguishes_installed_backend_failure_from_missing_dependency(
    tmp_path,
    monkeypatch,
) -> None:
    from docwen_application.preconversion import pre_converter
    from docwen_core.office_bridge import BridgeResult

    source = tmp_path / "legacy.doc"
    source.write_bytes(b"legacy")
    monkeypatch.setattr(
        pre_converter,
        "convert_with_backend_priority",
        lambda *_args, **_kwargs: BridgeResult(
            False,
            message="Microsoft Word opened the document but export failed",
            attempted_backend_ids=("msoffice_word",),
            available_backend_ids=("msoffice_word",),
            error_code="OFFICE_BACKEND_FAILED",
            cleanup_message="Private Office workspace cleanup failed: test workspace",
            cleanup_failed=True,
        ),
    )

    outcome = pre_converter.pre_convert(str(source), "doc", staging_dir=str(tmp_path))

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.cancelled is False
    assert outcome.error_type == "conversion_failed"
    assert outcome.message == "Microsoft Word opened the document but export failed"
    assert outcome.diagnostic_code == "OFFICE_BACKEND_FAILED"
    assert outcome.cleanup_message == "Private Office workspace cleanup failed: test workspace"
    assert outcome.cleanup_failed is True


@pytest.mark.parametrize(
    "error_code",
    [
        "",
        "OFFICE_BACKEND_FAILED",
        "OFFICE_SNAPSHOT_PREPARATION_FAILED",
        "OFFICE_BACKEND_EXCEPTION",
    ],
)
def test_pre_convert_does_not_misreport_non_inventory_bridge_failures_as_missing_dependencies(
    tmp_path,
    monkeypatch,
    error_code: str,
) -> None:
    from docwen_application.preconversion import pre_converter
    from docwen_core.office_bridge import BridgeResult

    source = tmp_path / "legacy.doc"
    source.write_bytes(b"legacy")
    monkeypatch.setattr(
        pre_converter,
        "convert_with_backend_priority",
        lambda *_args, **_kwargs: BridgeResult(
            False,
            message="bridge infrastructure failed",
            error_code=error_code,
        ),
    )

    outcome = pre_converter.pre_convert(str(source), "doc", staging_dir=str(tmp_path))

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.error_type == "conversion_failed"
    assert outcome.diagnostic_code == error_code
