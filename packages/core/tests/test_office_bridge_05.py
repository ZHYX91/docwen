"""Focused tests split from test_office_bridge.py."""

from __future__ import annotations

from ._office_bridge_support import (
    Path,
    _write_formula_writer_container,
    os,
    pytest,
    tempfile,
    urllib_parse,
)

pytestmark = pytest.mark.unit


def test_libreoffice_profile_cleanup_contains_a_persistent_windows_lock(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    profile = tmp_path / "docwen-lo-profile-persistent"
    profile.mkdir()

    def locked_rmtree(path: Path) -> None:
        raise PermissionError(32, "file is in use", str(path / "extensions.pmap"))

    monkeypatch.setattr(office_bridge.shutil, "rmtree", locked_rmtree)

    assert not office_bridge._remove_libreoffice_profile(
        profile,
        timeout_s=0.0,
        retry_interval_s=0.0,
    )
    assert profile.is_dir()


@pytest.mark.parametrize("document_format", ["docx", "odt"])
def test_writer_formula_pdf_uses_times_fallback_only_for_compatible_inputs(
    tmp_path: Path,
    document_format: str,
) -> None:
    from docwen_core import office_bridge

    compatible = tmp_path / f"compatible.{document_format}"
    explicit_liberation = tmp_path / f"explicit-liberation.{document_format}"
    _write_formula_writer_container(compatible, document_format=document_format)
    _write_formula_writer_container(
        explicit_liberation,
        document_format=document_format,
        font_name="Liberation Serif",
    )

    assert office_bridge._writer_conversion_needs_times_formula_fallback(compatible, "pdf:writer_pdf_Export")
    assert office_bridge._writer_conversion_needs_times_formula_fallback(compatible, "odt")
    assert not office_bridge._writer_conversion_needs_times_formula_fallback(compatible, "docx")
    assert not office_bridge._writer_conversion_needs_times_formula_fallback(
        explicit_liberation, "pdf:writer_pdf_Export"
    )


def test_libreoffice_conversion_seeds_owned_formula_font_profile(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    input_path = tmp_path / "input.docx"
    output_path = tmp_path / "requested.pdf"
    _write_formula_writer_container(input_path, document_format="docx")
    observed_profiles: list[Path] = []

    def fake_run(args: list[str], **_kwargs: object) -> bool:
        profile_uri = args[1].partition("=")[2]
        parsed_profile_path = urllib_parse.unquote(urllib_parse.urlparse(profile_uri).path)
        if os.name == "nt":
            parsed_profile_path = parsed_profile_path.lstrip("/")
        profile_path = Path(parsed_profile_path)
        registry = profile_path / "user" / "registrymodifications.xcu"
        registry_text = registry.read_text(encoding="utf-8")
        assert "Liberation Serif" in registry_text
        assert "Times New Roman" in registry_text
        assert "<value>true</value>" in registry_text
        observed_profiles.append(profile_path)
        output_dir = Path(args[args.index("--outdir") + 1])
        (output_dir / f"{Path(args[-1]).stem}.pdf").write_bytes(b"%PDF-1.4")
        return True

    monkeypatch.setattr(office_bridge, "find_soffice_path", lambda: "soffice")
    monkeypatch.setattr(office_bridge, "_run_libreoffice_process", fake_run)

    result = office_bridge._try_libreoffice_conversion(
        str(input_path),
        str(output_path),
        convert_to="pdf:writer_pdf_Export",
    )

    assert result == str(output_path.resolve())
    assert len(observed_profiles) == 1
    assert not observed_profiles[0].exists()


def test_libreoffice_conversion_fails_closed_when_owned_profile_cannot_cleanup(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core import office_bridge

    input_path = tmp_path / "input.docx"
    output_path = tmp_path / "requested.pdf"
    input_path.write_bytes(b"docx")
    cleanup_calls: list[Path] = []

    class _FakeProcess:
        returncode = 0

        def __init__(self, args: list[str], **_kwargs: object) -> None:
            output_dir = Path(args[args.index("--outdir") + 1])
            (output_dir / f"{Path(args[-1]).stem}.pdf").write_bytes(b"pdf")

        def poll(self) -> int:
            return self.returncode

        def communicate(self) -> tuple[str, str]:
            return "", ""

    def failed_cleanup(path: Path, **_kwargs: object) -> bool:
        cleanup_calls.append(path)
        office_bridge.shutil.rmtree(path)
        return False

    monkeypatch.setattr(office_bridge, "find_soffice_path", lambda: "soffice")
    monkeypatch.setattr(office_bridge.subprocess, "Popen", _FakeProcess)
    monkeypatch.setattr(
        office_bridge,
        "_remove_libreoffice_profile",
        failed_cleanup,
        raising=False,
    )

    result = office_bridge._try_libreoffice_conversion(
        str(input_path),
        str(output_path),
        convert_to="pdf:writer_pdf_Export",
    )

    assert result is None
    assert len(cleanup_calls) == 1
    assert not output_path.exists()
    assert not list(tmp_path.glob(f"{office_bridge._LIBREOFFICE_RESULT_PREFIX}*"))


def test_libreoffice_conversion_uses_extension_before_filter_name(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """LibreOffice uses a private profile and the filter's leading extension."""
    from docwen_core import office_bridge

    input_path = tmp_path / "input.docx"
    output_path = tmp_path / "requested.pdf"
    input_path.write_bytes(b"docx")
    commands: list[list[str]] = []
    profile_paths: list[Path] = []

    class _FakeProcess:
        returncode = 0

        def __init__(self, args: list[str], **kwargs: object) -> None:
            commands.append(args)
            self.args = args
            profile_uri = args[1].partition("=")[2]
            parsed_profile_path = urllib_parse.unquote(urllib_parse.urlparse(profile_uri).path)
            if os.name == "nt":
                parsed_profile_path = parsed_profile_path.lstrip("/")
            profile_path = Path(parsed_profile_path)
            assert profile_path.is_dir()
            profile_paths.append(profile_path)

        def poll(self) -> int:
            return self.returncode

        def communicate(self) -> tuple[str, str]:
            output_dir = Path(self.args[self.args.index("--outdir") + 1])
            (output_dir / f"{Path(self.args[-1]).stem}.pdf").write_bytes(b"%PDF-1.4")
            return "", ""

    monkeypatch.setattr(office_bridge, "find_soffice_path", lambda: "soffice")
    monkeypatch.setattr(office_bridge.subprocess, "Popen", _FakeProcess)

    result = office_bridge._try_libreoffice_conversion(
        str(input_path),
        str(output_path),
        convert_to="pdf:writer_pdf_Export",
    )

    assert result == str(output_path.resolve())
    assert output_path.read_bytes() == b"%PDF-1.4"
    assert len(commands) == 1
    assert commands[0][0] == "soffice"
    assert commands[0][1].startswith("-env:UserInstallation=file:")
    assert commands[0][2:6] == [
        "--headless",
        "--convert-to",
        "pdf:writer_pdf_Export",
        "--outdir",
    ]
    assert Path(commands[0][6]).parent == tmp_path
    assert Path(commands[0][6]).name.startswith(office_bridge._LIBREOFFICE_OUTPUT_PREFIX)
    assert commands[0][7] == str(input_path)
    assert len(profile_paths) == 1
    assert profile_paths[0].parent == Path(tempfile.gettempdir()).resolve()
    assert profile_paths[0].name.startswith(office_bridge._LIBREOFFICE_PROFILE_PREFIX)
    assert profile_paths[0].parent != output_path.parent
    assert not profile_paths[0].exists()
