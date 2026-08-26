"""Focused tests split from test_request_scoped_docx_policy.py."""

from __future__ import annotations

from ._request_scoped_docx_policy_support import (
    Any,
    Document,
    DocumentPlugin,
    FakeProgressSink,
    Path,
    _build_concurrency_probe_docx,
    _build_policy_probe_docx,
    _build_recursive_policy_probe_docx,
    _context,
    _markdown_from_result,
    _request_policy,
    build_docx_markdown_request_policy,
    pytest,
)

pytestmark = pytest.mark.unit


def test_request_syntax_reaches_sdt_and_nested_table_paths(tmp_path: Path) -> None:
    input_path = _build_recursive_policy_probe_docx(tmp_path)
    context = _context(
        tmp_path,
        input_path,
        request_id="recursive-policy",
        config=_request_policy(),
    )

    markdown = _markdown_from_result(DocumentPlugin().convert(context))

    assert "__sdt bold__" in markdown
    assert "__nested bold__" in markdown
    assert "**sdt bold**" not in markdown
    assert "**nested bold**" not in markdown
    assert not any(artifact.kind == "auxiliary" for artifact in context.workspace.registered_artifacts)


def test_converter_base64_path_consumes_explicit_request_compression_policy(
    tmp_path: Path,
) -> None:
    input_path = _build_policy_probe_docx(tmp_path, name="base64-policy.docx")
    config = _request_policy()
    config["export"]["to_md_image_extraction_mode"] = "base64"
    config["conversion"]["export"] = {
        "base64_compress_enabled": False,
        "base64_compress_threshold_kb": 1,
    }
    context = _context(
        tmp_path,
        input_path,
        request_id="base64-policy",
        config=config,
    )

    policy = build_docx_markdown_request_policy(context, {})
    assert policy.export.export_base64_compress_enabled is False
    assert policy.export.export_base64_compress_threshold_kb == 1

    markdown = _markdown_from_result(DocumentPlugin().convert(context))

    assert "data:image/png;base64," in markdown
    assert "](data:image/png;base64," in markdown
    assert "![[data:image/png;base64," not in markdown


def test_converter_ocr_sidecar_uses_request_link_without_main_title(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    input_path = _build_policy_probe_docx(tmp_path, name="ocr-policy.docx")
    config = _request_policy()
    config["export"]["to_md_image_extraction_mode"] = "file"
    config["export"]["to_md_ocr_placement_mode"] = "image_md"
    config["conversion"]["ocr_output"] = {
        "show_blockquote_title": True,
        "blockquote_title_override_by_locale": {"zh_CN": "Request OCR"},
    }
    config["link"]["format"] = {
        "image_link_style": "markdown_embed",
        "md_file_link_style": "markdown_link",
    }
    context = _context(
        tmp_path,
        input_path,
        request_id="ocr-policy",
        config=config,
        options={"to_md_enable_ocr": True},
    )

    monkeypatch.setattr(
        ocr,
        "run_ocr_outcome",
        lambda *_args, **_kwargs: ocr.OcrOutcome(
            ocr.OcrStatus.SUCCESS,
            text="request OCR text",
        ),
    )

    markdown = _markdown_from_result(DocumentPlugin().convert(context))
    sidecar = next(
        artifact for artifact in context.workspace.registered_artifacts if artifact.metadata.get("ocr") is True
    )
    sidecar_text = Path(sidecar.staging_path).read_text(encoding="utf-8")

    assert "__img_001_ocr.md)" in markdown
    assert "![[" not in markdown
    assert "](docx-image" not in markdown
    assert "](docx-image" in sidecar_text
    assert "Request OCR" not in sidecar_text
    assert "> request OCR text" in sidecar_text


def test_converter_main_ocr_uses_request_title_policy(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    input_path = _build_policy_probe_docx(tmp_path, name="ocr-main-policy.docx")
    config = _request_policy()
    config["export"]["to_md_image_extraction_mode"] = "file"
    config["export"]["to_md_ocr_placement_mode"] = "main_md"
    config["conversion"]["ocr_output"] = {
        "show_blockquote_title": True,
        "blockquote_title_override_by_locale": {"en_US": "🖼️ **Request OCR**:"},
    }
    context = _context(
        tmp_path,
        input_path,
        request_id="ocr-main-policy",
        config=config,
        options={"to_md_enable_ocr": True, "locale": "en_US"},
        ocr_blockquote_title="🖼️ **Request OCR**:",
    )

    monkeypatch.setattr(
        ocr,
        "run_ocr_outcome",
        lambda *_args, **_kwargs: ocr.OcrOutcome(
            ocr.OcrStatus.SUCCESS,
            text="request OCR text",
        ),
    )

    markdown = _markdown_from_result(DocumentPlugin().convert(context))

    assert "> 🖼️ **Request OCR**:" in markdown
    assert "> request OCR text" in markdown
    assert markdown.count("](docx-image") == 1
    assert not any(artifact.metadata.get("ocr") is True for artifact in context.workspace.registered_artifacts)


def test_converter_main_ocr_omits_disabled_title(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    input_path = _build_policy_probe_docx(tmp_path, name="ocr-title-disabled.docx")
    config = _request_policy()
    config["export"]["to_md_image_extraction_mode"] = "file"
    config["export"]["to_md_ocr_placement_mode"] = "main_md"
    config["conversion"]["ocr_output"] = {
        "show_blockquote_title": False,
        "blockquote_title_override_by_locale": {"zh_CN": "must not appear"},
    }
    context = _context(
        tmp_path,
        input_path,
        request_id="ocr-title-disabled",
        config=config,
        options={"to_md_enable_ocr": True},
    )

    monkeypatch.setattr(
        ocr,
        "run_ocr_outcome",
        lambda *_args, **_kwargs: ocr.OcrOutcome(
            ocr.OcrStatus.SUCCESS,
            text="plain OCR text",
        ),
    )

    markdown = _markdown_from_result(DocumentPlugin().convert(context))

    assert "must not appear" not in markdown
    assert "> plain OCR text" in markdown


def test_ocr_title_override_uses_request_locale(tmp_path: Path) -> None:
    input_path = _build_concurrency_probe_docx(tmp_path, name="ocr-locale.docx")
    config = _request_policy()
    config["conversion"]["ocr_output"] = {
        "show_blockquote_title": True,
        "blockquote_title_override_by_locale": {
            "zh_CN": "中文标题",
            "en_US": "English title",
        },
    }
    context = _context(
        tmp_path,
        input_path,
        request_id="ocr-locale",
        config=config,
        options={"locale": "en_US"},
        ocr_blockquote_title="English title",
    )

    policy = build_docx_markdown_request_policy(context, context.request.options)

    assert policy.ocr_blockquote_title == "English title"


def test_converter_main_ocr_uses_runtime_injected_localized_fallback(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text.ocr as ocr

    input_path = _build_policy_probe_docx(tmp_path, name="ocr-fallback.docx")
    config = _request_policy()
    config["export"]["to_md_image_extraction_mode"] = "file"
    config["export"]["to_md_ocr_placement_mode"] = "main_md"
    config["conversion"]["ocr_output"] = {
        "show_blockquote_title": True,
        "blockquote_title_override_by_locale": {},
    }
    context = _context(
        tmp_path,
        input_path,
        request_id="ocr-fallback",
        config=config,
        options={"to_md_enable_ocr": True, "locale": "en_US"},
        ocr_blockquote_title="🖼️ **Image OCR**:",
    )

    monkeypatch.setattr(
        ocr,
        "run_ocr_outcome",
        lambda *_args, **_kwargs: ocr.OcrOutcome(
            ocr.OcrStatus.SUCCESS,
            text="fallback OCR text",
        ),
    )

    markdown = _markdown_from_result(DocumentPlugin().convert(context))

    assert "> 🖼️ **Image OCR**:" in markdown
    assert "> fallback OCR text" in markdown


def test_table_images_and_ocr_remain_in_exact_cells_across_nested_content(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from hashlib import sha256

    from PIL import Image

    import docwen_core.text.ocr as ocr

    image_paths: list[Path] = []
    for index, color in enumerate(((255, 0, 0), (0, 255, 0), (0, 0, 255), (255, 255, 0)), start=1):
        path = tmp_path / f"cell-{index}.png"
        Image.new("RGB", (2, 2), color=color).save(path)
        image_paths.append(path)

    doc = Document()
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = "H|1"
    table.cell(0, 1).text = "H2"

    left_first = table.cell(1, 0).paragraphs[0]
    left_first.text = "left-before"
    left_first.add_run().add_picture(str(image_paths[0]))
    left_second = table.cell(1, 0).add_paragraph("left-after")
    left_second.add_run().add_picture(str(image_paths[1]))
    right = table.cell(1, 1).paragraphs[0]
    right.text = "right"
    right.add_run().add_picture(str(image_paths[2]))

    merged = table.cell(2, 0).merge(table.cell(2, 1))
    merged.text = "merged"
    nested = merged.add_table(rows=1, cols=1)
    nested_para = nested.cell(0, 0).paragraphs[0]
    nested_para.text = "nested"
    nested_para.add_run().add_picture(str(image_paths[3]))

    input_path = tmp_path / "table-cell-images.docx"
    doc.save(str(input_path))
    config = _request_policy()
    config["export"]["to_md_image_extraction_mode"] = "file"
    config["export"]["to_md_ocr_placement_mode"] = "main_md"
    context = _context(
        tmp_path,
        input_path,
        request_id="table-cell-images",
        config=config,
        options={"to_md_keep_images": True, "to_md_enable_ocr": True},
    )

    ocr_texts = iter(("OCR-left-first", "OCR-left-second", "OCR-right", "OCR-nested"))
    monkeypatch.setattr(
        ocr,
        "run_ocr_outcome",
        lambda *_args, **_kwargs: ocr.OcrOutcome(ocr.OcrStatus.SUCCESS, text=next(ocr_texts)),
    )

    result = DocumentPlugin().convert(context)
    markdown = _markdown_from_result(result)
    image_artifacts = [artifact for artifact in result.artifacts if artifact.kind == "image"]

    assert len(image_artifacts) == 4
    assert len({sha256(Path(artifact.staging_path).read_bytes()).hexdigest() for artifact in image_artifacts}) == 4
    assert next(artifact for artifact in result.artifacts if artifact.is_primary).metadata["image_count"] == 4
    names = [artifact.suggested_name for artifact in image_artifacts]
    table_lines = [line for line in markdown.splitlines() if line.startswith("|")]
    assert any(r"H\|1" in line for line in table_lines)

    data_line = next(line for line in table_lines if "left-before" in line)
    first_separator = data_line.find(" | ", 2)
    left_cell = data_line[:first_separator]
    right_cell = data_line[first_separator + 3 :]
    assert names[0] in left_cell and names[1] in left_cell
    assert names[2] not in left_cell
    assert left_cell.index(names[0]) < left_cell.index("left-after") < left_cell.index(names[1])
    assert "OCR-left-first" in left_cell and "OCR-left-second" in left_cell
    assert names[2] in right_cell and "OCR-right" in right_cell
    assert names[0] not in right_cell and names[1] not in right_cell

    nested_line = next(line for line in table_lines if "nested" in line)
    assert names[3] in nested_line
    assert "OCR-nested" in nested_line
    assert all(not any(name in line for name in names) for line in markdown.splitlines() if not line.startswith("|"))


def test_document_all_ocr_outcomes_warn_and_continue_later_images(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Every OCR outcome warns without losing base Markdown or later OCR."""
    from PIL import Image

    import docwen_core.text.ocr as ocr

    doc = Document()
    doc.add_paragraph("document base content")
    for index in range(6):
        image_path = tmp_path / f"document-ocr-{index}.png"
        Image.new("RGB", (2, 2), color=(index * 20, 0, 0)).save(image_path)
        doc.add_paragraph(f"image {index}").add_run().add_picture(str(image_path))
    input_path = tmp_path / "document-ocr-best-effort.docx"
    doc.save(str(input_path))

    config = _request_policy()
    config["export"]["to_md_image_extraction_mode"] = "file"
    context = _context(
        tmp_path,
        input_path,
        request_id="document-ocr-best-effort",
        config=config,
        options={
            "to_md_keep_images": True,
            "to_md_enable_ocr": True,
            "ocr_placement": "main_md",
        },
    )

    class _RecordingProgress(FakeProgressSink):
        def __init__(self) -> None:
            super().__init__()
            self.diagnostics: list[tuple[str, str, str, str]] = []

        def report_diagnostic(
            self,
            level: str,
            message: str,
            code: str = "",
            location: str = "",
        ) -> None:
            self.diagnostics.append((level, message, code, location))

    progress = _RecordingProgress()
    context._progress = progress
    outcomes = iter(
        [
            ocr.OcrOutcome(ocr.OcrStatus.UNAVAILABLE, message="private unavailable detail"),
            ocr.OcrOutcome(ocr.OcrStatus.MODEL_MISSING, message="private model path"),
            ocr.OcrOutcome(ocr.OcrStatus.INITIALIZATION_FAILED, message="private init detail"),
            ocr.OcrOutcome(ocr.OcrStatus.RECOGNITION_FAILED, message="private recognition detail"),
            ocr.OcrOutcome(ocr.OcrStatus.NO_TEXT),
            ocr.OcrOutcome(ocr.OcrStatus.SUCCESS, text="later document OCR"),
        ]
    )
    calls: list[str] = []

    def _run_ocr_outcome(path: str, **_kwargs: Any) -> ocr.OcrOutcome:
        calls.append(str(path))
        return next(outcomes)

    monkeypatch.setattr(ocr, "run_ocr_outcome", _run_ocr_outcome)

    markdown = _markdown_from_result(DocumentPlugin().convert(context))

    assert "document base content" in markdown
    assert "> later document OCR" in markdown
    assert len(calls) == 6
    assert len(progress.diagnostics) == 6
    assert [diagnostic[0] for diagnostic in progress.diagnostics] == ["warning"] * 6
    assert [diagnostic[2] for diagnostic in progress.diagnostics] == ["OCR-BEST-EFFORT"] * 6
    assert [
        message.split("status=", 1)[1].split(";", 1)[0] for _level, message, _code, _location in progress.diagnostics
    ] == [
        "unavailable",
        "model_missing",
        "initialization_failed",
        "recognition_failed",
        "no_text",
        "success",
    ]
    assert all(location for _level, _message, _code, location in progress.diagnostics)
    assert all("private" not in message for _level, message, _code, _location in progress.diagnostics)
