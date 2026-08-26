"""Source-tree CLI/runtime smoke tests for export default consumption."""

from __future__ import annotations

import argparse
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_CONFIGS = Path(__file__).resolve().parents[4] / "configs"


def _make_execution_args(**overrides: object) -> argparse.Namespace:
    ns = argparse.Namespace()
    ns.command = "convert"
    ns.json = False
    ns.quiet = True
    ns.verbose = False
    ns.timing = False
    ns.batch = False
    ns.jobs = 1
    ns.continue_on_error = False
    ns.output_path = None
    ns.output_dir = None
    ns.dry_run = False
    ns.to = "md"
    ns.template = None
    ns.check = []
    ns.extract_img = False
    ns.no_extract_img = False
    ns.ocr = False
    ns.image_mode = None
    ns.ocr_placement = None
    ns.action = ""
    ns.clean_numbering = None
    ns.add_numbering = None
    ns.heading_numbering_render_mode = None
    ns.heading_merge_mode = None
    ns.files = []
    ns.file = None
    ns.pages = None
    ns.dpi = None
    ns.mode = None
    ns.keep_alpha = False
    ns.ocr_language = None
    for key, value in overrides.items():
        setattr(ns, key, value)
    return ns


def _write_png(path: Path) -> None:
    from PIL import Image, ImageDraw

    image = Image.new("RGB", (120, 48), "white")
    draw = ImageDraw.Draw(image)
    draw.text((8, 16), "OCR", fill="black")
    path.parent.mkdir(parents=True, exist_ok=True)
    image.save(path)


def _single_image_markdown_node(output_dir: Path) -> Path:
    candidates = list(output_dir.glob("source_*_fromPng/source_*_fromPng.md"))
    assert len(candidates) == 1
    primary = candidates[0]
    assert primary.parent.stem == primary.stem
    return primary


def test_cli_image_without_ocr_flag_disables_ocr_even_when_config_default_is_enabled(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """CLI --ocr is opt-in; omitted flag must not inherit image OCR defaults."""
    import docwen_plugin_image.to_markdown.converter as image_converter
    from docwen_application.controller import ApplicationController
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_cli.commands.convert import execute_convert
    from docwen_core.text.ocr import OcrOutcome, OcrStatus

    calls: list[str] = []

    def fake_ocr(_path: str, **_kwargs: object) -> OcrOutcome:
        calls.append("called")
        return OcrOutcome(OcrStatus.SUCCESS, text="CLI SHOULD NOT OCR")

    monkeypatch.setattr(image_converter, "run_ocr_outcome", fake_ocr)

    source = tmp_path / "source.png"
    output_dir = tmp_path / "out"
    output_dir.mkdir()
    _write_png(source)

    config_port = ConfigPortAdapter(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "configs")
    assert config_port.get("image.to_md_enable_ocr") is True

    runtime_port = create_runtime_port()
    controller = ApplicationController(runtime_port=runtime_port, config_port=config_port)
    controller.start()
    try:
        args = _make_execution_args(
            files=[str(source)],
            output_dir=str(output_dir),
            to="md",
            ocr=False,
        )

        assert execute_convert(args, controller=controller) == 0
    finally:
        controller.stop()
        runtime_port.shutdown()

    primary = _single_image_markdown_node(output_dir)
    content = primary.read_text(encoding="utf-8")
    assert calls == []
    assert "CLI SHOULD NOT OCR" not in content
    assert not list(output_dir.rglob("*_ocr.md"))


def test_cli_image_ocr_uses_export_ocr_placement_default_without_flag(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """No --ocr-placement flag: runtime export.toml default controls placement."""
    import docwen_plugin_image.to_markdown.converter as image_converter
    from docwen_application.controller import ApplicationController
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_cli.commands.convert import execute_convert
    from docwen_core.text.ocr import OcrOutcome, OcrStatus

    monkeypatch.setattr(
        image_converter,
        "run_ocr_outcome",
        lambda _path, **_kwargs: OcrOutcome(OcrStatus.SUCCESS, text="CLI EXPORT DEFAULT OCR"),
    )

    source = tmp_path / "source.png"
    output_dir = tmp_path / "out"
    output_dir.mkdir()
    _write_png(source)

    config_port = ConfigPortAdapter(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "configs")
    assert config_port.get("export.to_md_ocr_placement_mode") == "main_md"

    runtime_port = create_runtime_port()
    controller = ApplicationController(runtime_port=runtime_port, config_port=config_port)
    controller.start()
    try:
        args = _make_execution_args(
            files=[str(source)],
            output_dir=str(output_dir),
            to="md",
            ocr=True,
            ocr_placement=None,
        )

        assert execute_convert(args, controller=controller) == 0
    finally:
        controller.stop()
        runtime_port.shutdown()

    primary = _single_image_markdown_node(output_dir)
    content = primary.read_text(encoding="utf-8")
    assert "CLI EXPORT DEFAULT OCR" in content
    assert "> CLI EXPORT DEFAULT OCR" in content
    assert not list(output_dir.rglob("*_ocr.md"))


def test_cli_image_ocr_uses_configured_ocr_language_default(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """No CLI OCR language flag: image.ocr_language config reaches plugin OCR."""
    import docwen_plugin_image.to_markdown.converter as image_converter
    from docwen_application.controller import ApplicationController
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_cli.commands.convert import execute_convert
    from docwen_core.text.ocr import OcrOutcome, OcrStatus

    ocr_calls: list[tuple[str | None, str]] = []

    def fake_ocr(_path: str, **kwargs: object) -> OcrOutcome:
        ocr_language = kwargs.get("ocr_language")
        assert ocr_language is None or isinstance(ocr_language, str)
        ocr_calls.append((ocr_language, str(kwargs.get("current_locale"))))
        return OcrOutcome(OcrStatus.SUCCESS, text="CLI OCR LANGUAGE DEFAULT")

    monkeypatch.setattr(image_converter, "run_ocr_outcome", fake_ocr)

    source = tmp_path / "source.png"
    output_dir = tmp_path / "out"
    output_dir.mkdir()
    _write_png(source)

    config_port = ConfigPortAdapter(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "configs")
    assert config_port.set("image.ocr_language", "japanese") is True

    runtime_port = create_runtime_port()
    controller = ApplicationController(runtime_port=runtime_port, config_port=config_port)
    controller.start()
    try:
        args = _make_execution_args(
            files=[str(source)],
            output_dir=str(output_dir),
            to="md",
            ocr=True,
            ocr_language=None,
        )

        assert execute_convert(args, controller=controller) == 0
    finally:
        controller.stop()
        runtime_port.shutdown()

    primary = _single_image_markdown_node(output_dir)
    assert "CLI OCR LANGUAGE DEFAULT" in primary.read_text(encoding="utf-8")
    assert ocr_calls == [("japanese", "zh_CN")]
