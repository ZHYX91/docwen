"""Derived Office clipboard bytes through MainWindow, worker and final output.

These source GUI regressions do not constitute native Office or packaged acceptance.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData

from docwen_application.controller import ApplicationController
from docwen_bundle.config_port import ConfigPortAdapter
from docwen_bundle.runtime_factory import create_runtime_port
from docwen_core.models.clipboard_document import (
    ClipboardImageRef,
    iter_clipboard_inlines,
    load_clipboard_document_bytes,
)
from docwen_gui.clipboard_office_provider import WORD_EMBED_SOURCE_MIME, WPS_DOCUMENT_MIME, WPS_IMAGE_DATA_MIME
from docwen_gui.main_window import MainWindow
from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = [pytest.mark.gui, pytest.mark.pr_gate, pytest.mark.release_gate]
_SAMPLES = Path(__file__).resolve().parents[4] / "tests/fixtures/files/clipboard-office"


@pytest.mark.parametrize("producer", ["word", "wps"])
def test_derived_office_clipboard_to_final_markdown(qapp, qtbot, tmp_path, producer) -> None:
    provenance = json.loads((_SAMPLES / "provenance.json").read_text(encoding="utf-8"))["files"]

    def sample(name):
        data = (_SAMPLES / name).read_bytes()
        assert len(data) == provenance[name]["bytes"]
        assert hashlib.sha256(data).hexdigest() == provenance[name]["sha256"]
        return data

    controller = ApplicationController(
        runtime_port=create_runtime_port(),
        config_port=ConfigPortAdapter(base_dir=_SAMPLES.parents[3] / "configs", user_dir=tmp_path / "configs"),
    )
    controller.start()
    window = MainWindow(
        view_model=MainWindowViewModel(controller=controller),
        clipboard_input_root=tmp_path / "clipboard-inputs",
    )
    qtbot.addWidget(window)
    window.show()
    try:
        config = controller.config_port
        assert config is not None
        assert config.set("output.directory.mode", "custom")
        assert config.set("output.directory.custom_path", str(tmp_path / "published"))
        assert config.set("output.directory.create_date_subfolder", False)
        assert config.set("export.to_md_image_extraction_mode", "file")
        assert config.set("link.format.image_link_style", "markdown_embed")
        mime = QMimeData()
        mime.setText(sample("rich-derived.txt").decode("utf-8"))
        mime.setHtml(sample("rich-derived.html").decode("utf-8"))
        if producer == "word":
            mime.setData(WORD_EMBED_SOURCE_MIME, sample("word-derived.ole"))
        else:
            mime.setData(WPS_DOCUMENT_MIME, sample("wps-writer-derived.zip"))
            mime.setData(WPS_IMAGE_DATA_MIME, sample("wps-images.bin"))
        qapp.clipboard().setMimeData(mime)
        window._on_paste_requested()
        qtbot.waitUntil(lambda: not window.view_model.inspection_busy)
        qtbot.waitUntil(lambda: window.view_model.selected_file is not None)
        selected = window.view_model.selected_file
        assert selected is not None and selected.format == "clipboard_document"
        model = load_clipboard_document_bytes(Path(selected.path).read_bytes())
        images = [inline for inline in iter_clipboard_inlines(model) if isinstance(inline, ClipboardImageRef)]
        assert [image.alt for image in images] == ["alpha-first", "nested-blue", "alpha-repeat"]
        expected_hashes = {resource.sha256 for resource in model.resources}
        assert len(expected_hashes) == 2
        window._action_area_vm.request_conversion("md")

        def finished():
            entry = window._batch_list_vm.get_file_entry(selected.path)
            return entry is not None and entry.status in {"completed", "failed"}

        qtbot.waitUntil(finished, timeout=30000)
        entry = window._batch_list_vm.get_file_entry(selected.path)
        assert entry is not None and entry.status == "completed", window._info_area_vm.history_rows
        assert entry.output_path
        output = Path(entry.output_path)
        text = output.read_text(encoding="utf-8-sig")
        assert text.index("BEFORE / 00123") < text.index("NESTED-BEFORE") < text.index("AFTER / repeated")
        assert all(text.count(f"![{alt}]") == 1 for alt in ["alpha-first", "nested-blue", "alpha-repeat"])
        pngs = list(output.parent.rglob("*.png"))
        assert len(pngs) == 2
        assert {hashlib.sha256(path.read_bytes()).hexdigest() for path in pngs} == expected_hashes
        assert Path(selected.path).is_file(), "History must retain the admitted bundle after completion"
        assert window._info_area_vm._task_summary.state == "success"
    finally:
        window.close()
        controller.stop()
        qapp.clipboard().clear()
        qapp.processEvents()
