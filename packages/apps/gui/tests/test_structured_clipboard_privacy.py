"""Real clipboard ingress must not copy provider metadata or raw exceptions to logs."""

from __future__ import annotations

import logging
from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData

from docwen_gui.main_window import MainWindow
from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = [pytest.mark.gui, pytest.mark.pr_gate, pytest.mark.release_gate]


def _cf_html(fragment: str, source_url: str) -> bytes:
    prefix = (
        "Version:1.0\r\n"
        "StartHTML:{start_html:010d}\r\n"
        "EndHTML:{end_html:010d}\r\n"
        "StartFragment:{start_fragment:010d}\r\n"
        "EndFragment:{end_fragment:010d}\r\n"
        f"SourceURL:{source_url}\r\n"
    )
    html = f"<html><body><!--StartFragment-->{fragment}<!--EndFragment--></body></html>".encode()
    start_html = len(prefix.format(start_html=0, end_html=0, start_fragment=0, end_fragment=0).encode())
    header = prefix.format(
        start_html=start_html,
        end_html=start_html + len(html),
        start_fragment=start_html + html.index(b"<!--StartFragment-->") + len(b"<!--StartFragment-->"),
        end_fragment=start_html + html.index(b"<!--EndFragment-->"),
    ).encode()
    return header + html


@pytest.mark.parametrize("case", ["structured", "invalid-offsets", "projection-exception"])
def test_main_window_provider_metadata_and_raw_exception_privacy(qapp, qtbot, tmp_path, monkeypatch, caplog, case):
    source_url = "file:///C:/CF_HTML_PRIVATE_52a9/source.docx"
    alt = "ALT_PRIVATE_b71e"
    image_url = "https://example.invalid/IMG_URL_PRIVATE_d836/private.png"
    raw_exception = "RAW_EXCEPTION_PRIVATE_f420"
    fragment = f'<p>before<img src="{image_url}" alt="{alt}"></p>'
    if case != "projection-exception":
        fragment += "<table><tr><td>00123</td></tr></table><p>after</p>"
    payload = _cf_html(fragment, source_url)
    if case == "invalid-offsets":
        payload = payload.replace(b"EndFragment:", b"BadFragment:", 1)
    if case == "projection-exception":
        from docwen_gui import clipboard_content

        def fail_projection(_html):
            raise ValueError(raw_exception)

        monkeypatch.setattr(clipboard_content, "project_clipboard_html", fail_projection)
    mime = QMimeData()
    mime.setData("text/html", payload)
    mime.setText("original plain body")
    window = MainWindow(
        view_model=MainWindowViewModel(controller=None),
        clipboard_input_root=tmp_path / "clipboard-inputs",
    )
    qtbot.addWidget(window)
    caplog.set_level(logging.DEBUG)
    try:
        qapp.clipboard().setMimeData(mime)
        window._on_paste_requested()
        if case == "invalid-offsets":
            assert window.view_model.files == []
        else:
            qtbot.waitUntil(lambda: not window.view_model.inspection_busy and len(window.view_model.files) == 1)
            saved = Path(window.view_model.files[0].path).read_text(encoding="utf-8")
            if case == "structured":
                assert alt in saved and "00123" in saved
                assert source_url not in saved and image_url not in saved
            else:
                assert saved == "original plain body"
                assert "Clipboard HTML projection failed" in caplog.text
        for private_value in (source_url, alt, image_url, raw_exception):
            assert private_value not in caplog.text
    finally:
        window.close()
        qapp.clipboard().clear()
        qapp.processEvents()
