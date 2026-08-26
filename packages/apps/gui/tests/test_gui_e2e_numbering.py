"""Phase B-GUI-E2E: GUI end-to-end numbering real-effect assertions.

These tests verify that switching numbering options in the
ActionArea widget (checkboxes, combos) produces a visible effect
on the conversion output — the plan's '真实序号生效路径'.

The pattern: dock a Markdown file, toggle numbering controls on
the ActionArea, trigger conversion to DOCX, verify the output
DOCX contains the expected structural evidence.

**Marks**: ``pytest.mark.gui`` (excluded from default unit suite).
These tests require ``main_window_with_controller`` (real runtime,
real window, offscreen QPA). They are slower than unit tests and
should be run explicitly:

    pytest packages/apps/gui/tests -m gui -k "numbering_e2e"
"""

from __future__ import annotations

from pathlib import Path

import pytest

# These tests inherit the e2e markers and conftest from
# test_gui_e2e_conversion.py — same fixture, same pattern.

pytestmark = pytest.mark.gui

_E2E_CONVERSION_TIMEOUT_MS = 60000


def _wait_for(
    condition,
    timeout_ms: int = _E2E_CONVERSION_TIMEOUT_MS,
    interval_ms: int = 100,
) -> bool:
    from PySide6.QtCore import QEventLoop, QTimer

    done: list[bool] = [False]
    loop = QEventLoop()

    def _check() -> None:
        if condition():
            done[0] = True
            loop.quit()

    periodic_timer = QTimer()
    periodic_timer.setInterval(interval_ms)
    periodic_timer.timeout.connect(_check)
    periodic_timer.start()

    safety_timer = QTimer()
    safety_timer.setSingleShot(True)
    safety_timer.timeout.connect(loop.quit)
    safety_timer.start(timeout_ms)

    try:
        loop.exec()
    finally:
        periodic_timer.stop()
        periodic_timer.deleteLater()
        safety_timer.stop()
        safety_timer.deleteLater()

    return done[0]


class TestMdToDocxNumberingE2E:
    """E2E: MD->DOCX with numbering controls toggled in the ActionArea.

    Uses a real main window + runtime (enqueued by
    ``main_window_with_controller``). The test docks a Markdown file,
    toggles numbering options, triggers conversion, and asserts the
    output DOCX contains the expected numbering evidence.
    """

    def test_remove_numbering_strips_heading_prefix(self, main_window_with_controller, tmp_path) -> None:
        """``remove_numbering=True`` strips existing numbering from output DOCX."""
        from PySide6.QtWidgets import QApplication

        window = main_window_with_controller
        vm = window.view_model

        # Write a Markdown file with handwritten numbering.
        md_path = tmp_path / "e2e_remove_numbering.md"
        md_path.write_text(
            "# 一、引言\n\n正文一。\n\n## （一）背景\n\n正文二。\n",
            encoding="utf-8",
        )

        # Dock the Markdown file — this triggers setup_for_md_to_document.
        vm.add_files([str(md_path)])
        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        assert vm.has_files
        assert window._action_area_vm.visible

        # Toggle: remove existing numbering, do NOT add new.
        aa = window._action_area_vm
        aa.set_md_to_doc_option("remove_numbering", True)
        aa.set_md_to_doc_option("add_numbering", False)

        if app is not None:
            app.processEvents()

        # Trigger conversion to DOCX.
        from docwen_gui.main_window import _normalize_path

        normalized = _normalize_path(str(md_path))

        window._action_area_vm.request_conversion("docx")
        if app is not None:
            app.processEvents()

        def _done() -> bool:
            entry = window._batch_list_vm.get_file_entry(normalized)
            if entry is None:
                return False
            return entry.status in ("completed", "failed")

        assert _wait_for(_done), "Conversion did not complete within timeout"

        if app is not None:
            app.processEvents()

        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == "completed", f"Conversion failed: {entry.status}"
        assert entry.output_path, "No output path from conversion"

        output_path = Path(entry.output_path)
        assert output_path.exists()
        assert output_path.stat().st_size > 0

        # Verify the output DOCX — heading text should be bare
        # (handwritten prefix stripped, no add).
        import zipfile

        from lxml import etree

        WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        with zipfile.ZipFile(str(output_path)) as zf:
            doc = etree.fromstring(zf.read("word/document.xml"))
            styles = etree.fromstring(zf.read("word/styles.xml"))
        body = doc.find(f"{{{WML}}}body")
        assert body is not None, "No body element in document.xml"

        heading_style_ids: set[str] = set()
        for style_el in styles.findall(f".//{{{WML}}}style"):
            style_id = style_el.get(f"{{{WML}}}styleId", "")
            name_el = style_el.find(f"{{{WML}}}name")
            name = (name_el.get(f"{{{WML}}}val", "") if name_el is not None else "").lower()
            if style_id.startswith("Heading") or name.startswith("heading "):
                heading_style_ids.add(style_id)

        headings: list[str] = []
        for p in body.findall(f".//{{{WML}}}p"):
            ppr = p.find(f"{{{WML}}}pPr")
            style = ""
            if ppr is not None:
                ps = ppr.find(f"{{{WML}}}pStyle")
                if ps is not None:
                    style = ps.get(f"{{{WML}}}val", "")
            if style in heading_style_ids:
                text = "".join((t.text or "") for t in p.findall(f".//{{{WML}}}t"))
                headings.append(text)

        assert headings, f"No headings found in output DOCX: {output_path}"

        # With remove=True and add=False, headings should be bare
        # (no 一、/（一） prefixes).
        for h in headings:
            assert not h.startswith("一、"), f"heading '{h}' should have had '一、' removed"
            assert not h.startswith("（一）"), f"heading '{h}' should have had '（一）' removed"
