from __future__ import annotations

import os
import threading
import time
from pathlib import Path

import pytest
from tests.support.gui import shutdown_main_window

from docwen_gui.i18n import t
from docwen_gui.view_models.batch_list_vm import BatchFileEntry

_E2E_CONVERSION_TIMEOUT_MS = 60000


def _create_ocr_smoke_png(output_path: Path, text: str = "HELLO DOCWEN OCR") -> Path:
    from PIL import Image, ImageDraw, ImageFont

    image = Image.new("RGB", (720, 220), "white")
    draw = ImageDraw.Draw(image)
    font_path = Path("C:/Windows/Fonts/arialbd.ttf")
    font = ImageFont.truetype(str(font_path), 56) if font_path.exists() else ImageFont.load_default()
    draw.text((34, 76), text, fill="black", font=font)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(output_path)
    return output_path


def _has_registered_com_prog_id(*prog_ids: str) -> bool:
    import sys

    if sys.platform != "win32":
        return False
    try:
        import winreg
    except ImportError:
        return False

    for prog_id in prog_ids:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{prog_id}\\CLSID"):
                return True
        except OSError:
            continue
    return False


def _has_document_pdf_backend() -> bool:
    from docwen_core.office_bridge import find_soffice_path

    return _has_registered_com_prog_id("Word.Application", "KWPS.Application", "Kwps.Application") or bool(
        find_soffice_path()
    )


def _has_spreadsheet_pdf_backend() -> bool:
    from docwen_core.office_bridge import find_soffice_path

    return _has_registered_com_prog_id("Excel.Application", "KET.Application") or bool(find_soffice_path())


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


def _read_pdf_text(path: Path) -> str:
    import fitz

    with fitz.open(path) as document:
        text_parts: list[str] = []
        for page in document:
            page_text = page.get_text("text")
            assert isinstance(page_text, str)
            text_parts.append(" ".join(page_text.split()))
        return " ".join(text_parts)


def _assert_markdown_node(
    output_path: Path,
    *,
    source_stem: str,
    source_tag: str,
    output_root: Path | None = None,
) -> None:
    assert output_path.exists()
    assert output_path.parent.name == output_path.stem
    assert output_path.stem.startswith(f"{source_stem}_")
    assert output_path.stem.endswith(f"_from{source_tag}")
    assert (output_path.parent / "docwen-node.json").is_file()
    if output_root is not None:
        assert output_path.parent.parent == output_root


__all__ = (
    "_E2E_CONVERSION_TIMEOUT_MS",
    "BatchFileEntry",
    "Path",
    "_assert_markdown_node",
    "_create_ocr_smoke_png",
    "_has_document_pdf_backend",
    "_has_spreadsheet_pdf_backend",
    "_read_pdf_text",
    "_wait_for",
    "os",
    "pytest",
    "shutdown_main_window",
    "t",
    "threading",
    "time",
)
