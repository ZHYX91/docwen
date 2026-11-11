"""RTL filename display preserves input identity across preview and rejection."""

from __future__ import annotations

import pytest

from docwen_gui.i18n import get_locale, set_locale, t

pytestmark = pytest.mark.gui


@pytest.mark.parametrize(
    ("locale", "filename", "display_name"),
    [
        ("ar_SA", "01-docx.docx", "\u206801-docx.docx\u2069"),
        ("ar_SA", "01-تقرير.docx", "\u206801-تقرير.docx\u2069"),
        ("en_US", "01-docx.docx", "01-docx.docx"),
    ],
)
def test_current_filename_direction_keeps_original_identity_after_rejection(
    main_window, tmp_path, request, qtbot, locale, filename, display_name
) -> None:
    from docx import Document

    vm = main_window._input_area_vm
    main_vm = main_window._view_model
    original_locale = get_locale()
    request.addfinalizer(lambda: set_locale(original_locale))
    set_locale(locale)
    file_path = tmp_path / filename
    document = Document()
    document.add_paragraph("Synthetic document")
    document.save(file_path)
    original_bytes = file_path.read_bytes()

    preview = vm.build_drag_preview([str(file_path)])
    assert preview.message == t("components.file_drop.drag_preview_single", filename=display_name)
    assert preview.tooltip == str(file_path)

    vm.add_files([str(file_path)])
    qtbot.waitUntil(lambda: bool(main_vm.files) and not main_vm.inspection_busy)

    expected = t("components.file_drop.file_selected_msg", filename=display_name)
    assert vm.selection_message == expected
    selected = main_vm.selected_file
    assert selected is not None
    assert selected.path == str(file_path)
    metadata = dict(selected.metadata)
    assert vm.selected_file_path == str(file_path)
    assert vm.selection_detail == str(tmp_path)

    vm.add_files([str(tmp_path / "missing.docx")])

    assert vm.selection_message.split("\n", 1)[0] == expected
    assert main_vm.selected_file is selected
    assert selected.path == str(file_path)
    assert selected.metadata == metadata
    assert file_path.name == filename
    assert file_path.read_bytes() == original_bytes
