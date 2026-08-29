"""
GUI 截图脚本（新项目版本）

用于导出主窗口、设置对话框和关键场景的截图，支持多种主题批量输出。基于当前新项目代码库（阶段 1/2 完成后）重写。

使用方式：
    python scripts/gui_capture.py
    python scripts/gui_capture.py --targets main settings about batch status progress completed failed cancelled template conversion-document conversion-spreadsheet conversion-image conversion-layout --themes light dark
    python scripts/gui_capture.py --targets progress completed failed cancelled --output-dir tmp/gui-shots-runtime
    python scripts/gui_capture.py --targets settings --output-dir tmp/gui-shots
    python scripts/gui_capture.py --lang en_US --targets main settings --output-dir tmp/gui-shots-en
"""

from __future__ import annotations

import argparse
import os
import sys
from contextlib import suppress
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
LOCAL_SRC_PATHS = [
    PROJECT_ROOT,
    PROJECT_ROOT / "packages" / "core" / "src",
    PROJECT_ROOT / "packages" / "application" / "src",
    PROJECT_ROOT / "packages" / "runtime" / "src",
    PROJECT_ROOT / "packages" / "bundle" / "src",
    PROJECT_ROOT / "packages" / "apps" / "gui" / "src",
]
for path in reversed(LOCAL_SRC_PATHS):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

os.environ.setdefault("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")

from PySide6.QtCore import QSize  # noqa: E402
from PySide6.QtWidgets import QApplication, QLabel, QScrollArea, QTabWidget, QVBoxLayout, QWidget  # noqa: E402

from docwen_gui.styles.theme_manager import ThemeManager  # noqa: E402

DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "tmp" / "gui-shots"
THEME_PRESETS = {
    "light": "light",
    "dark": "dark",
    "system": "system",
}


def _available_locales() -> list[str]:
    locales_dir = PROJECT_ROOT / "i18n" / "locales"
    return sorted(path.stem for path in locales_dir.glob("*.toml"))


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export DocWen GUI screenshots")
    parser.add_argument(
        "--targets",
        nargs="+",
        choices=[
            "main",
            "settings",
            "about",
            "batch",
            "status",
            "progress",
            "completed",
            "failed",
            "cancelled",
            "template",
            "conversion",
            "conversion-document",
            "conversion-spreadsheet",
            "conversion-image",
            "conversion-layout",
        ],
        default=[
            "main",
            "settings",
            "about",
            "batch",
            "status",
            "progress",
            "completed",
            "failed",
            "cancelled",
            "template",
            "conversion-document",
            "conversion-spreadsheet",
            "conversion-image",
            "conversion-layout",
        ],
        help="Screenshot targets",
    )
    parser.add_argument(
        "--themes",
        nargs="+",
        choices=sorted(THEME_PRESETS),
        default=["light", "dark"],
        help="Themes to capture (light, dark, system)",
    )
    parser.add_argument(
        "--output-dir",
        default=str(DEFAULT_OUTPUT_DIR),
        help="Output directory for screenshots",
    )
    parser.add_argument("--width", type=int, default=1200, help="Main window width")
    parser.add_argument("--height", type=int, default=760, help="Main window height")
    parser.add_argument(
        "--lang",
        choices=_available_locales(),
        default=None,
        help="GUI locale to use for captured widgets",
    )
    parser.add_argument("--list-themes", action="store_true", help="List available themes and exit")
    return parser.parse_args(argv)


def _ensure_app() -> QApplication:
    app = QApplication.instance()
    if isinstance(app, QApplication):
        return app
    from docwen_gui.app import create_qapplication

    return create_qapplication(["docwen_capture"])


def _normalize_theme(theme_name: str) -> str:
    normalized = theme_name.strip().lower()
    return THEME_PRESETS.get(normalized, normalized or "light")


def _prepare_widget(widget: QWidget, size: QSize) -> None:
    widget.resize(size)
    widget.show()
    widget.raise_()
    widget.ensurePolished()
    layout = widget.layout()
    if layout is not None:
        layout.activate()
    app = QApplication.instance()
    if isinstance(app, QApplication):
        # Fluent widgets schedule parts of their geometry update.  Settle those
        # queued passes before grabbing so captures represent the final layout.
        for _ in range(3):
            app.sendPostedEvents()
            app.processEvents()


def _capture_widget(widget: QWidget, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    pixmap = widget.grab()
    if not pixmap.save(str(output_path)):
        raise RuntimeError(f"Failed to save screenshot: {output_path}")


# ── Main window capture ────────────────────────────────────────────────────


def _capture_main_window(theme_name: str, output_dir: Path, size: QSize) -> Path:
    window, _vm = _create_capture_main_window()
    try:
        _prepare_widget(window, size)
        output_path = output_dir / f"main-{theme_name}.png"
        _capture_widget(window, output_path)
        return output_path
    finally:
        window.close()


# ── Settings dialog capture ─────────────────────────────────────────────────


def _capture_settings_dialog(theme_name: str, output_dir: Path, size: QSize) -> list[Path]:
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, SettingsDialog

    vm = _create_settings_view_model(SettingsViewModel)
    dialog = SettingsDialog(parent=None, view_model=vm)
    try:
        _prepare_widget(dialog, size)
        app = QApplication.instance()
        output_paths: list[Path] = []
        output_path = output_dir / f"settings-{theme_name}.png"
        _capture_widget(dialog, output_path)
        output_paths.append(output_path)

        tab_widget = dialog.findChild(QTabWidget)
        if tab_widget is not None:
            for index in range(tab_widget.count()):
                tab_widget.setCurrentIndex(index)
                if isinstance(app, QApplication):
                    app.processEvents()
                tab_key = (
                    TAB_KEYS[index]
                    if index < len(TAB_KEYS)
                    else _safe_slug(tab_widget.tabText(index) or f"tab-{index + 1}")
                )
                tab_output_path = output_dir / f"settings-tab-{tab_key}-{theme_name}.png"
                _capture_widget(dialog, tab_output_path)
                output_paths.append(tab_output_path)
        return output_paths
    finally:
        dialog.close()


def _create_settings_view_model(settings_vm_cls):
    try:
        from docwen_application.controller import ApplicationController
        from docwen_bundle.config_port import ConfigPortAdapter

        config_port = ConfigPortAdapter(
            base_dir=PROJECT_ROOT / "configs",
            user_dir=PROJECT_ROOT / "tmp" / "gui-capture-config",
        )
        controller = ApplicationController(config_port=config_port)
        return settings_vm_cls(controller=controller)
    except Exception:
        return settings_vm_cls(controller=None)


# ── About dialog capture ────────────────────────────────────────────────────


def _capture_about_dialog(theme_name: str, output_dir: Path) -> list[Path]:
    from docwen_gui.dialogs.about import AboutDialog

    dialog = AboutDialog(parent=None)
    try:
        _prepare_widget(dialog, dialog.size())
        output_paths: list[Path] = []
        output_path = output_dir / f"about-{theme_name}.png"
        _capture_widget(dialog, output_path)
        output_paths.append(output_path)

        scroll_area = dialog.findChild(QScrollArea, "aboutScrollArea")
        if scroll_area is not None:
            scroll_bar = scroll_area.verticalScrollBar()
            if scroll_bar.maximum() > 0:
                scroll_bar.setValue(scroll_bar.maximum())
                app = QApplication.instance()
                if isinstance(app, QApplication):
                    app.processEvents()
                bottom_output_path = output_dir / f"about-bottom-{theme_name}.png"
                _capture_widget(dialog, bottom_output_path)
                output_paths.append(bottom_output_path)
        return output_paths
    finally:
        dialog.close()


# ── Batch list preview ──────────────────────────────────────────────────────


def _capture_batch_preview(theme_name: str, output_dir: Path, size: QSize) -> Path:
    """Create a standalone batch list widget with sample entries for screenshot."""
    from docwen_gui.view_models.batch_list_vm import BatchListViewModel
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel
    from docwen_gui.widgets.batch_list import BatchList

    main_vm = MainWindowViewModel(controller=None)
    vm = BatchListViewModel(main_vm=main_vm)
    sample_paths = [
        r"S:\Projects\Archive\reports\2026\very_long_filename_for_status_balance_review_document_v3.docx",
        r"S:\Projects\Archive\reports\2026\completed-contract-review-summary.docx",
        r"S:\Projects\Archive\reports\2026\skipped-legacy-template-copy.docx",
        r"S:\Projects\Data\incoming\invoice-summary-2026.xlsx",
        r"S:\Projects\Media\layout\batch\failed-scan-with-mixed-footnotes-and-stamps.pdf",
        r"S:\Projects\Books\imports\weekly-digest.epub",
    ]
    added, _failed = vm.add_files(sample_paths)
    for path in added:
        entry = vm.get_file_entry(path)
        if entry is None:
            continue
        if entry.file_name == "very_long_filename_for_status_balance_review_document_v3.docx":
            entry.size_bytes = 245_760
        elif entry.file_name == "completed-contract-review-summary.docx":
            entry.size_bytes = 1_048_576
            entry.warning_message = _capture_sample_text("batch_template_warning")
            vm.set_file_status(
                path,
                "completed",
                output_path=r"S:\Projects\Output\completed-contract-review-summary.md",
            )
        elif entry.file_name == "skipped-legacy-template-copy.docx":
            entry.size_bytes = 776_192
            vm.set_file_status(path, "skipped", skip_reason=_capture_sample_text("batch_skip_existing"))
        elif entry.file_name == "invoice-summary-2026.xlsx":
            entry.size_bytes = 1_126_400
            vm.set_file_status(path, "completed", output_path=r"S:\Projects\Output\invoice-summary-2026.md")
        elif entry.file_name == "failed-scan-with-mixed-footnotes-and-stamps.pdf":
            entry.size_bytes = 3_145_728
            vm.set_file_status(
                path,
                "failed",
                error_message=_capture_sample_text("batch_broken_layout"),
                error_count=2,
            )
        elif entry.file_name == "weekly-digest.epub":
            entry.size_bytes = 786_432
            vm.set_file_status(path, "skipped", skip_reason=_capture_sample_text("batch_plugin_disabled"))
    vm.activate_tab("document")
    widget = BatchList(view_model=vm)
    widget.select_file(r"S:/Projects/Archive/reports/2026/completed-contract-review-summary.docx")
    try:
        _prepare_widget(widget, size)
        output_path = output_dir / f"batch-{theme_name}.png"
        _capture_widget(widget, output_path)
        return output_path
    finally:
        widget.close()


# ── Status / info-area preview ──────────────────────────────────────────────


def _capture_status_preview(theme_name: str, output_dir: Path, size: QSize) -> Path:
    """Create a standalone info area with representative history/task states."""
    from docwen_gui.view_models.info_area_vm import InfoAreaViewModel
    from docwen_gui.widgets.info_area import InfoArea

    widget = QWidget()
    widget.setWindowTitle("DocWen status preview")
    layout = QVBoxLayout(widget)
    layout.setContentsMargins(12, 12, 12, 12)
    layout.setSpacing(8)

    title = QLabel(_capture_sample_text("status_title"), widget)
    title_font = title.font()
    title_font.setBold(True)
    title.setFont(title_font)
    layout.addWidget(title)

    description = QLabel(_capture_sample_text("status_description"), widget)
    description.setWordWrap(True)
    layout.addWidget(description)

    vm = InfoAreaViewModel(parent=widget)
    info_area = InfoArea(view_model=vm, parent=widget)
    vm.set_transient_message(
        "progress:preview",
        _capture_sample_text("status_progress"),
        "info",
        ttl_ms=0,
        source="preview",
    )
    vm.add_message(_capture_sample_text("status_copied"), "info")
    vm.add_message(
        _capture_sample_text("status_completed"),
        "success",
        show_location=True,
        file_path=r"S:\Projects\Reports\2026Q2-预算执行月报.docx",
        navigate_file_path=r"S:\Projects\Reports\2026Q2-预算执行月报.docx",
        operation_id="preview-success",
    )
    vm.add_message(
        _capture_sample_text("status_format_warning"),
        "warning",
    )
    vm.add_message(
        _capture_sample_text("status_merge_failed"),
        "danger",
        show_location=True,
        file_path=r"S:\Projects\Layouts\merged\issue-17.pdf",
        navigate_file_path=r"S:\Projects\Layouts\merged\issue-17.pdf",
        operation_id="preview-failed",
    )
    vm.set_task_summary(
        operation_id="preview-batch",
        current_file="annual-report-layout.pdf",
        completed_count=42,
        total_count=48,
        failed_count=3,
        state="partial",
        tone="warning",
        navigate_file_path=r"S:\Projects\Output\failed-details.json",
        guide_actions=InfoAreaViewModel.compute_guide_actions(
            "partial",
            output_dir=r"S:\Projects\Output",
            failed_details_path=r"S:\Projects\Output\failed-details.json",
            retry_available=True,
        ),
    )
    layout.addWidget(info_area, stretch=1)

    try:
        _prepare_widget(widget, size)
        output_path = output_dir / f"status-{theme_name}.png"
        _capture_widget(widget, output_path)
        return output_path
    finally:
        vm.stop_all_timers()
        widget.close()


# ── Runtime status scenes ───────────────────────────────────────────────────


def _capture_runtime_status_scene(
    theme_name: str,
    output_dir: Path,
    size: QSize,
    *,
    state: str,
) -> Path:
    """Capture the main window in representative runtime terminal states."""
    from docwen_gui.view_models.info_area_vm import InfoAreaViewModel

    window, vm = _create_capture_main_window()

    active_path = r"S:\Projects\Archive\reports\2026\annual-summary-2026.docx"
    queued_path = r"S:\Projects\Archive\reports\2026\board-review-appendix.docx"
    fake_ref = _FakeFileRef(active_path, "document", "docx")

    _sync_capture_mode(window, vm, "batch")
    _sync_capture_selected_file(window, vm, fake_ref)
    _sync_capture_category(window, "document")

    added, _failed = window._batch_list_vm.add_files([active_path, queued_path])
    for path in added:
        entry = window._batch_list_vm.get_file_entry(path)
        if entry is not None:
            entry.size_bytes = 1_245_184 if path == active_path else 842_752
    window._batch_list.select_file(active_path)
    window._action_area_vm.setup_for_document_file(active_path)

    if state == "progress":
        window._batch_list_vm.set_file_status(
            active_path,
            "processing",
            operation_id="preview-runtime",
        )
        window._action_area_vm.show_cancel()
        window._info_area_vm.set_transient_message(
            "progress:preview-runtime",
            _capture_sample_text("runtime_progress"),
            "info",
            ttl_ms=0,
            source="preview-runtime",
        )
        window._info_area_vm.set_task_summary(
            operation_id="preview-runtime",
            current_file=Path(active_path).name,
            current_file_path=active_path,
            completed_count=0,
            total_count=2,
            state="active",
            tone="info",
            navigation_kind="current",
        )
        window._info_area_vm.start_activity_animation(_capture_sample_text("runtime_processing"))
    elif state == "completed":
        completed_output_dir = r"S:\Projects\Output\2026-06-28"
        active_output = completed_output_dir + r"\annual-summary-2026.md"
        queued_output = completed_output_dir + r"\board-review-appendix.md"
        window._batch_list_vm.set_file_status(
            active_path,
            "completed",
            output_path=active_output,
            operation_id="preview-runtime",
        )
        window._batch_list_vm.set_file_status(
            queued_path,
            "completed",
            output_path=queued_output,
            operation_id="preview-runtime-2",
        )
        window._action_area_vm.hide_cancel()
        window._info_area_vm.add_message(
            _capture_sample_text("runtime_completed_active"),
            "success",
            show_location=True,
            file_path=active_output,
            navigate_file_path=active_output,
            operation_id="preview-runtime",
        )
        window._info_area_vm.add_message(
            _capture_sample_text("runtime_completed_batch"),
            "success",
            show_location=True,
            file_path=completed_output_dir,
            navigate_file_path=completed_output_dir,
            operation_id="preview-runtime-batch",
        )
        window._info_area_vm.set_task_summary(
            operation_id="preview-runtime",
            current_file=Path(active_path).name,
            current_file_path=active_path,
            completed_count=2,
            total_count=2,
            failed_count=0,
            state="success",
            tone="success",
            navigate_file_path=completed_output_dir,
            navigation_kind="output",
            guide_actions=InfoAreaViewModel.compute_guide_actions("success", output_dir=completed_output_dir),
        )
    elif state == "failed":
        failed_details_path = r"S:\Projects\Output\2026-06-28\failed-items.json"
        failed_output_dir = r"S:\Projects\Output\2026-06-28"
        active_error = _capture_sample_text("runtime_failed_active")
        queued_error = _capture_sample_text("runtime_failed_queued")
        window._batch_list_vm.set_file_status(
            active_path,
            "failed",
            error_message=active_error,
            operation_id="preview-runtime",
        )
        window._batch_list_vm.set_file_status(
            queued_path,
            "failed",
            error_message=queued_error,
            operation_id="preview-runtime-2",
        )
        window._action_area_vm.hide_cancel()
        window._info_area_vm.add_message(
            active_error,
            "danger",
            show_location=True,
            file_path=active_path,
            navigate_file_path=active_path,
            operation_id="preview-runtime",
        )
        window._info_area_vm.add_message(
            _capture_sample_text("runtime_failed_summary"),
            "danger",
            show_location=True,
            file_path=failed_details_path,
            navigate_file_path=failed_details_path,
            operation_id="preview-runtime-failed",
        )
        window._info_area_vm.set_task_summary(
            operation_id="preview-runtime",
            current_file=Path(active_path).name,
            current_file_path=active_path,
            completed_count=0,
            total_count=2,
            failed_count=2,
            state="failed",
            tone="danger",
            navigate_file_path=failed_details_path,
            navigation_kind="failed",
            guide_actions=InfoAreaViewModel.compute_guide_actions(
                "failed",
                output_dir=failed_output_dir,
                failed_details_path=failed_details_path,
                retry_available=True,
            ),
        )
    elif state == "cancelled":
        window._batch_list_vm.set_file_status(
            active_path,
            "cancelled",
            error_message=_capture_sample_text("runtime_cancelled"),
            operation_id="preview-runtime",
        )
        window._info_area_vm.add_message(
            _capture_sample_text("runtime_cancelled"),
            "warning",
            show_location=True,
            file_path=active_path,
            navigate_file_path=active_path,
            operation_id="preview-runtime",
        )
        window._info_area_vm.set_task_summary(
            operation_id="preview-runtime",
            current_file=Path(active_path).name,
            current_file_path=active_path,
            completed_count=0,
            total_count=2,
            failed_count=0,
            cancelled_count=1,
            state="cancelled",
            tone="warning",
            navigate_file_path=active_path,
            navigation_kind="failed",
            guide_actions=InfoAreaViewModel.compute_guide_actions("cancelled"),
        )
    else:
        raise ValueError(f"Unsupported runtime status scene: {state!r}")

    try:
        _prepare_widget(window, size)
        output_path = output_dir / f"{state}-{theme_name}.png"
        _capture_widget(window, output_path)
        return output_path
    finally:
        window._info_area_vm.stop_all_timers()
        window.close()


# ── Template selector capture ───────────────────────────────────────────────


def _capture_template_scene(theme_name: str, output_dir: Path, size: QSize) -> Path:
    """Capture the main window with a markdown file to show template selector."""
    window, vm = _create_capture_main_window()

    # Simulate a markdown file selection to trigger template projection
    vm.set_mode("single")
    fake_ref = _FakeFileRef("/tmp/sample.md", "markdown", "md")
    _sync_capture_selected_file(window, vm, fake_ref)
    _sync_capture_category(window, "markdown")
    # Populate template selector with sample templates for visual review
    ts = window._template_selector
    if ts is not None:
        ts.load_templates("docx", ["Corporate-Report-Blue", "Academic-Thesis-CN", "Meeting-Minutes-Minimal"])
        ts.load_templates("xlsx", ["Quarterly-Dashboard", "Inventory-Checklist"])
        ts.activate_and_select("docx")
    try:
        _prepare_widget(window, size)
        output_path = output_dir / f"template-{theme_name}.png"
        _capture_widget(window, output_path)
        return output_path
    finally:
        window.close()


# ── Conversion panel capture ────────────────────────────────────────────────

_CONVERSION_SCENARIOS = {
    "document": (
        r"S:\Projects\Documents\annual-summary-2026.docx",
        "document",
        "docx",
        "conversion-document",
    ),
    "spreadsheet": (
        r"S:\Projects\Data\incoming\invoice-summary-2026.xlsx",
        "spreadsheet",
        "xlsx",
        "conversion-spreadsheet",
    ),
    "image": (
        r"S:\Projects\Images\scanned-contract-page-01.png",
        "image",
        "png",
        "conversion-image",
    ),
    "layout": (
        r"S:\Projects\Media\layout\annual-report-2026.pdf",
        "layout",
        "pdf",
        "conversion-layout",
    ),
}


def _capture_conversion_scene(
    theme_name: str,
    output_dir: Path,
    size: QSize,
    *,
    scenario: str = "document",
    legacy_name: bool = False,
) -> Path:
    """Capture the main window with a file selected to show conversion panel."""
    window, vm = _create_capture_main_window()

    file_path, category, fmt, file_stem = _CONVERSION_SCENARIOS[scenario]
    fake_ref = _FakeFileRef(file_path, category, fmt)
    _sync_capture_selected_file(window, vm, fake_ref)
    _sync_capture_category(window, category)
    if scenario == "layout":
        window._conversion_panel_vm.set_pdf_info(18, Path(file_path).name)
    try:
        _prepare_widget(window, size)
        output_name = f"conversion-{theme_name}.png" if legacy_name else f"{file_stem}-{theme_name}.png"
        output_path = output_dir / output_name
        _capture_widget(window, output_path)
        return output_path
    finally:
        window.close()


# ── Helpers ─────────────────────────────────────────────────────────────────


def _create_capture_main_window() -> tuple[QWidget, object]:
    """Create the main window through the same composition root as the GUI."""
    from docwen_gui.app import create_main_window

    window = create_main_window(controller=None)
    return window, window._view_model


class _FakeFileRef:
    """Minimal FileRef stand-in for screenshot capture."""

    def __init__(self, path: str, category: str, fmt: str) -> None:
        self.path = path
        self.category = category
        self.format = fmt
        self.warning_message = ""


def _safe_slug(text: str) -> str:
    slug = "".join(ch.lower() if ch.isalnum() else "-" for ch in text.strip())
    slug = "-".join(part for part in slug.split("-") if part)
    return slug or "tab"


_CAPTURE_SAMPLE_TEXT: dict[str, dict[str, str]] = {
    "batch_template_warning": {
        "zh_CN": "检测到模板字段缺失，已按默认字段继续导出。",
        "zh_TW": "偵測到範本欄位缺失，已依預設欄位繼續匯出。",
        "en_US": "Template fields are missing; export continued with default fields.",
    },
    "batch_skip_existing": {
        "zh_CN": "输出文件已存在，按设置跳过。",
        "zh_TW": "輸出檔案已存在，已依設定略過。",
        "en_US": "Output file already exists; skipped according to settings.",
    },
    "batch_broken_layout": {
        "zh_CN": "第 17 页包含损坏对象引用，无法继续解析版式。",
        "zh_TW": "第 17 頁包含損壞的物件參照，無法繼續解析版面。",
        "en_US": "Page 17 contains a broken object reference; layout parsing cannot continue.",
    },
    "batch_plugin_disabled": {
        "zh_CN": "未启用出版物转换插件。",
        "zh_TW": "未啟用出版物轉換外掛。",
        "en_US": "Publication conversion plugin is not enabled.",
    },
    "status_title": {
        "zh_CN": "状态栏预览",
        "zh_TW": "狀態列預覽",
        "en_US": "Status Preview",
    },
    "status_description": {
        "zh_CN": "覆盖即时消息、长历史消息、定位按钮、任务摘要和完成引导。",
        "zh_TW": "涵蓋即時訊息、長歷史訊息、定位按鈕、任務摘要與完成引導。",
        "en_US": "Covers live messages, long history items, location buttons, task summaries, and completion guidance.",
    },
    "status_progress": {
        "zh_CN": "正在转换第 12 / 48 个文件，请稍候。",
        "zh_TW": "正在轉換第 12 / 48 個檔案，請稍候。",
        "en_US": "Converting file 12 of 48. Please wait.",
    },
    "status_copied": {
        "zh_CN": "已复制输出目录到剪贴板。",
        "zh_TW": "已將輸出目錄複製到剪貼簿。",
        "en_US": "Output directory copied to clipboard.",
    },
    "status_completed": {
        "zh_CN": "已完成：2026Q2-预算执行月报.docx",
        "zh_TW": "已完成：2026Q2-預算執行月報.docx",
        "en_US": "Completed: 2026Q2-budget-execution-report.docx",
    },
    "status_format_warning": {
        "zh_CN": "检测到 3 个文件的扩展名与实际内容不一致，已按实际格式继续处理并写入结果摘要。",
        "zh_TW": "偵測到 3 個檔案的副檔名與實際內容不一致，已依實際格式繼續處理並寫入結果摘要。",
        "en_US": "Three files have extensions that differ from their actual content; processing continued with detected formats and was recorded in the summary.",
    },
    "status_merge_failed": {
        "zh_CN": "合并版式文件失败：第 17 页包含损坏对象引用，已跳过该页并等待用户确认是否继续导出。",
        "zh_TW": "合併版面檔案失敗：第 17 頁包含損壞的物件參照，已略過該頁並等待使用者確認是否繼續匯出。",
        "en_US": "Layout merge failed: page 17 contains a broken object reference. The page was skipped while waiting for confirmation to continue exporting.",
    },
    "runtime_progress": {
        "zh_CN": "进度：35% 正在转换 annual-summary-2026.docx",
        "zh_TW": "進度：35% 正在轉換 annual-summary-2026.docx",
        "en_US": "Progress: 35% Converting annual-summary-2026.docx",
    },
    "runtime_processing": {
        "zh_CN": "正在处理",
        "zh_TW": "正在處理",
        "en_US": "Processing",
    },
    "runtime_completed_active": {
        "zh_CN": "已完成：annual-summary-2026.docx → annual-summary-2026.md",
        "zh_TW": "已完成：annual-summary-2026.docx → annual-summary-2026.md",
        "en_US": "Completed: annual-summary-2026.docx -> annual-summary-2026.md",
    },
    "runtime_completed_batch": {
        "zh_CN": "批量转换完成：2 / 2 个文件已写入输出目录。",
        "zh_TW": "批次轉換完成：2 / 2 個檔案已寫入輸出目錄。",
        "en_US": "Batch conversion complete: 2 of 2 files were written to the output directory.",
    },
    "runtime_cancelled": {
        "zh_CN": "任务已取消：annual-summary-2026.docx",
        "zh_TW": "任務已取消：annual-summary-2026.docx",
        "en_US": "Task was cancelled: annual-summary-2026.docx",
    },
    "runtime_failed_active": {
        "zh_CN": "转换失败：annual-summary-2026.docx。无法读取文档结构，请查看失败详情。",
        "zh_TW": "轉換失敗：annual-summary-2026.docx。無法讀取文件結構，請查看失敗詳細資訊。",
        "en_US": "Conversion failed: annual-summary-2026.docx. The document structure could not be read; view failure details.",
    },
    "runtime_failed_queued": {
        "zh_CN": "转换失败：board-review-appendix.docx。文件受保护或已损坏。",
        "zh_TW": "轉換失敗：board-review-appendix.docx。檔案受保護或已損壞。",
        "en_US": "Conversion failed: board-review-appendix.docx. The file is protected or damaged.",
    },
    "runtime_failed_summary": {
        "zh_CN": "批量转换失败：2 / 2 个文件未完成，已生成失败详情。",
        "zh_TW": "批次轉換失敗：2 / 2 個檔案未完成，已產生失敗詳細資訊。",
        "en_US": "Batch conversion failed: 2 of 2 files did not finish; failure details were generated.",
    },
}


def _capture_sample_text(key: str) -> str:
    from docwen_gui.i18n import get_locale

    locale = get_locale()
    values = _CAPTURE_SAMPLE_TEXT[key]
    if locale in values:
        return values[locale]
    if locale.startswith("zh"):
        return values["zh_TW"]
    return values["en_US"]


def _sync_capture_category(window: QWidget, category: str) -> None:
    """Keep the synthetic screenshot scene's left category tab in sync."""
    tab_category = "text" if category in {"markdown", "text"} else category
    batch_vm = getattr(window, "_batch_list_vm", None)
    if batch_vm is not None and hasattr(batch_vm, "activate_tab"):
        batch_vm.activate_tab(tab_category)


def _sync_capture_mode(window: QWidget, vm: object, mode: str) -> None:
    """Keep synthetic screenshot scenes' main and input modes in sync."""
    if hasattr(vm, "set_mode"):
        vm.set_mode(mode)
    input_area = getattr(window, "_input_area", None)
    input_vm = getattr(input_area, "view_model", None)
    if input_vm is not None and hasattr(input_vm, "set_mode"):
        input_vm.set_mode(mode)


def _sync_capture_selected_file(window: QWidget, vm: object, file_ref: _FakeFileRef) -> None:
    """Keep synthetic screenshot scenes internally consistent."""
    if hasattr(vm, "set_selected_file"):
        vm.set_selected_file(file_ref)

    input_area = getattr(window, "_input_area", None)
    if input_area is None:
        return

    file_path = str(file_ref.path)
    with suppress(Exception):
        input_area._selected_file = file_path

    if Path(file_path).is_file() and hasattr(input_area, "update_display"):
        input_area.update_display(file_path)
        return

    display_vm = getattr(input_area, "view_model", None)
    if display_vm is None or not hasattr(display_vm, "_emit_message"):
        return

    from docwen_gui.i18n import t

    display_vm._emit_message(
        t(
            "components.file_drop.file_selected_msg",
            "Selected: {filename}",
            filename=Path(file_path).name,
        ),
        "success",
    )


# ── Main ────────────────────────────────────────────────────────────────────


def main() -> int:
    args = _parse_args()

    if args.list_themes:
        for theme_name in THEME_PRESETS:
            print(theme_name)
        return 0

    if args.lang:
        from docwen_gui.i18n import set_locale

        set_locale(args.lang)

    app = _ensure_app()
    theme_manager = ThemeManager.get_instance()
    theme_manager.initialize(app, "light")

    output_dir = Path(args.output_dir)
    main_size = QSize(args.width, args.height)
    captured_files: list[Path] = []

    for requested_theme in args.themes:
        theme_name = _normalize_theme(requested_theme)
        theme_manager.apply_theme(theme_name)

        if "main" in args.targets:
            captured_files.append(_capture_main_window(theme_name, output_dir, main_size))

        if "settings" in args.targets:
            captured_files.extend(_capture_settings_dialog(theme_name, output_dir, QSize(1080, 760)))

        if "about" in args.targets:
            captured_files.extend(_capture_about_dialog(theme_name, output_dir))

        if "batch" in args.targets:
            captured_files.append(_capture_batch_preview(theme_name, output_dir, QSize(720, 760)))

        if "status" in args.targets:
            captured_files.append(_capture_status_preview(theme_name, output_dir, QSize(760, 520)))

        if "progress" in args.targets:
            captured_files.append(_capture_runtime_status_scene(theme_name, output_dir, main_size, state="progress"))

        if "completed" in args.targets:
            captured_files.append(_capture_runtime_status_scene(theme_name, output_dir, main_size, state="completed"))

        if "failed" in args.targets:
            captured_files.append(_capture_runtime_status_scene(theme_name, output_dir, main_size, state="failed"))

        if "cancelled" in args.targets:
            captured_files.append(_capture_runtime_status_scene(theme_name, output_dir, main_size, state="cancelled"))

        if "template" in args.targets:
            captured_files.append(_capture_template_scene(theme_name, output_dir, main_size))

        if "conversion" in args.targets:
            captured_files.append(_capture_conversion_scene(theme_name, output_dir, main_size, legacy_name=True))

        for scenario in ("document", "spreadsheet", "image", "layout"):
            target = f"conversion-{scenario}"
            if target in args.targets:
                captured_files.append(_capture_conversion_scene(theme_name, output_dir, main_size, scenario=scenario))

    for file_path in captured_files:
        print(file_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
