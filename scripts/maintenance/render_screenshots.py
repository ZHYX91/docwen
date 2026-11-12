"""Render actual DocWen widgets; no UI text replacement or image retouching."""

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import time
from datetime import UTC, datetime
from pathlib import Path

_SAMPLE_TEXT = {
    "ar_SA": (
        "تقرير المشروع",
        "التقدم",
        "اكتملت مراجعة المستندات والتحقق من البيانات.",
        "المهمة",
        "الحالة",
        "مراجعة المستندات",
        "مكتمل",
    ),
    "de_DE": (
        "Projektbericht",
        "Fortschritt",
        "Die Dokumentenprüfung und die Datenkontrolle sind abgeschlossen.",
        "Aufgabe",
        "Status",
        "Dokumentenprüfung",
        "Abgeschlossen",
    ),
    "en_US": (
        "Project report",
        "Progress",
        "Document review and data checks are complete.",
        "Task",
        "Status",
        "Document review",
        "Complete",
    ),
    "es_ES": (
        "Informe del proyecto",
        "Avances",
        "La revisión de documentos y la comprobación de datos han finalizado.",
        "Tarea",
        "Estado",
        "Revisión de documentos",
        "Completada",
    ),
    "fr_FR": (
        "Rapport de projet",
        "Avancement",
        "La révision des documents et la vérification des données sont terminées.",
        "Tâche",
        "État",
        "Révision des documents",
        "Terminée",
    ),
    "ja_JP": (
        "プロジェクト報告",
        "進捗状況",
        "文書の確認とデータの照合が完了しました。",
        "作業",
        "状況",
        "文書の確認",
        "完了",
    ),
    "ko_KR": (
        "프로젝트 보고서",
        "진행 상황",
        "문서 검토와 데이터 확인을 완료했습니다.",
        "작업",
        "상태",
        "문서 검토",
        "완료",
    ),
    "pt_BR": (
        "Relatório do projeto",
        "Andamento",
        "A revisão dos documentos e a conferência dos dados foram concluídas.",
        "Tarefa",
        "Status",
        "Revisão dos documentos",
        "Concluída",
    ),
    "ru_RU": (
        "Отчёт по проекту",
        "Ход работы",
        "Проверка документов и данных завершена.",
        "Задача",
        "Статус",
        "Проверка документов",
        "Завершена",
    ),
    "vi_VN": (
        "Báo cáo dự án",
        "Tiến độ",
        "Đã hoàn tất rà soát tài liệu và kiểm tra dữ liệu.",
        "Công việc",
        "Trạng thái",
        "Rà soát tài liệu",
        "Hoàn tất",
    ),
    "zh_CN": (
        "项目周报",
        "工作进展",
        "本周完成文档整理与数据核对。",
        "任务",
        "状态",
        "文档整理",
        "已完成",
    ),
    "zh_TW": (
        "專案週報",
        "工作進度",
        "本週完成文件整理與資料核對。",
        "工作",
        "狀態",
        "文件整理",
        "已完成",
    ),
}


def write_sample(directory: Path, locale: str, *, document: bool) -> Path:
    """Create a fictional input in the screenshot language, without an English fallback."""
    title, heading, summary, task_header, status_header, task, status = _SAMPLE_TEXT[locale]
    sample = directory / f"{title}{'.docx' if document else '.md'}"
    if document:
        from docx import Document

        source = Document()
        source.add_heading(title, 0)
        source.add_heading(heading, 1)
        source.add_paragraph(summary)
        table = source.add_table(rows=2, cols=2)
        for row, texts in zip(table.rows, ((task_header, status_header), (task, status)), strict=True):
            for cell, value in zip(row.cells, texts, strict=True):
                cell.text = value
        source.save(sample)
    else:
        sample.write_text(
            f"# {title}\n\n## {heading}\n\n{summary}\n\n"
            f"| {task_header} | {status_header} |\n| --- | --- |\n| {task} | {status} |\n",
            encoding="utf-8",
        )
    return sample


def main(argv: list[str] | None = None) -> None:
    tool_repo = Path(__file__).resolve().parents[2]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, default=tool_repo)
    parser.add_argument("--workspace-root", type=Path)
    parser.add_argument(
        "locale", choices=tuple(path.stem for path in sorted((tool_repo / "i18n/locales").glob("*.toml")))
    )
    parser.add_argument("theme", choices=("light", "dark"))
    parser.add_argument("--scale", type=int, choices=(90, 100, 110, 125, 150), default=100)
    parser.add_argument("--font", choices=("small", "default", "large"), default="default")
    parser.add_argument("--section", choices=("none", "templates", "general", "gongwen"), default="templates")
    parser.add_argument("--input", choices=("file", "document", "clipboard", "empty"), default="file")
    parser.add_argument("--width", type=int, default=1040)
    parser.add_argument("--height", type=int, default=820)
    parser.add_argument("--settings-width", type=int, default=960)
    parser.add_argument("--settings-height", type=int, default=880)
    parser.add_argument("--device-pixel-ratio", type=int, choices=(1, 2), default=1)
    parser.add_argument(
        "--show-recent-template",
        action="store_true",
        help="Populate sample session history with the selected real template",
    )
    parser.add_argument(
        "--filter-selected-template",
        action="store_true",
        help="Use the existing search field to show only the selected template",
    )
    parser.add_argument("--output", type=Path, required=True, help="Dedicated directory for reviewed media")
    args = parser.parse_args(argv)
    repo = args.repo.resolve(strict=True)
    args.output = args.output.resolve()
    sys.path.insert(0, str(tool_repo))
    from tools.run_lease import managed_run
    from tools.workspace_root import resolve_workspace_root

    workspace = resolve_workspace_root(repo, explicit=args.workspace_root)
    for source in sorted((repo / "packages").rglob("src")):
        if source.is_dir():
            sys.path.insert(0, str(source))
    with managed_run(
        workspace / "temp", prefix="docwen-qt-media-", owner="docwen.qt-store-media", kind="real-qt-widget-render"
    ) as run:
        previous_directory = Path.cwd()
        try:
            # Qt separates platform options with colons. A relative file name
            # in the leased run works even when the checkout is on another drive.
            os.chdir(run.root)
            _render(args, repo, run.root)
        finally:
            os.chdir(previous_directory)


def _render(args: argparse.Namespace, repo: Path, root: Path) -> None:
    locale, theme = args.locale, args.theme
    out = args.output.resolve()
    out.mkdir(parents=True, exist_ok=True)
    profile = root / "profile"
    (profile / "configs").mkdir(parents=True)
    cfg = (repo / "configs/gui.toml").read_text(encoding="utf-8-sig")
    for k, v in {
        "locale": f'"{locale}"',
        "default_theme": f'"{theme}"',
        "scale_percent": str(args.scale),
        "size_preset": f'"{args.font}"',
    }.items():
        cfg = re.sub(r"(?m)^" + k + r"\s*=.*$", f"{k} = {v}", cfg)
    (profile / "configs/gui.toml").write_text(cfg, encoding="utf-8")
    (root / "temp").mkdir()
    examples = root / "examples"
    screen_config = root / "offscreen-screen.json"
    screen_config.write_text(
        json.dumps(
            {
                "windowFrameMargins": True,
                "screens": [
                    {
                        "name": "media",
                        "x": 0,
                        "y": 0,
                        "width": max(1280, args.width + 80, args.settings_width + 80) * args.device_pixel_ratio,
                        "height": max(1024, args.height + 80, args.settings_height + 80) * args.device_pixel_ratio,
                        "logicalDpi": 96,
                        "logicalBaseDpi": 96,
                        "dpr": 1.0,
                    }
                ],
            }
        ),
        encoding="utf-8",
    )
    assert not examples.exists(), "Do not overwrite an existing example directory"
    examples.mkdir()
    sample = write_sample(examples, locale, document=args.input == "document")
    for k in ("DOCWEN_CONFIG_DIR", "DOCWEN_LOG_DIR", "DOCWEN_LOG_TO_TEMP", "QT_SCREEN_SCALE_FACTORS"):
        os.environ.pop(k, None)
    os.environ.update(
        DOCWEN_RUNTIME_ROOT=str(root / "runtime"),
        DOCWEN_DATA_DIR=str(profile),
        QT_SCALE_FACTOR=str(args.device_pixel_ratio),
        QT_QPA_PLATFORM="offscreen:configfile=" + screen_config.name,
        PYTHONDONTWRITEBYTECODE="1",
        DOCWEN_GUI_DISABLE_STATE_SAVE="1",
    )
    for k in ("TEMP", "TMP", "TMPDIR"):
        os.environ[k] = str(root / "temp")
    from PySide6.QtCore import QTimer
    from PySide6.QtWidgets import QApplication, QLineEdit

    from docwen_bundle import gui_entry
    from docwen_gui import app as gui_app

    original = gui_app.create_main_window
    main_label = "document" if args.input == "document" else "markdown"
    record = {
        "sourceCommit": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=repo, text=True).strip(),
        "locale": locale,
        "theme": theme,
        "scale": args.scale,
        "font": args.font,
        "input": args.input,
        "devicePixelRatio": args.device_pixel_ratio,
        "showRecentTemplate": args.show_recent_template,
        "filterSelectedTemplate": args.filter_selected_template,
        "requestedSizes": {
            main_label: {"width": args.width, "height": args.height},
            **(
                {args.section: {"width": args.settings_width, "height": args.settings_height}}
                if args.section != "none"
                else {}
            ),
        },
        "runtimeRoot": str(root),
        "sourceWorkingFiles": {},
        "method": "Real source application startup, QWidget.grab; original widgets, theme, controlled sample and actual template discovery. Clipboard input uses the isolated offscreen Qt clipboard and shipping paste handler, not native interaction evidence. No cursor or OS border; no image retouching.",
        "images": [],
    }
    for entry in (
        subprocess.check_output(["git", "status", "--porcelain=v1", "-z"], cwd=repo).decode("utf-8").split("\0")
    ):
        if not entry:
            continue
        relative = entry[3:]
        path = repo / relative
        if path.is_file():
            record["sourceWorkingFiles"][relative] = hashlib.sha256(path.read_bytes()).hexdigest()
    failed = []

    def capture(widget, label):
        widget.repaint()
        path = out / f"{locale}-{theme}-{label}.png"
        pix = widget.grab()
        requested = record["requestedSizes"][label]
        assert (widget.width(), widget.height()) == (requested["width"], requested["height"]), (
            label,
            "Requested widget size was not applied",
            requested,
            (widget.width(), widget.height()),
        )
        assert (pix.width(), pix.height(), pix.devicePixelRatio()) == (
            requested["width"] * args.device_pixel_ratio,
            requested["height"] * args.device_pixel_ratio,
            args.device_pixel_ratio,
        ), (label, "Unexpected captured pixels", pix.size(), pix.devicePixelRatio())
        assert pix.save(str(path), "PNG")
        record["images"].append(
            {
                "path": str(path),
                "logicalWidth": widget.width(),
                "logicalHeight": widget.height(),
                "width": pix.width(),
                "height": pix.height(),
                "devicePixelRatio": pix.devicePixelRatio(),
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            }
        )

    def make_window(**kwargs):
        window = original(**kwargs)

        def finish():
            try:
                capture(window._settings_dialog, args.section)
                window._settings_dialog.reject()
                window.close()
            except Exception as e:
                failed.append(repr(e))
                window.close()

        def settings():
            try:
                capture(window, main_label)
                if args.section == "none":
                    window.close()
                    return
                assert window.open_settings(None)["accepted"]
                assert window._settings_dialog.activate_section(args.section)
                if args.section == "templates":
                    selected = window._template_selector.get_selected_template_resource()
                    if selected is not None:
                        window._settings_dialog.focus_template("docx", selected[1])
                window._settings_dialog.resize(args.settings_width, args.settings_height)
                QTimer.singleShot(1200, finish)
            except Exception as e:
                failed.append(repr(e))
                window.close()

        def size():
            try:
                if args.input == "clipboard":
                    assert window._clipboard_store is None
                    window._clipboard_input_root = root / "clipboard"
                    QApplication.clipboard().setText(sample.read_text(encoding="utf-8"))
                    window._on_paste_requested(plain_text_only=True)
                deadline = time.monotonic() + 30
                ready_ticks = 0

                def wait_for_input():
                    nonlocal ready_ticks
                    try:
                        projection = window._view_model.ui_projection
                        ready = not window._view_model.inspection_busy and (
                            args.input == "empty"
                            or (args.input == "document" and projection.right_panel_visible)
                            or (
                                args.input != "document"
                                and projection.right_panel_visible
                                and window._template_selector.get_selected_template_resource() is not None
                            )
                        )
                        ready_ticks = ready_ticks + 1 if ready else 0
                        if ready_ticks >= 3:
                            if args.show_recent_template:
                                selected = window._template_selector.get_selected_template_resource()
                                assert selected is not None
                                window._template_selector.mark_used(selected[1])
                                record["sampleRecentTemplate"] = selected[1]
                            if args.filter_selected_template:
                                selected = window._template_selector.get_selected_template_resource()
                                assert selected is not None
                                searches = [
                                    field
                                    for field in window._template_selector.findChildren(
                                        QLineEdit, "templateSelectorSearch"
                                    )
                                    if field.isVisible()
                                ]
                                assert len(searches) == 1
                                label = window._template_selector.template_label(selected[1])
                                searches[0].setText(label)
                                record["sampleTemplateSearch"] = label
                            # Input inspection reveals the template panel and expands the
                            # normal window. Resize only after that transition settles;
                            # an earlier resize would become the collapsed-window width.
                            window.resize(args.width, args.height)
                            QTimer.singleShot(1500, settings)
                        elif time.monotonic() < deadline:
                            QTimer.singleShot(100, wait_for_input)
                        else:
                            raise TimeoutError("Input inspection and template projection did not settle")
                    except Exception as e:
                        failed.append(repr(e))
                        window.close()

                wait_for_input()
            except Exception as e:
                failed.append(repr(e))
                window.close()

        QTimer.singleShot(2500, size)
        return window

    gui_app.create_main_window = make_window
    state = "retained-failure"
    primary: BaseException | None = None
    try:
        code = gui_entry.main(["DocWen", str(sample)] if args.input in {"file", "document"} else ["DocWen"])
        expected_count = 1 if args.section == "none" else 2
        assert code == 0 and len(record["images"]) == expected_count and not failed, (code, failed)
        state = "completed-success"
    except BaseException as exc:
        primary = exc
        raise
    finally:
        gui_app.create_main_window = original
        record.update(at=datetime.now(UTC).isoformat(), state=state, errors=failed)
        try:
            (out / f"{locale}-{theme}.json").write_text(
                json.dumps(record, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
            )
        except OSError as exc:
            if primary is None:
                primary = exc
                raise
            primary.add_note(f"Could not save render evidence at {out}: {exc}")
        finally:
            try:
                import logging

                logging.shutdown()
                from loguru import logger

                logger.remove()
            except Exception as exc:
                if primary is None:
                    raise
                primary.add_note(f"Could not close render logging: {exc}")
    print(json.dumps({key: record[key] for key in ("at", "state", "images", "errors")}, ensure_ascii=False))


if __name__ == "__main__":
    main()
