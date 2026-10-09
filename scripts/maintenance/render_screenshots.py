"""Render actual DocWen widgets; no UI text replacement or image retouching."""

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
from datetime import UTC, datetime
from pathlib import Path


def main(argv: list[str] | None = None) -> None:
    tool_repo = Path(__file__).resolve().parents[2]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, default=tool_repo)
    parser.add_argument("--workspace-root", type=Path)
    parser.add_argument("locale", choices=("en_US", "zh_CN"))
    parser.add_argument("theme", choices=("light", "dark"))
    parser.add_argument("--scale", type=int, choices=(90, 100, 110, 125, 150), default=100)
    parser.add_argument("--font", choices=("small", "default", "large"), default="default")
    parser.add_argument("--section", choices=("templates", "general"), default="templates")
    parser.add_argument("--input", choices=("file", "clipboard"), default="file")
    parser.add_argument("--width", type=int, default=1040)
    parser.add_argument("--height", type=int, default=820)
    parser.add_argument("--settings-width", type=int, default=960)
    parser.add_argument("--settings-height", type=int, default=880)
    parser.add_argument("--output", type=Path, required=True, help="Dedicated directory for reviewed media")
    args = parser.parse_args(argv)
    repo = args.repo.resolve(strict=True)
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
        _render(args, repo, run.root)


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
        "center_panel_width": "620",
        "default_height": "900",
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
                        "width": max(1280, args.width + 80, args.settings_width + 80),
                        "height": max(1024, args.height + 80, args.settings_height + 80),
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
    sample = examples / ("项目周报.md" if locale == "zh_CN" else "Project report.md")
    sample.write_text(
        "# 项目周报\n\n## 工作进展\n\n本周完成文档整理与数据核对。\n\n| 任务 | 状态 |\n| --- | --- |\n| 文档整理 | 已完成 |\n"
        if locale == "zh_CN"
        else "# Project report\n\n## Progress\n\nDocument review and data checks are complete.\n\n| Task | Status |\n| --- | --- |\n| Document review | Complete |\n",
        encoding="utf-8",
    )
    for k in ("DOCWEN_CONFIG_DIR", "DOCWEN_LOG_DIR", "DOCWEN_LOG_TO_TEMP", "QT_SCREEN_SCALE_FACTORS"):
        os.environ.pop(k, None)
    os.environ.update(
        DOCWEN_RUNTIME_ROOT=str(root / "runtime"),
        DOCWEN_DATA_DIR=str(profile),
        QT_SCALE_FACTOR="1",
        QT_QPA_PLATFORM="offscreen:configfile=" + os.path.relpath(screen_config, repo).replace("\\", "/"),
        PYTHONDONTWRITEBYTECODE="1",
        DOCWEN_GUI_DISABLE_STATE_SAVE="1",
    )
    for k in ("TEMP", "TMP", "TMPDIR"):
        os.environ[k] = str(root / "temp")
    os.chdir(repo)
    from PySide6.QtCore import QTimer
    from PySide6.QtWidgets import QApplication

    from docwen_bundle import gui_entry
    from docwen_gui import app as gui_app

    original = gui_app.create_main_window
    record = {
        "sourceCommit": subprocess.check_output(["git", "rev-parse", "HEAD"], text=True).strip(),
        "locale": locale,
        "theme": theme,
        "scale": args.scale,
        "font": args.font,
        "input": args.input,
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
        assert pix.save(str(path), "PNG")
        record["images"].append(
            {
                "path": str(path),
                "width": pix.width(),
                "height": pix.height(),
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
                capture(window, "markdown")
                assert window.open_settings(None)["accepted"]
                assert window._settings_dialog.activate_section(args.section)
                if args.section == "templates":
                    selected = window._template_selector.get_selected_template_resource()
                    window._settings_dialog.focus_template("docx", selected[1])
                window._settings_dialog.resize(args.settings_width, args.settings_height)
                QTimer.singleShot(1200, finish)
            except Exception as e:
                failed.append(repr(e))
                window.close()

        def size():
            if args.input == "clipboard":
                assert window._clipboard_store is None
                window._clipboard_input_root = root / "clipboard"
                QApplication.clipboard().setText(sample.read_text(encoding="utf-8"))
                window._on_paste_requested(plain_text_only=True)
            window.resize(args.width, args.height)
            QTimer.singleShot(1500, settings)

        QTimer.singleShot(2500, size)
        return window

    gui_app.create_main_window = make_window
    state = "retained-failure"
    primary: BaseException | None = None
    try:
        code = gui_entry.main(["DocWen", str(sample)] if args.input == "file" else ["DocWen"])
        assert code == 0 and len(record["images"]) == 2 and not failed, (code, failed)
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
