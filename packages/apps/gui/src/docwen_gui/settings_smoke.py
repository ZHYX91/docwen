"""Internal packaged-GUI probe for eager Settings page construction."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from PySide6.QtWidgets import QApplication

    from .main_window import MainWindow


def _schedule_test_settings_report(app: QApplication, window: MainWindow) -> None:
    """Write a settings-page construction report when requested by release tests.

    The production dialog intentionally imports each page at runtime so one page
    can fail without taking down the rest.  This probe exercises that exact path
    inside the frozen executable and reports placeholders as failures.
    """

    report_path_raw = os.environ.get("DOCWEN_GUI_TEST_SETTINGS_REPORT", "").strip()
    if not report_path_raw:
        return
    report_path = Path(report_path_raw)

    def _probe() -> None:
        dialog = None
        report: dict[str, object] = {
            "success": False,
            "expectedTabs": [],
            "loadedTabs": [],
            "failedTabs": [],
            "missingTabs": [],
            "unexpectedTabs": [],
            "pageObjectNames": {},
            "error": None,
        }
        try:
            from .view_models.settings_vm import SettingsViewModel
            from .widgets.settings.dialog import TAB_KEYS, SettingsDialog

            controller = getattr(getattr(window, "view_model", None), "controller", None)
            dialog = SettingsDialog(parent=window, view_model=SettingsViewModel(controller=controller))
            pages = dict(dialog._tabs)  # pyright: ignore[reportPrivateUsage]
            expected_tabs = list(TAB_KEYS)
            missing_tabs = [key for key in expected_tabs if key not in pages]
            unexpected_tabs = [key for key in pages if key not in TAB_KEYS]
            object_names = {key: page.objectName() for key, page in pages.items()}
            failed_tabs = [key for key in expected_tabs if object_names.get(key) == "settingsTabLoadErrorPage"]
            loaded_tabs = [key for key in expected_tabs if key in pages and key not in failed_tabs]
            report.update(
                {
                    "success": not missing_tabs and not unexpected_tabs and not failed_tabs,
                    "expectedTabs": expected_tabs,
                    "loadedTabs": loaded_tabs,
                    "failedTabs": failed_tabs,
                    "missingTabs": missing_tabs,
                    "unexpectedTabs": unexpected_tabs,
                    "pageObjectNames": object_names,
                }
            )
        except Exception as exc:
            report["error"] = f"{type(exc).__name__}: {exc}"
        finally:
            if dialog is not None:
                dialog.close()
                dialog.deleteLater()
            report_path.parent.mkdir(parents=True, exist_ok=True)
            report_path.write_text(
                f"{json.dumps(report, ensure_ascii=False, indent=2)}\n",
                encoding="utf-8",
            )

    from PySide6.QtCore import QTimer

    QTimer.singleShot(0, _probe)


__all__ = ["_schedule_test_settings_report"]
