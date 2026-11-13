"""One catalog owner shared by settings and the conversion selector."""

from __future__ import annotations

from PySide6.QtCore import QObject, Signal

from docwen_gui.qt_bridge.background_operation import BackgroundOperation
from docwen_runtime.templates import TemplateInfo, TemplateManager


class TemplateViewModel(QObject):
    changed = Signal()
    failed = Signal(str)

    def __init__(self, parent: QObject | None = None, *, manager: TemplateManager | None = None) -> None:
        super().__init__(parent)
        self.manager = manager or TemplateManager.default()
        self.templates: tuple[TemplateInfo, ...] = ()
        self.enabled: dict[str, bool] = {}
        self.defaults: dict[str, str | None] = {}
        self.error: str | None = None
        self._refresh_operation = BackgroundOperation(self)

    @staticmethod
    def _read_catalog(manager, preferred_name):
        with manager.state_store.locked():
            templates = tuple(manager.list_templates())
            state = manager.state_store.load()
            enabled = {item.id: state["enabled"].get(item.id, True) for item in templates}
            defaults = dict(state["defaults"])
            if not state["defaults_initialized"]:
                for target in ("docx", "xlsx"):
                    candidates = [
                        item
                        for item in templates
                        if item.target == target and enabled[item.id] and not manager.is_custom(item)
                    ]
                    chosen = next(
                        (item for item in candidates if item.name == preferred_name),
                        candidates[0] if candidates else None,
                    )
                    defaults[target] = chosen.id if chosen else None
            else:
                for target, selected in defaults.items():
                    if selected is not None and not enabled.get(selected, False):
                        defaults[target] = None
            if defaults != state["defaults"] or not state["defaults_initialized"]:
                state["defaults"] = defaults
                state["defaults_initialized"] = True
                manager.state_store.save(state)
        return templates, enabled, defaults

    @property
    def busy(self) -> bool:
        return self._refresh_operation.busy

    def cancel_refresh(self) -> None:
        self._refresh_operation.cancel()

    def refresh_async(self) -> None:
        from docwen_gui.i18n import t

        manager = self.manager
        preferred_name = t("meta.template_name")

        def work(token):
            token.check()
            result = self._read_catalog(manager, preferred_name)
            token.check()
            return result

        self._refresh_operation.submit(work, self._receive_catalog)

    def refresh(self) -> None:
        from docwen_gui.i18n import t

        self.cancel_refresh()
        try:
            result = self._read_catalog(self.manager, t("meta.template_name"))
        except Exception as exc:
            self._receive_catalog(None, exc)
        else:
            self._receive_catalog(result, None)

    def _receive_catalog(self, result, error) -> None:
        try:
            if error is not None:
                raise error
            templates, enabled, defaults = result
            changed = (templates, enabled, defaults) != (self.templates, self.enabled, self.defaults)
            self.templates, self.enabled, self.defaults = templates, enabled, defaults
            was_failed = self.error is not None
            self.error = None
            if changed or was_failed:
                self.changed.emit()
        except Exception as exc:
            self.error = str(exc)
            self.templates = ()
            self.failed.emit(self.error)

    def find(self, template_id: str) -> TemplateInfo | None:
        return next((item for item in self.templates if item.id == template_id), None)
