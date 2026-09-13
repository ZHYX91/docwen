"""One catalog owner shared by settings and the conversion selector."""

from __future__ import annotations

from PySide6.QtCore import QObject, Signal

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

    def refresh(self) -> None:
        try:
            with self.manager.state_store.locked():
                templates = tuple(self.manager.list_templates())
                state = self.manager.state_store.load()
                enabled = {item.id: state["enabled"].get(item.id, True) for item in templates}
                defaults = dict(state["defaults"])
                if not state["defaults_initialized"]:
                    from docwen_gui.i18n import t

                    for target in ("docx", "xlsx"):
                        candidates = [
                            item
                            for item in templates
                            if item.target == target and enabled[item.id] and not self.manager.is_custom(item)
                        ]
                        chosen = next(
                            (item for item in candidates if item.name == t("meta.template_name")),
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
                    self.manager.state_store.save(state)
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
