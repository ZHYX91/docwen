from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit

__all__ = (
    "pytest",
    "pytestmark",
)


@pytest.fixture(autouse=True)
def synthetic_navigation_resources(monkeypatch):
    """Exercise all registered pages with explicitly multilingual synthetic resources."""
    from tests.support.gui_vm_fakes import FakeController, optimization_capability_projection

    from docwen_application.optimization_selection import OptimizationSource
    from docwen_gui.i18n import get_locale
    from docwen_gui.view_models._optimization_filter import discover_optimization_choices
    from docwen_gui.view_models.settings_vm import SettingsViewModel

    controller = FakeController()

    def projection():
        value = optimization_capability_projection()
        locales = [
            "zh_CN",
            "zh_TW",
            "en_US",
            "de_DE",
            "fr_FR",
            "es_ES",
            "pt_BR",
            "ru_RU",
            "ja_JP",
            "ko_KR",
            "vi_VN",
            "ar_SA",
        ]
        value["optimization_applicable_locales"] = {
            "gongwen": locales,
            "invoice_cn": locales,
        }
        return value

    monkeypatch.setattr(controller, "describe_gui_runtime_capabilities", projection)

    def choices(self, locale=None, source_category=None):
        return discover_optimization_choices(
            controller,
            locale=locale or get_locale(),
            sources=(OptimizationSource(detected_format="", source_category=source_category),)
            if source_category
            else (),
        )

    monkeypatch.setattr(SettingsViewModel, "get_optimization_choices_result", choices)
