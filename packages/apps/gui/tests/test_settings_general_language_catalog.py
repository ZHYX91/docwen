from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.gui

PROJECT_ROOT = Path(__file__).resolve().parents[4]
PROJECT_CONFIGS = PROJECT_ROOT / "configs"
LOCALES_DIR = PROJECT_ROOT / "i18n" / "locales"
PACKAGED_LOCALES = tuple(sorted(path.stem for path in LOCALES_DIR.glob("*.toml")))


def _new_tab(user_dir: Path):
    tab, _view_model, _config_port = _new_tab_context(user_dir)
    return tab


def _new_tab_context(user_dir: Path):
    from docwen_application.controller import ApplicationController
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.general_tab import GeneralTab

    config_port = ConfigPortAdapter(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
    view_model = SettingsViewModel(controller=ApplicationController(config_port=config_port))
    return GeneralTab(view_model), view_model, config_port


def _combo_codes(tab) -> tuple[str, ...]:
    combo = tab._language_combo  # pyright: ignore[reportPrivateUsage]
    return tuple(str(combo.itemData(index)) for index in range(combo.count()))


def test_general_language_codes_equal_packaged_locale_stems(qapp, tmp_path: Path) -> None:
    tab = _new_tab(tmp_path / "user")
    codes = _combo_codes(tab)

    assert len(codes) == len(set(codes))
    assert set(codes) == set(PACKAGED_LOCALES)


@pytest.mark.parametrize("active_locale", PACKAGED_LOCALES)
def test_general_language_labels_resolve_for_every_packaged_locale(
    qapp,
    tmp_path: Path,
    active_locale: str,
) -> None:
    from docwen_gui.i18n import get_locale, set_locale, t

    previous_locale = get_locale()
    set_locale(active_locale)
    try:
        tab = _new_tab(tmp_path / "user")
        combo = tab._language_combo  # pyright: ignore[reportPrivateUsage]
        labels = {str(combo.itemData(index)): combo.itemText(index) for index in range(combo.count())}

        assert set(labels) == set(PACKAGED_LOCALES)
        for code in PACKAGED_LOCALES:
            resolved = t(f"settings.general.languages.{code}", "")
            assert resolved.strip(), f"{active_locale} has no language label for {code}"
            assert labels[code] == resolved
    finally:
        set_locale(previous_locale)


def test_general_language_user_selection_persists_through_apply(qapp, tmp_path: Path) -> None:
    user_dir = tmp_path / "user"
    tab, view_model, config_port = _new_tab_context(user_dir)
    combo = tab._language_combo  # pyright: ignore[reportPrivateUsage]
    initial = view_model.config.gui.language
    target = next(locale for locale in PACKAGED_LOCALES if locale != initial)

    index = combo.findData(target)
    assert index >= 0
    combo.setCurrentIndex(index)

    assert view_model.config.gui.language == target
    assert view_model.is_dirty is True
    assert view_model.apply_changes() is True
    config_port.reload()
    assert config_port.get("gui.language.locale") == target

    reloaded = _new_tab(user_dir)
    assert reloaded._language_combo.currentData() == target  # pyright: ignore[reportPrivateUsage]
