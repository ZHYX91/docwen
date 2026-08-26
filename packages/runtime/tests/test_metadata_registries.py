"""Runtime registry tests for templates and numbering schemes."""

from __future__ import annotations

import pytest
from tests.support.numbering import repository_numbering_registry

pytestmark = pytest.mark.unit


def test_template_registry_discovers_real_templates() -> None:
    from docwen_runtime.templates import TemplateRegistry

    templates = TemplateRegistry.default().list_templates()
    names = {template.name for template in templates}

    assert "简体中文通用模板" in names
    assert "English General Template" in names
    assert any(template.target == "xlsx" for template in templates)


def test_template_registry_filters_by_target() -> None:
    from docwen_runtime.templates import TemplateRegistry

    docx_templates = TemplateRegistry.default().list_templates("docx")
    assert docx_templates
    assert {template.target for template in docx_templates} == {"docx"}


def test_template_registry_gets_template_by_canonical_id() -> None:
    from docwen_runtime.templates import TemplateRegistry

    registry = TemplateRegistry.default()
    expected = next(item for item in registry.list_templates("docx") if item.name == "简体中文通用模板")
    template = registry.get_template(expected.id)

    assert template.target == "docx"
    assert template.path.name == "简体中文通用模板.docx"
    assert template.size_bytes > 0


def test_resource_root_discovery_supports_package_adjacent_resources(tmp_path) -> None:
    from docwen_runtime.resources.registry import _find_root_from_module_path

    package_root = tmp_path / "site-packages" / "docwen_runtime"
    module_file = package_root / "resources" / "registry.py"
    module_file.parent.mkdir(parents=True)
    module_file.write_text("", encoding="utf-8")
    (package_root / "templates").mkdir()

    assert _find_root_from_module_path(module_file) == package_root


def test_resource_root_discovery_skips_inaccessible_ancestors(tmp_path, monkeypatch) -> None:
    from docwen_runtime.resources import registry

    package_root = tmp_path / "site-packages" / "docwen_runtime"
    module_file = package_root / "resources" / "registry.py"
    module_file.parent.mkdir(parents=True)
    module_file.write_text("", encoding="utf-8")
    (package_root / "configs").mkdir()

    real_exists = registry.Path.exists

    def guarded_exists(candidate):
        if candidate.name == "templates" and candidate.parent == tmp_path:
            raise PermissionError("simulated inaccessible ancestor")
        return real_exists(candidate)

    monkeypatch.setattr(registry.Path, "exists", guarded_exists)

    assert registry._find_root_from_module_path(module_file) == package_root


def test_numbering_registry_explicit_snapshot_ignores_flat_heading_numbering_file(tmp_path) -> None:
    from docwen_runtime.config.loader import ConfigLoader
    from docwen_runtime.config.registry import CONFIG_FILES
    from docwen_runtime.numbering import NumberingSchemeRegistry

    root = tmp_path
    configs_dir = root / "configs"
    (configs_dir / "numbering").mkdir(parents=True)
    (configs_dir / "proofread").mkdir(parents=True)
    (root / "i18n" / "locales").mkdir(parents=True)
    # Create all base files
    for spec in CONFIG_FILES:
        path = configs_dir / spec.rel_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n", encoding="utf-8")
    # Write the numbering add data under the new registry path
    (configs_dir / "numbering" / "add.toml").write_text(
        '[settings]\norder = ["markdown_only"]\n'
        "[schemes.markdown_only]\n"
        'name = "Markdown Only"\n'
        'description = "canonical"\n'
        "enabled = true\n"
        "is_system = false\n"
        'locales = ["*"]\n'
        "[schemes.markdown_only.level_1]\n"
        'format = "{1.arabic_half} "\n',
        encoding="utf-8",
    )
    # A removed flat config location must not become an implicit source again.
    (configs_dir / "heading_numbering_add.toml").write_text(
        '[settings]\norder = ["legacy_only"]\n'
        "[schemes.legacy_only]\n"
        'name = "Legacy Only"\n'
        'description = "legacy"\n'
        "enabled = true\n"
        "is_system = false\n"
        'locales = ["*"]\n'
        "[schemes.legacy_only.level_1]\n"
        'format = "{1.arabic_half} "\n',
        encoding="utf-8",
    )
    (root / "i18n" / "locales" / "zh_CN.toml").write_text("", encoding="utf-8")

    loader = ConfigLoader(base_dir=configs_dir, user_dir=root / "user")
    registry = NumberingSchemeRegistry.from_config_snapshot(
        loader.config.as_dict(),
        locale_path=root / "i18n" / "locales" / "zh_CN.toml",
    )

    ids = [scheme.scheme_id for scheme in registry.list_schemes()]

    assert "markdown_only" in ids
    assert "legacy_only" not in ids


def test_numbering_registry_discovers_real_schemes() -> None:
    schemes = repository_numbering_registry().list_schemes()
    ids = [scheme.scheme_id for scheme in schemes]

    assert ids[:4] == [
        "gongwen_standard",
        "hierarchical_standard",
        "hierarchical_h2_start",
        "legal_standard",
    ]


def test_numbering_registry_resolves_translation_keys_to_display_text() -> None:
    scheme = repository_numbering_registry().get_scheme("hierarchical_standard")

    assert scheme.name == "层级数字标准"
    assert scheme.description != "hierarchical_standard_desc"
    assert "层级递进格式" in scheme.description


def test_numbering_registry_resolves_translation_keys_from_requested_locale() -> None:
    scheme = repository_numbering_registry(locale="en_US").get_scheme("hierarchical_standard")

    assert scheme.name == "Hierarchical Number Standard"
    assert scheme.description != "hierarchical_standard_desc"
    assert "Hierarchical format" in scheme.description


def test_numbering_registry_uses_explicit_snapshot_and_locale_path(tmp_path) -> None:
    from docwen_runtime.config.loader import ConfigLoader
    from docwen_runtime.config.registry import CONFIG_FILES
    from docwen_runtime.numbering import NumberingSchemeRegistry

    root = tmp_path
    configs_dir = root / "configs"
    user_dir = root / "user"
    configs_dir.mkdir()
    user_dir.mkdir()
    (configs_dir / "numbering").mkdir(parents=True)
    (configs_dir / "proofread").mkdir(parents=True)
    (root / "i18n" / "locales").mkdir(parents=True)
    for spec in CONFIG_FILES:
        path = configs_dir / spec.rel_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n", encoding="utf-8")
    # Write locale under gui.toml (new registry namespace)
    (configs_dir / "gui.toml").write_text(
        '[language]\nlocale = "en_US"\n',
        encoding="utf-8",
    )
    (configs_dir / "numbering" / "add.toml").write_text(
        '[settings]\norder = ["hierarchical_standard"]\n'
        "[schemes.hierarchical_standard]\n"
        'name_key = "hierarchical_standard"\n'
        'description_key = "hierarchical_standard_desc"\n'
        "enabled = true\n"
        "is_system = true\n"
        'locales = ["*"]\n'
        "[schemes.hierarchical_standard.level_1]\n"
        'format = "{1.arabic_half} "\n',
        encoding="utf-8",
    )
    (root / "i18n" / "locales" / "en_US.toml").write_text(
        "[cli.numbering_schemes]\n"
        'hierarchical_standard = "Hierarchical Number Standard"\n'
        "[editors.numbering_add.descriptions]\n"
        'hierarchical_standard_desc = "Hierarchical format from active locale"\n',
        encoding="utf-8",
    )
    loader = ConfigLoader(base_dir=configs_dir, user_dir=user_dir)
    registry = NumberingSchemeRegistry.from_config_snapshot(
        loader.config.as_dict(),
        locale_path=root / "i18n" / "locales" / "en_US.toml",
    )

    scheme = registry.get_scheme("hierarchical_standard")

    assert scheme.name == "Hierarchical Number Standard"
    assert scheme.description == "Hierarchical format from active locale"


def test_numbering_registry_gets_scheme_details() -> None:
    scheme = repository_numbering_registry().get_scheme("gongwen_standard")

    assert scheme.name == "公文标准"
    assert scheme.enabled is True
    assert scheme.levels["level_1"] == "{1.chinese_lower}、"


def test_numbering_registry_filters_locale_specific_schemes_and_honors_wildcard(tmp_path) -> None:
    from docwen_runtime.numbering import NumberingSchemeRegistry

    snapshot = {
        "gui": {"language": {"locale": "zh_CN"}},
        "numbering": {
            "add": {
                "settings": {"order": ["universal", "english", "chinese", "unrestricted"]},
                "schemes": {
                    "universal": {"name": "Universal", "locales": ["*"]},
                    "english": {"name": "English", "locales": ["en_US"]},
                    "chinese": {"name": "Chinese", "locales": ["zh_CN"]},
                    "unrestricted": {"name": "Unrestricted"},
                },
            }
        },
    }
    zh_locale = tmp_path / "zh_CN.toml"
    en_locale = tmp_path / "en_US.toml"
    zh_locale.write_text("", encoding="utf-8")
    en_locale.write_text("", encoding="utf-8")
    registry = NumberingSchemeRegistry.from_config_snapshot(snapshot, locale_path=zh_locale)

    assert [scheme.scheme_id for scheme in registry.list_schemes()] == [
        "universal",
        "chinese",
        "unrestricted",
    ]
    assert [scheme.scheme_id for scheme in registry.with_config_snapshot(snapshot, locale="en_US").list_schemes()] == [
        "universal",
        "english",
        "unrestricted",
    ]


def test_numbering_registry_from_config_snapshot_is_immutable_and_loader_independent(tmp_path) -> None:
    from docwen_runtime.numbering import NumberingSchemeRegistry

    snapshot = {
        "gui": {"language": {"locale": "zh_CN"}},
        "numbering": {
            "add": {
                "settings": {"order": ["request_scheme"]},
                "schemes": {
                    "request_scheme": {
                        "name": "Request scheme",
                        "description": "frozen",
                        "enabled": True,
                        "is_system": False,
                        "locales": ["*"],
                        "level_1": {"format": "REQ-{1.arabic_half} "},
                    }
                },
            }
        },
    }

    locale_path = tmp_path / "zh_CN.toml"
    locale_path.write_text("", encoding="utf-8")
    registry = NumberingSchemeRegistry.from_config_snapshot(
        snapshot,
        locale_path=locale_path,
    )
    snapshot["numbering"]["add"]["schemes"]["request_scheme"]["name"] = "mutated"
    snapshot["numbering"]["add"]["settings"]["order"] = []

    scheme = registry.get_scheme("request_scheme")
    assert scheme.name == "Request scheme"
    assert scheme.levels["level_1"] == "REQ-{1.arabic_half} "


def test_numbering_registry_with_snapshot_reuses_owned_locale_root(tmp_path) -> None:
    from docwen_runtime.numbering import NumberingSchemeRegistry

    locales = tmp_path / "locales"
    locales.mkdir()
    zh_locale = locales / "zh_CN.toml"
    zh_locale.write_text(
        '[cli.numbering_schemes]\nrequest_scheme = "中文方案"\n',
        encoding="utf-8",
    )
    (locales / "en_US.toml").write_text(
        '[cli.numbering_schemes]\nrequest_scheme = "English scheme"\n',
        encoding="utf-8",
    )
    owner = NumberingSchemeRegistry.from_config_snapshot({}, locale_path=zh_locale)
    snapshot = {
        "numbering": {
            "add": {
                "settings": {"order": ["request_scheme"]},
                "schemes": {
                    "request_scheme": {
                        "name_key": "request_scheme",
                        "enabled": True,
                        "level_1": {"format": "{1.arabic_half}. "},
                    }
                },
            }
        }
    }

    localized = owner.with_config_snapshot(snapshot, locale="en_US")
    fallback = owner.with_config_snapshot(snapshot, locale="missing_LOCALE")

    assert owner.list_schemes() == []
    assert localized.get_scheme("request_scheme").name == "English scheme"
    assert fallback.get_scheme("request_scheme").name == "中文方案"
