from __future__ import annotations

from pathlib import Path

import pytest
import tomlkit

from docwen.config.toml_operations import read_toml_file, validate_toml_syntax
from docwen.i18n.locale_validator import get_reference_locale_path


LOCALES_DIR = Path(__file__).resolve().parents[2] / "src" / "docwen" / "i18n" / "locales"


pytestmark = pytest.mark.unit


def test_locales_dir_exists() -> None:
    assert LOCALES_DIR.is_dir()


def test_all_locales_toml_parse_and_have_meta() -> None:
    toml_files = sorted(LOCALES_DIR.glob("*.toml"))
    assert toml_files, "No locale TOML files found"

    for toml_path in toml_files:
        assert validate_toml_syntax(toml_path) is True, f"Invalid TOML syntax: {toml_path.name}"
        data = read_toml_file(toml_path)
        assert "meta" in data, f"Missing [meta] in {toml_path.name}"
        assert "name" in data["meta"], f"Missing meta.name in {toml_path.name}"
        assert "native_name" in data["meta"], f"Missing meta.native_name in {toml_path.name}"


def test_key_groups_order_and_styles_key_order() -> None:
    reference_path = get_reference_locale_path()
    reference_text = reference_path.read_text(encoding="utf-8")
    reference_doc = tomlkit.parse(reference_text)
    reference_styles_keys = list(reference_doc.get("styles", {}).keys())

    key_sections = ["styles", "style_formats", "placeholders", "yaml_keys"]
    missing_reference_sections = [name for name in key_sections if name not in reference_doc]
    assert not missing_reference_sections, f"Missing required sections in {reference_path.name}: {missing_reference_sections}"

    reference_section_order = [k for k in reference_doc.keys() if k in key_sections]
    assert reference_section_order == key_sections

    toml_files = sorted(LOCALES_DIR.glob("*.toml"))
    for toml_path in toml_files:
        text = toml_path.read_text(encoding="utf-8")
        actual_doc = tomlkit.parse(text)
        missing = [name for name in key_sections if name not in actual_doc]
        assert not missing, f"Missing sections in {toml_path.name}: {missing}"

        actual_section_order = [k for k in actual_doc.keys() if k in key_sections]
        assert actual_section_order == reference_section_order, (
            f"Section order mismatch in {toml_path.name}: {actual_section_order} != {reference_section_order}"
        )

        actual_styles_table = actual_doc.get("styles", {})
        actual_styles_keys = list(actual_styles_table.keys())
        if toml_path == reference_path:
            assert actual_styles_keys == reference_styles_keys
            continue

        positions_in_styles = {k: actual_styles_keys.index(k) for k in reference_styles_keys if k in actual_styles_keys}
        missing_style_keys = [k for k in reference_styles_keys if k not in positions_in_styles]
        assert not missing_style_keys, f"Missing styles keys in {toml_path.name}: {missing_style_keys}"

        last = -1
        for key in reference_styles_keys:
            idx = positions_in_styles[key]
            assert idx > last, f"Styles key order incorrect in {toml_path.name}: {key}"
            last = idx

        extras = [k for k in actual_styles_keys if k not in set(reference_styles_keys)]
        if extras:
            extra_positions = [actual_styles_keys.index(k) for k in extras]
            assert all(p > last for p in extra_positions), f"Extra styles keys must be appended in {toml_path.name}"
