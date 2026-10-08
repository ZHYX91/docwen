"""Compare the closed SDT formatting profile with the pinned standard types."""

from copy import deepcopy
from importlib.resources import files

import pytest
from lxml import etree

from docwen_core.docx_host_metadata import has_owned_tag_properties
from docwen_plugin_document.shared import word_metadata
from docwen_plugin_document.shared.word_metadata import is_valid_word_metadata

pytestmark = pytest.mark.integration
_Q = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _enum_cases():
    schemas = [
        etree.fromstring(files("docwen_plugin_document").joinpath("resources", "word_metadata", name).read_bytes())
        for name in ("wml.xsd", "shared-commonSimpleTypes.xsd")
    ]
    cases = []
    for name, attribute, type_name, base in [
        ("bdr", "val", "ST_Border", {}),
        ("u", "val", "ST_Underline", {}),
        ("shd", "val", "ST_Shd", {}),
        ("effect", "val", "ST_TextEffect", {}),
        ("em", "val", "ST_Em", {}),
        ("highlight", "val", "ST_HighlightColor", {}),
        ("vertAlign", "val", "ST_VerticalAlignRun", {}),
        ("color", "themeColor", "ST_ThemeColor", {"val": "auto"}),
        ("rFonts", "cstheme", "ST_Theme", {}),
        ("rFonts", "hint", "ST_Hint", {}),
    ]:
        values = [
            value
            for schema in schemas
            for value in schema.xpath(
                f'x:simpleType[@name="{type_name}"]//x:enumeration/@value',
                namespaces={"x": "http://www.w3.org/2001/XMLSchema"},
            )
        ]
        assert values, type_name
        for value in [*values, "not-a-schema-value"]:
            cases.append((name, {**base, attribute: value}))
    return cases


@pytest.mark.parametrize("name,attributes", _enum_cases())
def test_supported_formatting_enumerations_match_official_schema(name, attributes):
    _assert_profile_matches_schema(name, attributes)


@pytest.mark.parametrize(
    "name,attribute,base",
    [
        ("b", "val", {}),
        ("sz", "val", {}),
        ("position", "val", {}),
        ("w", "val", {}),
        ("color", "val", {}),
        ("color", "themeTint", {"val": "auto"}),
        ("fitText", "val", {}),
        ("eastAsianLayout", "id", {}),
        ("bdr", "space", {"val": "single"}),
    ],
)
@pytest.mark.parametrize(
    "value",
    [
        "",
        "0",
        "-0",
        "+2",
        " 2 ",
        "2 pt",
        "2pt",
        " 2pt ",
        "50%",
        " 50% ",
        "false",
        " false ",
        " off ",
        "FF",
        " FF ",
        "auto",
        " auto ",
        "18446744073709551616",
    ],
)
def test_supported_formatting_lexical_members_match_official_schema(name, attribute, base, value):
    _assert_profile_matches_schema(name, {**base, attribute: value})


def _assert_profile_matches_schema(name, attributes):
    properties = etree.Element(_Q + "rPr")
    etree.SubElement(properties, _Q + name, {_Q + key: value for key, value in attributes.items()})
    control = etree.Element(_Q + "sdtPr")
    control.append(deepcopy(properties))
    etree.SubElement(control, _Q + "tag", {_Q + "val": "owned"})
    assert has_owned_tag_properties(control, "owned") == is_valid_word_metadata(properties), etree.tostring(properties)


def test_font_hint_cs_is_not_a_transitional_hint():
    _assert_profile_matches_schema("rFonts", {"hint": "cs"})


@pytest.mark.parametrize("nested", [False, True])
@pytest.mark.parametrize("name", ["vanish", "b", "sz"])
def test_duplicate_run_properties_cannot_establish_metadata_validity(nested, name):
    properties = etree.Element(_Q + "rPr")
    target = properties
    if nested:
        change = etree.SubElement(properties, _Q + "rPrChange", {_Q + "id": "1", _Q + "author": "Editor"})
        target = etree.SubElement(change, _Q + "rPr")
    values = ("2", "3") if name == "sz" else ("false", "true")
    for value in values:
        etree.SubElement(target, _Q + name, {_Q + "val": value})
    assert not is_valid_word_metadata(properties)


def test_sdt_end_properties_can_contain_multiple_distinct_run_property_sets():
    end = etree.Element(_Q + "sdtEndPr")
    for name in ("b", "i"):
        etree.SubElement(etree.SubElement(end, _Q + "rPr"), _Q + name)
    assert is_valid_word_metadata(end)


@pytest.mark.parametrize(
    "name,attribute",
    [
        ("sz", "val"),
        ("szCs", "val"),
        ("kern", "val"),
        ("fitText", "val"),
        ("tabIndex", "val"),
        ("bdr", "sz"),
        ("bdr", "space"),
    ],
)
@pytest.mark.parametrize("value", ["+2", "-0", " +2 ", " -0 "])
def test_unsigned_metadata_policy_survives_a_permissive_validator(monkeypatch, name, attribute, value):
    # Model the newer validator's acceptance explicitly: the product policy must
    # still reject signs rather than depend on this host's libxml2 behavior.
    class PermissiveSchema:
        def validate(self, element):
            return True

    monkeypatch.setattr(word_metadata, "_metadata_schema", lambda: PermissiveSchema())
    properties = etree.Element(_Q + ("sdtPr" if name == "tabIndex" else "rPr"))
    attributes = {_Q + attribute: value}
    if name == "bdr":
        attributes[_Q + "val"] = "single"
    etree.SubElement(properties, _Q + name, attributes)
    assert not is_valid_word_metadata(properties)


@pytest.mark.parametrize("value", ["0", " 02 ", "18446744073709551615"])
def test_unsigned_sdt_tab_index_remains_valid(value):
    properties = etree.Element(_Q + "sdtPr")
    etree.SubElement(properties, _Q + "tabIndex", {_Q + "val": value})
    assert is_valid_word_metadata(properties)
