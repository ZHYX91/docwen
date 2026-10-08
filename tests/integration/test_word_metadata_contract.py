"""Compare the closed SDT formatting profile with the pinned standard types."""

from copy import deepcopy
from importlib.resources import files

import pytest
from lxml import etree

from docwen_core.docx_host_metadata import has_owned_tag_properties
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
