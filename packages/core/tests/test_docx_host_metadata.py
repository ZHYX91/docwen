"""Word's control formatting must not change ownership proofs."""

import pytest
from docx.oxml import parse_xml

from docwen_core.docx_host_metadata import has_owned_tag_properties

pytestmark = pytest.mark.contract

_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_FORMAT = (
    '<w:rPr><w:b/><w:bCs/><w:color w:val="4F81BD" w:themeColor="accent1"/>'
    '<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
)


def test_word_saved_control_formatting_preserves_owned_tag():
    properties = parse_xml(f'<w:sdtPr {_NS}>{_FORMAT}<w:tag w:val="owned"/><w:id w:val="-42"/></w:sdtPr>')
    assert has_owned_tag_properties(properties, "owned")
    assert not has_owned_tag_properties(properties, "other")


@pytest.mark.parametrize(
    "extra",
    [
        '<w:tag w:val="owned"/>',
        '<w:id w:val="3"/>',
        '<w:lock w:val="sdtLocked"/>',
        "<w:rPr/>",
    ],
)
def test_formatting_does_not_relax_control_identity(extra):
    properties = parse_xml(f'<w:sdtPr {_NS}>{_FORMAT}<w:tag w:val="owned"/><w:id w:val="-42"/>{extra}</w:sdtPr>')
    assert not has_owned_tag_properties(properties, "owned")


@pytest.mark.parametrize(
    "content",
    ["<w:t>hidden payload</w:t>", "<w:b><w:t>payload</w:t></w:b>", "<w:rPrChange/>", "<w:b/>text", "<w:b/><w:b/>"],
)
def test_control_formatting_rejects_content_and_revision_containers(content):
    properties = parse_xml(f'<w:sdtPr {_NS}><w:rPr>{content}</w:rPr><w:tag w:val="owned"/></w:sdtPr>')
    assert not has_owned_tag_properties(properties, "owned")


@pytest.mark.parametrize(
    "content",
    [
        '<w:b w:val="perhaps"/>',
        '<w:b w:color="FF0000"/>',
        '<w:sz w:val="not-a-number"/>',
        '<w:szCs w:val="-18"/>',
        "<w:sz/>",
        '<w:sz w:val="18" w:themeColor="accent1"/>',
        '<w:color w:val="red"/>',
        '<w:color w:val="12345"/>',
        '<w:color w:themeColor="unknown"/>',
        '<w:color w:themeColor="accent1" w:themeTint="XYZ"/>',
        '<w:color w:themeColor="accent1" w:themeShade="FFF"/>',
        '<w:color w:val="auto" w:unknown="1"/>',
    ],
)
def test_control_defaults_reject_invalid_attribute_values(content):
    properties = parse_xml(f'<w:sdtPr {_NS}><w:rPr>{content}</w:rPr><w:tag w:val="owned"/></w:sdtPr>')
    assert not has_owned_tag_properties(properties, "owned")


@pytest.mark.parametrize(
    "content",
    [
        '<w:b w:val="off"/><w:i w:val="1"/><w:vanish w:val="false"/>',
        '<w:sz w:val="18"/><w:szCs w:val="24"/>',
        '<w:color w:val="auto"/>',
        '<w:color w:val="aBcD12"/>',
        '<w:color w:val="auto" w:themeColor="accent1"/>',
        '<w:color w:val="auto" w:themeColor="text1" w:themeTint="FF" w:themeShade="a0"/>',
    ],
)
def test_control_defaults_keep_valid_word_values(content):
    properties = parse_xml(f'<w:sdtPr {_NS}><w:rPr>{content}</w:rPr><w:tag w:val="owned"/></w:sdtPr>')
    assert has_owned_tag_properties(properties, "owned")


@pytest.mark.parametrize(
    "content,valid",
    [
        ('<w:effect w:val="sparkle"/>', True),
        ('<w:effect w:val="unknown"/>', False),
        ('<w:em w:val="underDot"/>', True),
        ('<w:em w:val="unknown"/>', False),
        ('<w:highlight w:val="yellow"/>', True),
        ('<w:highlight w:val="unknown"/>', False),
        ('<w:vertAlign w:val="superscript"/>', True),
        ('<w:vertAlign w:val="top"/>', False),
        ('<w:position w:val="-2"/><w:spacing w:val="-1.5pt"/>', True),
        ('<w:position w:val="two"/>', False),
        ("<w:spacing/>", False),
        ("<w:w/>", True),
        ('<w:w w:val="600"/>', True),
        ('<w:w w:val="601"/>', False),
        ('<w:fitText w:val="31680" w:id="-1"/>', True),
        ('<w:fitText w:val="31681"/>', True),
        ('<w:fitText w:val="1" w:id="2147483648"/>', True),
        ('<w:rStyle w:val="Strong"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/>', True),
        ("<w:rStyle/>", False),
        ('<w:rFonts w:ascii="Arial" w:cstheme="minorBidi" w:hint="eastAsia"/>', True),
        ('<w:rFonts w:asciiTheme="not-a-theme"/>', False),
        ('<w:eastAsianLayout w:id="1" w:combine="on" w:combineBrackets="round"/>', True),
        ('<w:eastAsianLayout w:combine="perhaps"/>', False),
        ('<w:eastAsianLayout w:combineBrackets="unknown"/>', False),
        ('<w:u w:val="wave" w:color="auto" w:themeColor="accent2"/>', True),
        ('<w:u w:val="unknown"/>', False),
        ('<w:u w:color="red"/>', False),
        ('<w:shd w:val="pct25" w:fill="ABCDEF" w:themeFill="background1"/>', True),
        ('<w:shd w:val="unknown"/>', False),
        ('<w:shd w:val="clear" w:themeFill="unknown"/>', False),
        ('<w:bdr w:val="single" w:sz="8" w:space="1" w:shadow="false"/>', True),
        ('<w:bdr w:val="zigZagStitch"/>', True),
        ('<w:bdr w:val="unknown"/>', False),
        ('<w:bdr w:val="single" w:space="32"/>', True),
        ('<w:bdr w:val="single" w:sz="-1"/>', False),
    ],
)
def test_control_property_families_validate_values_and_unknown_attributes(content, valid):
    properties = parse_xml(f'<w:sdtPr {_NS}><w:rPr>{content}</w:rPr><w:tag w:val="owned"/></w:sdtPr>')
    assert has_owned_tag_properties(properties, "owned") is valid
    properties[0][0].set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}unexpected", "1")
    assert not has_owned_tag_properties(properties, "owned")


@pytest.mark.parametrize(
    "content,valid",
    [
        ('<w:color w:themeColor="accent1"/>', False),
        ("<w:color/>", False),
        ('<w:shd w:val="clear" w:themeTint="F"/>', False),
        ('<w:u w:themeShade="F"/>', False),
        ('<w:bdr w:val="single" w:themeTint="F"/>', False),
        ('<w:shd w:val="clear" w:themeTint="0F"/>', True),
        ('<w:sz w:val="18446744073709551615"/>', True),
        ('<w:sz w:val="18446744073709551616"/>', False),
        ('<w:szCs w:val="18446744073709551616"/>', False),
        ('<w:kern w:val="18446744073709551616"/>', False),
        ('<w:sz w:val="+000000000000000000018"/>', True),
        ('<w:sz w:val="-0"/>', True),
        ('<w:sz w:val=" 18 "/>', True),
        ('<w:sz w:val="12.5pt"/>', True),
        ('<w:rFonts w:ascii="A long font family name beyond thirty one characters"/>', True),
        (f'<w:lang w:val="{"a" * 85}"/>', True),
        ('<w:fitText w:val="12.5pt"/>', True),
        ('<w:fitText w:val="18446744073709551616"/>', False),
        ('<w:eastAsianLayout w:id="2147483648"/>', True),
        ('<w:w w:val="0"/>', True),
        ('<w:w w:val="050%"/>', True),
        ('<w:w w:val="601%"/>', False),
        ('<w:w w:val="00000000000000000000100"/>', True),
        ('<w:bdr w:val="single" w:sz="18446744073709551615"/>', True),
        ('<w:bdr w:val="single" w:space="18446744073709551616"/>', False),
    ],
)
def test_control_defaults_follow_transitional_schema_not_sdk_limits(content, valid):
    properties = parse_xml(f'<w:sdtPr {_NS}><w:rPr>{content}</w:rPr><w:tag w:val="owned"/></w:sdtPr>')
    assert has_owned_tag_properties(properties, "owned") is valid
