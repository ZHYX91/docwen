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
