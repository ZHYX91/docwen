import pytest

from docwen_core.docx_parsing.xml_ns import (
    NS_A,
    NS_M,
    NS_MC,
    NS_PIC,
    NS_R,
    NS_V,
    NS_W,
    NS_WP,
    NSMAP_M,
    NSMAP_W,
    NSMAP_WPV,
)

pytestmark = pytest.mark.unit


def test_namespace_constants_match_ooxml_uris():
    assert NS_W == "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    assert NS_M == "http://schemas.openxmlformats.org/officeDocument/2006/math"
    assert NS_WP == "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    assert NS_A == "http://schemas.openxmlformats.org/drawingml/2006/main"
    assert NS_R == "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    assert NS_PIC == "http://schemas.openxmlformats.org/drawingml/2006/picture"
    assert NS_V == "urn:schemas-microsoft-com:vml"
    assert NS_MC == "http://schemas.openxmlformats.org/markup-compatibility/2006"


def test_namespace_maps_are_ready_for_lxml_findall():
    assert NSMAP_W == {"w": NS_W}
    assert NSMAP_M["m"] == NS_M
    assert NSMAP_M["w"] == NS_W
    assert NSMAP_WPV["w"] == NS_W
    assert NSMAP_WPV["wp"] == NS_WP
    assert NSMAP_WPV["a"] == NS_A
    assert NSMAP_WPV["r"] == NS_R
    assert NSMAP_WPV["v"] == NS_V
    assert NSMAP_WPV["pic"] == NS_PIC
