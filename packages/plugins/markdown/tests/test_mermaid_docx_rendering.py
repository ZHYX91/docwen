from __future__ import annotations

from io import BytesIO

import pytest
from docx import Document
from docx.oxml.ns import qn
from PIL import Image

from docwen_plugin_markdown.renderer import MdToDocxRenderer

pytestmark = pytest.mark.unit


def _png_bytes() -> bytes:
    stream = BytesIO()
    Image.new("RGB", (160, 80), "white").save(stream, format="PNG")
    return stream.getvalue()


def _mermaid_node() -> dict:
    return {
        "type": "block_code",
        "attrs": {"info": "mermaid"},
        "raw": "graph TD\nA-->B\n",
    }


def test_mermaid_image_mode_embeds_png_instead_of_visible_source() -> None:
    rendered_sources: list[str] = []
    renderer = MdToDocxRenderer(
        Document(),
        mermaid_mode="image",
        mermaid_render=lambda source: rendered_sources.append(source) or _png_bytes(),
    )

    paragraph = renderer._handle_block_code(_mermaid_node())

    assert rendered_sources == ["graph TD\nA-->B\n"]
    assert paragraph.text == ""
    assert paragraph._p.find(".//" + qn("w:drawing")) is not None
    assert renderer.warnings == ()


def test_mermaid_image_failure_falls_back_to_code_with_warning() -> None:
    def fail(_source: str) -> bytes:
        raise RuntimeError("renderer unavailable")

    renderer = MdToDocxRenderer(
        Document(),
        mermaid_mode="image",
        mermaid_render=fail,
    )

    paragraph = renderer._handle_block_code(_mermaid_node())

    assert paragraph.text == "graph TD\nA-->B"
    assert renderer.warnings[0][0] == "MD2DOCX-MERMAID-FALLBACK"
    assert "renderer unavailable" in renderer.warnings[0][1]


def test_non_mermaid_fence_never_calls_mermaid_renderer() -> None:
    def unexpected(_source: str) -> bytes:
        raise AssertionError("renderer must not be called")

    renderer = MdToDocxRenderer(
        Document(),
        mermaid_mode="image",
        mermaid_render=unexpected,
    )
    node = {"type": "block_code", "attrs": {"info": "python"}, "raw": "print('ok')\n"}

    paragraph = renderer._handle_block_code(node)

    assert paragraph.text == "print('ok')"
    assert renderer.warnings == ()
