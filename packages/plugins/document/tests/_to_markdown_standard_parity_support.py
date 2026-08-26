import re
import tomllib
from contextlib import nullcontext
from pathlib import Path
from types import SimpleNamespace
from typing import Any, cast
from unittest.mock import MagicMock, patch

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Twips
from docx.styles.style import ParagraphStyle
from lxml import etree
from tests.support.numbering import repository_numbering_registry

from docwen_core.docx_parsing.format_features import DocxMarkdownSyntaxConfig
from docwen_plugin_document.shared.note_extraction import NoteExtractor
from docwen_plugin_document.to_markdown.converter import DocxToMarkdownConverter

pytestmark = pytest.mark.contract


class _ExactSchemeRegistry:
    def __init__(self, *, enabled: bool = True, levels: dict[str, str] | None = None) -> None:
        self._scheme = SimpleNamespace(
            enabled=enabled,
            levels={"level_1": "{1.arabic_half} "} if levels is None else levels,
        )

    def get_scheme(self, scheme_id: str) -> object:
        if scheme_id != "exact":
            raise LookupError(scheme_id)
        return self._scheme


def _inject_numpr(para: Any, *, num_id: str = "9", ilvl: int = 0) -> None:
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        para._p.insert(0, pPr)
    existing = pPr.find(qn("w:numPr"))
    if existing is not None:
        pPr.remove(existing)

    numPr = OxmlElement("w:numPr")
    ilvl_elem = OxmlElement("w:ilvl")
    ilvl_elem.set(qn("w:val"), str(ilvl))
    num_id_elem = OxmlElement("w:numId")
    num_id_elem.set(qn("w:val"), num_id)
    numPr.append(ilvl_elem)
    numPr.append(num_id_elem)
    pPr.append(numPr)


def _convert_document_fixture_to_markdown(tmp_path: Path, document: Any, *, request_id: str) -> str:
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    source = tmp_path / f"{request_id}.docx"
    document.save(str(source))
    staging = tmp_path / f"{request_id}-staging"
    staging.mkdir()
    context = FakeExecutionContext(
        request=ConversionRequest(
            request_id=request_id,
            input_refs=[FileRef(path=str(source), format="docx", category="document")],
            target_format="md",
            options={},
            output_policy=OutputPolicy(),
        ),
        workspace=FakeWorkspaceHandle(str(source), str(staging)),
        config=FakeConfigView(),
        progress=FakeProgressSink(),
        cancellation=CancellationToken().view(),
        logger=FakePluginLogger(),
        numbering_registry=None,
    )
    result = DocxToMarkdownConverter().convert(context)
    assert result.success, result.error.message if result.error else "conversion failed"
    return Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")


def _inject_outline_level(para: Any, level: int) -> None:
    pPr = para._p.get_or_add_pPr()
    existing = pPr.find(qn("w:outlineLvl"))
    if existing is not None:
        pPr.remove(existing)
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), str(level))
    pPr.append(outline)


def _parse_markdown_ast(markdown: str) -> list[dict[str, Any]]:
    import mistune

    parsed = cast(list[dict[str, Any]], mistune.create_markdown(renderer="ast")(markdown))
    return [node for node in parsed if node["type"] != "blank_line"]


_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _append_formula(parent: Any, text: str) -> None:
    omath = OxmlElement("m:oMath")
    math_run = OxmlElement("m:r")
    math_text = OxmlElement("m:t")
    math_text.text = text
    math_run.append(math_text)
    omath.append(math_run)
    parent.append(omath)


def _append_alternate_content_formula(
    paragraph: Any,
    *,
    choice_text: str | None,
    fallback_text: str,
) -> None:
    alternate = etree.Element(f"{{{_MC_NS}}}AlternateContent", nsmap={"mc": _MC_NS})
    choice = etree.SubElement(alternate, f"{{{_MC_NS}}}Choice")
    choice.set("Requires", "m")
    if choice_text is not None:
        _append_formula(choice, choice_text)
    fallback = etree.SubElement(alternate, f"{{{_MC_NS}}}Fallback")
    _append_formula(fallback, fallback_text)
    paragraph._p.append(alternate)


def _mock_para_style(para, name: str):
    """Return a context manager that patches ``para.style`` to return
    an object whose ``.name`` is *name*."""
    mock_style = MagicMock()
    mock_style.name = name
    return patch.object(type(para), "style", property(lambda s: mock_style))


def _inject_pPr_without_numPr(para):
    """Ensure the paragraph has a ``<w:pPr>`` element without ``<w:numPr>``,
    so that ``detect_list_item`` reaches the pStyle fallback path."""
    from docx.oxml.ns import qn

    existing = para._p.find(qn("w:pPr"))
    if existing is None:
        pPr = OxmlElement("w:pPr")
        para._p.insert(0, pPr)
    else:
        numPr = existing.find(qn("w:numPr"))
        if numPr is not None:
            existing.remove(numPr)


def _make_pstyle_numbering_index(
    style_id: str = "1Heading1",
    num_fmt: str = "chineseCountingThousand",
    lvl_text: str = "%1、",
    num_id: str = "10",
    abs_id: str = "0",
    ilvl: int = 0,
):
    """Build a NumberingIndex with a known pStyle mapping."""
    from docwen_plugin_document.shared.numbering_index import NumberingIndex

    idx = NumberingIndex.__new__(NumberingIndex)
    idx._num_to_abstract = {num_id: abs_id}
    idx._abstract_num_style_links = {}
    idx._abstract_levels = {
        abs_id: {
            ilvl: {"numFmt": num_fmt, "lvlText": lvl_text, "pStyle": style_id},
        },
    }
    return idx


def _build_heading_formatter(scheme_id: str) -> Any:
    """Build a HeadingFormatter for the given numbering scheme ID."""
    registry = repository_numbering_registry()
    scheme_info = registry.get_scheme(scheme_id)
    scheme_config: dict[str, Any] = {}
    for key, fmt in scheme_info.levels.items():
        scheme_config[key] = {"format": fmt}
    from docwen_core.text.heading_numbering import HeadingFormatter

    return HeadingFormatter(scheme_config)


__all__ = (
    "WD_BREAK",
    "WD_COLOR_INDEX",
    "WD_STYLE_TYPE",
    "_MC_NS",
    "Any",
    "Document",
    "DocxMarkdownSyntaxConfig",
    "DocxToMarkdownConverter",
    "MagicMock",
    "NoteExtractor",
    "OxmlElement",
    "ParagraphStyle",
    "Path",
    "Twips",
    "_ExactSchemeRegistry",
    "_append_alternate_content_formula",
    "_append_formula",
    "_build_heading_formatter",
    "_convert_document_fixture_to_markdown",
    "_inject_numpr",
    "_inject_outline_level",
    "_inject_pPr_without_numPr",
    "_make_pstyle_numbering_index",
    "_mock_para_style",
    "_parse_markdown_ast",
    "cast",
    "etree",
    "nullcontext",
    "patch",
    "pytest",
    "pytestmark",
    "qn",
    "re",
    "tomllib",
)
