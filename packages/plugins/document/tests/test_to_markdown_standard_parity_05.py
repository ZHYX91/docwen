"""Focused tests split from test_to_markdown_standard_parity.py."""

from __future__ import annotations

from ._to_markdown_standard_parity_support import (
    Document,
    DocxToMarkdownConverter,
    MagicMock,
    Path,
    _build_heading_formatter,
    _ExactSchemeRegistry,
    patch,
    pytest,
    re,
)

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    ("scheme", "registry", "error_type", "diagnostic_code"),
    [
        ("", _ExactSchemeRegistry(), "invalid_input", "NUMBERING-SCHEME-REQUIRED"),
        ("exact", None, "capability_unavailable", "NUMBERING-REGISTRY-UNAVAILABLE"),
        ("missing", _ExactSchemeRegistry(), "resource_not_found", "NUMBERING-SCHEME-NOT-FOUND"),
        (
            "exact",
            _ExactSchemeRegistry(enabled=False),
            "capability_unavailable",
            "NUMBERING-SCHEME-DISABLED",
        ),
        ("exact", _ExactSchemeRegistry(levels={}), "invalid_input", "NUMBERING-SCHEME-NO-LEVELS"),
    ],
)
def test_docx_to_md_rejects_unusable_exact_numbering_scheme(
    tmp_path: Path,
    scheme: str,
    registry: object,
    error_type: str,
    diagnostic_code: str,
) -> None:
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    source = tmp_path / "numbering-failure.docx"
    document = Document()
    document.add_heading("Heading", level=1)
    document.save(str(source))
    staging = tmp_path / "staging"
    staging.mkdir()
    context = FakeExecutionContext(
        request=ConversionRequest(
            request_id="docx-numbering-failure",
            input_refs=[FileRef(path=str(source), format="docx", category="document")],
            target_format="md",
            options={"add_numbering": True, "numbering_scheme": scheme},
            output_policy=OutputPolicy(),
        ),
        workspace=FakeWorkspaceHandle(str(source), str(staging)),
        config=FakeConfigView(),
        progress=FakeProgressSink(),
        cancellation=CancellationToken().view(),
        logger=FakePluginLogger(),
        numbering_registry=registry,
    )

    result = DocxToMarkdownConverter().convert(context)

    assert not result.success
    assert result.error is not None
    assert result.error.error_type == error_type
    assert result.error.diagnostic_code == diagnostic_code
    assert result.artifacts == []


def test_add_numbering_gongwen_standard():
    """add_numbering with gongwen_standard → Chinese numbering."""
    doc = Document()
    h1 = doc.add_heading("Introduction", level=1)
    h2 = doc.add_heading("Background", level=2)
    h1b = doc.add_heading("Methods", level=1)

    fmt = _build_heading_formatter("gongwen_standard")
    converter = DocxToMarkdownConverter()

    lines1, _ = converter._process_paragraph(
        h1._element,
        {id(h1._element): h1},
        remove_numbering=True,
        heading_formatter=fmt,
    )
    assert "# 一、Introduction" in "\n".join(lines1)

    lines2, _ = converter._process_paragraph(
        h2._element,
        {id(h2._element): h2},
        remove_numbering=True,
        heading_formatter=fmt,
    )
    assert "## （一）Background" in "\n".join(lines2)

    lines3, _ = converter._process_paragraph(
        h1b._element,
        {id(h1b._element): h1b},
        remove_numbering=True,
        heading_formatter=fmt,
    )
    assert "# 二、Methods" in "\n".join(lines3)


def test_add_numbering_hierarchical_standard():
    """add_numbering with hierarchical_standard → 1, 1.1, 2 numbering."""
    doc = Document()
    h1 = doc.add_heading("Introduction", level=1)
    h2 = doc.add_heading("Background", level=2)
    h1b = doc.add_heading("Methods", level=1)

    fmt = _build_heading_formatter("hierarchical_standard")
    converter = DocxToMarkdownConverter()

    lines1, _ = converter._process_paragraph(
        h1._element,
        {id(h1._element): h1},
        remove_numbering=True,
        heading_formatter=fmt,
    )
    assert "# 1 Introduction" in "\n".join(lines1)

    lines2, _ = converter._process_paragraph(
        h2._element,
        {id(h2._element): h2},
        remove_numbering=True,
        heading_formatter=fmt,
    )
    assert "## 1.1 Background" in "\n".join(lines2)

    lines3, _ = converter._process_paragraph(
        h1b._element,
        {id(h1b._element): h1b},
        remove_numbering=True,
        heading_formatter=fmt,
    )
    assert "# 2 Methods" in "\n".join(lines3)


def test_add_numbering_hierarchical_h2_start():
    """hierarchical_h2_start → level 1 no numbering, level 2 starts at 1."""
    doc = Document()
    h1 = doc.add_heading("Introduction", level=1)
    h2 = doc.add_heading("Background", level=2)
    h1b = doc.add_heading("Methods", level=1)

    fmt = _build_heading_formatter("hierarchical_h2_start")
    converter = DocxToMarkdownConverter()

    lines1, _ = converter._process_paragraph(
        h1._element,
        {id(h1._element): h1},
        remove_numbering=True,
        heading_formatter=fmt,
    )
    assert "# Introduction" in "\n".join(lines1)

    lines2, _ = converter._process_paragraph(
        h2._element,
        {id(h2._element): h2},
        remove_numbering=True,
        heading_formatter=fmt,
    )
    assert "## 1 Background" in "\n".join(lines2)

    lines3, _ = converter._process_paragraph(
        h1b._element,
        {id(h1b._element): h1b},
        remove_numbering=True,
        heading_formatter=fmt,
    )
    assert "# Methods" in "\n".join(lines3)


def test_add_numbering_legal_standard():
    """legal_standard → 第一编, 第一章, 第二编 numbering."""
    doc = Document()
    h1 = doc.add_heading("Introduction", level=1)
    h2 = doc.add_heading("Background", level=2)
    h1b = doc.add_heading("Methods", level=1)

    fmt = _build_heading_formatter("legal_standard")
    converter = DocxToMarkdownConverter()

    lines1, _ = converter._process_paragraph(
        h1._element,
        {id(h1._element): h1},
        remove_numbering=True,
        heading_formatter=fmt,
    )
    assert "# 第一编　Introduction" in "\n".join(lines1)

    lines2, _ = converter._process_paragraph(
        h2._element,
        {id(h2._element): h2},
        remove_numbering=True,
        heading_formatter=fmt,
    )
    assert "## 第一章　Background" in "\n".join(lines2)

    lines3, _ = converter._process_paragraph(
        h1b._element,
        {id(h1b._element): h1b},
        remove_numbering=True,
        heading_formatter=fmt,
    )
    assert "# 第二编　Methods" in "\n".join(lines3)


def test_add_numbering_remove_then_add():
    """remove_numbering=True strips existing prefix, add_numbering adds new one."""
    doc = Document()
    para = doc.add_paragraph("一、Existing Title")

    mock_style = MagicMock()
    mock_style.name = "Heading 1"
    mock_style.style_id = "1Heading1"

    fmt = _build_heading_formatter("hierarchical_standard")
    converter = DocxToMarkdownConverter()

    cleanup_rules = (("chinese_顿号", re.compile(r"^[一二三四五六七八九十百千万]+、"), 1),)
    with patch.object(type(para), "style", property(lambda s: mock_style)):
        lines, _ = converter._process_paragraph(
            para._element,
            {id(para._element): para},
            remove_numbering=True,
            heading_formatter=fmt,
            heading_cleanup_rules=cleanup_rules,
        )
    output = "\n".join(lines)
    assert "# 1 Existing Title" in output


def test_add_numbering_false_no_change():
    """add_numbering=False → no numbering added (existing behavior preserved)."""
    doc = Document()
    para = doc.add_heading("Plain Heading", level=1)

    converter = DocxToMarkdownConverter()
    lines, _ = converter._process_paragraph(
        para._element,
        {id(para._element): para},
        remove_numbering=False,
        heading_formatter=None,
    )
    output = "\n".join(lines)
    assert "# Plain Heading" in output
