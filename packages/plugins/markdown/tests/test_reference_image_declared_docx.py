"""Declared reference images reach the DOCX media part selected by each binding."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any

import pytest
from docx import Document
from docx.document import Document as DocumentObject
from PIL import Image
from tests.support.cancellation import FakeCancellationTokenView
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

from docwen_core.links.declared_resources import DeclaredResourceError, DeclaredResourceResolver
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_plugin_markdown.mistune_extensions import create_extended_markdown
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_runtime.config import build_document_style_catalog

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    "token,definition",
    [
        ("![ref][image]", "[image]: active.png\n"),
        ("![image][]", "[image]: active.png\n"),
        ("![image]", "[image]: active.png\n"),
        ("![ref][IMAGE   NAME]", '[image\nname]:\n <active.png>\n "title"\n'),
        ("![ref][image]", '[image]: active.png "title\ncontinued"\n'),
        ("![ref][image]", "[image]: active.png\n[IMAGE]: ignored.png\n"),
        ("![image]", "[^note]: text\n[image]: active.png\n"),
        ("![image]", "[^note]:\n    body\n[image]: active.png\n"),
    ],
)
def test_reference_images_embed_the_authenticated_resources(
    token: str, definition: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = token + "\n\n![other](active.png)\n\n" + definition
    _assert_authenticated_images(source, token, tmp_path, monkeypatch)


@pytest.mark.parametrize(
    "source,token,alt",
    [
        ("> ![im\n> age]\n>\n> [im\n> age]: active.png\n", "![im\n> age]", "im\nage"),
        ("> ![ref][im\n> age]\n>\n> [im age]: active.png\n", "![ref][im\n> age]", "ref"),
        ("> ![im\n> age][]\n>\n> [im age]: active.png\n", "![im\n> age][]", "im\nage"),
        ("- ![im\n  age]\n\n  [im age]: active.png\n", "![im\n  age]", "im\nage"),
        ("> ![im\nage]\n\n[im age]: active.png\n", "![im\nage]", "im\nage"),
        ("> - ![im\n>   age]\n>\n> [im age]: active.png\n", "![im\n>   age]", "im\nage"),
        ("- [im age]: active.png\n  ![im\n  age]\n", "![im\n  age]", "im\nage"),
        ("> [im age]: active.png\n> ![im\n> age]\n", "![im\n> age]", "im\nage"),
        (
            "> ![im\n> a > ge]\n>\n> [im a > ge]: active.png\n",
            "![im\n> a > ge]",
            "im\na > ge",
        ),
    ],
)
@pytest.mark.parametrize("eol", ["\n", "\r\n"])
def test_container_reference_images_keep_source_tokens_and_embed_the_correct_media(
    source: str, token: str, alt: str, eol: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = (source + "\n![other](active.png)\n").replace("\n", eol)
    token = token.replace("\n", eol)
    document = _assert_authenticated_images(source, token, tmp_path, monkeypatch)
    assert document.inline_shapes[0]._inline.docPr.get("descr") == alt.replace("\n", " ")


def _assert_authenticated_images(
    source: str, token: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> DocumentObject:
    source_path = tmp_path / "source.md"
    source_path.write_bytes(source.encode("utf-8"))
    resources = []
    hashes = []
    for directory, color in [("assets", "green"), ("other", "red")]:
        image = tmp_path / f"{directory}.png"
        Image.new("RGB", (6, 4), color).save(image)
        hashes.append(hashlib.sha256(image.read_bytes()).hexdigest())
        resources.append(
            FileRef(
                path=str(image),
                format="png",
                category="image",
                input_kind="resource",
                input_role="linked_resource",
                logical_path=f"{directory}/active.png",
                media_type="image/png",
            )
        )
    source_ref = FileRef(
        path=str(source_path),
        format="markdown",
        category="document",
        input_kind="document",
        input_role="source",
        logical_path="notes/note.md",
        media_type="text/markdown",
    )
    request = ConversionRequest(
        request_id="reference-images",
        input_refs=[source_ref, *resources],
        target_format="docx",
        output_policy=OutputPolicy(),
        options={
            "remove_numbering": False,
            "markdown_resource_bindings": {
                "authored_sha256": hashlib.sha256(source.encode()).hexdigest(),
                "images": [
                    {"authored_token": token, "logical_path": "assets/active.png"},
                    {"authored_token": "![other](active.png)", "logical_path": "other/active.png"},
                ],
            },
        },
    )
    context = FakeExecutionContext(
        request,
        FakeWorkspaceHandle(str(source_path), str(tmp_path), tuple(request.input_refs)),
        FakeConfigView(),
        FakeProgressSink(),
        FakeCancellationTokenView(),
        FakePluginLogger(),
        document_style_catalog=build_document_style_catalog({}),
    )

    def forbidden(*args: object, **kwargs: object) -> None:
        raise AssertionError("declared image attempted filesystem discovery")

    monkeypatch.setattr("docwen_core.links._embed_dispatch.resolve_file_path", forbidden)

    def images(nodes: list[dict[str, Any]]) -> list[dict[str, Any]]:
        return [
            image
            for node in nodes
            for image in ([node] if node["type"] == "image" else images(node.get("children", [])))
        ]

    ast = create_extended_markdown()(source)
    assert isinstance(ast, list)
    assert [image["attrs"]["url"] for image in images(ast)] == ["active.png", "active.png"]
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    output = next(artifact.staging_path for artifact in result.artifacts if artifact.kind == "primary")
    document = Document(output)
    actual = []
    for shape in document.inline_shapes:
        relationship_id = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
        assert relationship_id is not None
        actual.append(hashlib.sha256(document.part.related_parts[relationship_id].blob).hexdigest())
    assert actual == hashes
    assert source_path.read_bytes() == source.encode("utf-8")
    return document


@pytest.mark.parametrize(
    "destination,accepted",
    [
        (r"<a\>b.png>", False),
        (r"<a\(b.png>", False),
        ("<a b.png>", True),
        ("a(b.png", True),
        ("a)b.png", True),
        ("a.png junk", False),
        ("a\x00b.png", False),
    ],
)
def test_reference_destination_qualification_matches_the_frozen_renderer(destination: str, accepted: bool) -> None:
    # Mistune's reference grammar currently rejects backslashes in angle
    # targets and accepts unbalanced parentheses in bare block destinations.
    # Source authentication must not silently broaden that renderer grammar.
    source = f"![image]\n\n[image]: {destination}\n"
    ast = create_extended_markdown()(source)
    assert isinstance(ast, list)
    assert (ast[0]["children"][0]["type"] == "image") is accepted
    resolver = DeclaredResourceResolver("note.md", {"asset.png": "/declared.png"})
    binding = {
        "authored_sha256": hashlib.sha256(source.encode()).hexdigest(),
        "images": [{"authored_token": "![image]", "logical_path": "asset.png"}],
    }
    if accepted:
        assert resolver.with_bindings(source, binding).resolve_image("![image]", "unused") == "/declared.png"
    else:
        with pytest.raises(DeclaredResourceError, match="visible authored image"):
            resolver.with_bindings(source, binding)


@pytest.mark.parametrize(
    "definition,accepted",
    [
        *[(f"paragraph\n{char}\n[image]: active.png\n", False) for char in ["\u00a0", "\u2003", "\x85", "\u2028"]],
        ("[^note]:\n[image]: active.png\n", False),
        ("[^note]: text\n[image]: active.png\n", True),
        ("[^note]:\nbody\n[image]: active.png\n", True),
        ("[^note]: text\n\n   \vcontinuation\n[image]: active.png\n", True),
    ],
)
def test_reference_definition_block_eligibility_matches_the_frozen_renderer(definition: str, accepted: bool) -> None:
    source = "![image]\n\n" + definition
    ast = create_extended_markdown()(source)
    assert isinstance(ast, list)
    assert (ast[0]["children"][0]["type"] == "image") is accepted
    resolver = DeclaredResourceResolver("note.md", {"asset.png": "/declared.png"})
    binding = {
        "authored_sha256": hashlib.sha256(source.encode()).hexdigest(),
        "images": [{"authored_token": "![image]", "logical_path": "asset.png"}],
    }
    if accepted:
        assert resolver.with_bindings(source, binding).resolve_image("![image]", "unused") == "/declared.png"
    else:
        with pytest.raises(DeclaredResourceError, match="visible authored image"):
            resolver.with_bindings(source, binding)
