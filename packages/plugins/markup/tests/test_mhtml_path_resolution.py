"""Relative MHTML resolution preserves empty path segments and rootless bases."""

from __future__ import annotations

from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pytest

from docwen_plugin_markup.web_archive.resource_identity import mhtml_resource_identity, resolve_mhtml_uri

from ._input_routes_support import _run_request, _test_png_bytes
from ._input_routes_support import pipeline as pipeline

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


@pytest.mark.parametrize("prefix", ["https://archive.example/", "thismessage:/"])
@pytest.mark.parametrize("decoy", [False, True], ids=["correct-only", "with-decoy"])
@pytest.mark.parametrize("path", ["images//p.png", "images///p.png", "images//../p.png"])
def test_empty_path_segments_select_exact_png(pipeline, tmp_path: Path, prefix: str, decoy: bool, path: str) -> None:
    correct, wrong = _test_png_bytes((240, 0, 0)), _test_png_bytes((0, 0, 240))
    # Removing '..' consumes only the immediately preceding empty segment.
    resolved_path = "images/p.png" if ".." in path else path
    archive = MIMEMultipart("related", type="text/html")
    archive.attach(MIMEText(f'<base href="{prefix}dir/page.html"><img src="{path}">', "html", "utf-8"))
    resources = [(prefix + "dir/" + resolved_path, correct)]
    if decoy:
        flattened = "p.png" if ".." in path else "images/p.png"
        resources.append((prefix + "dir/" + flattened, wrong))
    _assert_published_image(pipeline[1], tmp_path, archive, resources, correct)


@pytest.mark.parametrize("reference", ["#v", "#", "?red#v", "?#v"])
@pytest.mark.parametrize("decoy", [False, True], ids=["correct-only", "with-decoy"])
def test_rootless_synthetic_base_keeps_its_path(pipeline, tmp_path: Path, reference: str, decoy: bool) -> None:
    correct, wrong = _test_png_bytes((240, 0, 0)), _test_png_bytes((0, 0, 240))
    query = "?" if reference == "?#v" else "?red"
    archive = MIMEMultipart("related", type="text/html")
    archive.attach(MIMEText(f'<base href="thismessage:g?red"><img src="{reference}">', "html", "utf-8"))
    resources = [("thismessage:g" + query, correct)]
    if decoy:
        resources.append(("thismessage:" + query, wrong))
    _assert_published_image(pipeline[1], tmp_path, archive, resources, correct)


def _assert_published_image(
    task_manager, root: Path, archive, resources: list[tuple[str, bytes]], correct: bytes
) -> None:
    for location, data in resources:
        part = MIMEImage(data, _subtype="png")
        part["Content-Location"] = location
        archive.attach(part)
    source = root / "path.mhtml"
    original = archive.as_bytes()
    source.write_bytes(original)
    output = root / "output"
    output.mkdir()
    result = _run_request(task_manager, source, "mhtml", output, image_link_style="markdown_embed")
    assert result.success, result.error
    assert [Path(item.staging_path).read_bytes() for item in result.artifacts if item.kind == "image"] == [correct]
    assert source.read_bytes() == original


@pytest.mark.parametrize(
    ("base", "reference", "expected"),
    [
        ("thismessage:g?red", "other.png", "thismessage:other.png"),
        ("thismessage:g?red", "?", "thismessage:g?"),
        ("thismessage:g?red", "#v", "thismessage:g?red#v"),
        ("thismessage:g?red", "", "thismessage:g?red"),
        ("https://archive.example", "p.png", "https://archive.example/p.png"),
        ("thismessage:g", "g", "thismessage:g"),
        ("thismessage:/dir//page", "../p.png", "thismessage:/dir/p.png"),
        ("thismessage:/dir//page", "p.png", "thismessage:/dir//p.png"),
        ("https://archive.example/dir/page", "./a//b/../p.png", "https://archive.example/dir/a//p.png"),
        (
            "thismessage:/a/page",
            "images/%2e%2e/%2f/P.png?q=/../#v/./",
            "thismessage:/a/images/%2e%2e/%2f/P.png?q=/../#v/./",
        ),
        ("https://archive.example/a", "http:g//p?", "http:g//p?"),
        ("https://archive.example/a", "///g//p?", "https:///g//p?"),
    ],
)
def test_component_merging_does_not_rewrite_other_uri_parts(base: str, reference: str, expected: str) -> None:
    assert resolve_mhtml_uri(base, reference) == expected
    identity = expected.partition("#")[0]
    assert mhtml_resource_identity(reference, base_uri=base) == identity
    assert mhtml_resource_identity(identity) == identity


@pytest.mark.parametrize(
    ("reference", "expected"),
    [
        ("g:h", "g:h"),
        ("g", "http://a/b/c/g"),
        ("./g", "http://a/b/c/g"),
        ("g/", "http://a/b/c/g/"),
        ("/g", "http://a/g"),
        ("//g", "http://g"),
        ("?y", "http://a/b/c/d;p?y"),
        ("g?y", "http://a/b/c/g?y"),
        ("#s", "http://a/b/c/d;p?q#s"),
        ("g#s", "http://a/b/c/g#s"),
        ("g?y#s", "http://a/b/c/g?y#s"),
        (";x", "http://a/b/c/;x"),
        ("g;x", "http://a/b/c/g;x"),
        ("g;x?y#s", "http://a/b/c/g;x?y#s"),
        ("", "http://a/b/c/d;p?q"),
        (".", "http://a/b/c/"),
        ("./", "http://a/b/c/"),
        ("..", "http://a/b/"),
        ("../", "http://a/b/"),
        ("../g", "http://a/b/g"),
        ("../..", "http://a/"),
        ("../../", "http://a/"),
        ("../../g", "http://a/g"),
        ("../../../g", "http://a/g"),
        ("../../../../g", "http://a/g"),
        ("/./g", "http://a/g"),
        ("/../g", "http://a/g"),
        ("g.", "http://a/b/c/g."),
        (".g", "http://a/b/c/.g"),
        ("g..", "http://a/b/c/g.."),
        ("..g", "http://a/b/c/..g"),
        ("./../g", "http://a/b/g"),
        ("./g/.", "http://a/b/c/g/"),
        ("g/./h", "http://a/b/c/g/h"),
        ("g/../h", "http://a/b/c/h"),
        ("g;x=1/./y", "http://a/b/c/g;x=1/y"),
        ("g;x=1/../y", "http://a/b/c/y"),
        ("g?y/./x", "http://a/b/c/g?y/./x"),
        ("g?y/../x", "http://a/b/c/g?y/../x"),
        ("g#s/./x", "http://a/b/c/g#s/./x"),
        ("g#s/../x", "http://a/b/c/g#s/../x"),
        ("http:g", "http:g"),
    ],
)
def test_rfc3986_normal_and_abnormal_relative_vectors(reference: str, expected: str) -> None:
    """RFC 3986 section5.4 vectors, independent of the resolver implementation."""
    assert resolve_mhtml_uri("http://a/b/c/d;p?q", reference) == expected
