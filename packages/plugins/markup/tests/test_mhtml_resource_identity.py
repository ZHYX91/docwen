"""Real MIME resources retain their identities through publication."""

from __future__ import annotations

import base64
import re
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pytest

from ._input_routes_support import _run_request, _test_png_bytes
from ._input_routes_support import pipeline as pipeline

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


def _archive(
    path: Path, html: str, resources: list[tuple[str, bytes]], *, root: str = "https://archive.example/docs/page.html"
) -> None:
    archive = MIMEMultipart("related")
    body = MIMEText(html, "html", "utf-8")
    body["Content-Location"] = root
    archive.attach(body)
    for location, payload in resources:
        image = MIMEImage(payload, _subtype="png")
        if location.startswith("cid:"):
            image["Content-ID"] = "<" + location[4:] + ">"
        else:
            image["Content-Location"] = location
        archive.attach(image)
    path.write_bytes(archive.as_bytes())


@pytest.mark.parametrize("mode", ["file", "base64"])
@pytest.mark.parametrize(
    "locations",
    [
        ("https://one.example/logo.png", "https://two.example/logo.png"),
        ("https://one.example/logo.png?v=1", "https://one.example/logo.png?v=2"),
        ("https://one.example/Logo.png", "https://one.example/logo.png"),
        ("cid:Image-A", "cid:image-a"),
    ],
    ids=["host", "query", "path-case", "cid-case"],
)
def test_mhtml_distinct_images_never_share_payload(
    pipeline, tmp_path: Path, mode: str, locations: tuple[str, str]
) -> None:
    first, second = _test_png_bytes((240, 0, 0)), _test_png_bytes((0, 0, 240))
    source = tmp_path / "distinct.mhtml"
    _archive(
        source,
        f'<p><img src="{locations[0]}"/><img src="{locations[1]}"/></p>',
        [(locations[0], first), (locations[1], second)],
    )
    original = source.read_bytes()
    output = tmp_path / "output"
    output.mkdir()
    result = _run_request(pipeline[1], source, "mhtml", output, image_mode=mode, image_link_style="markdown_embed")
    assert result.success, result.error
    assert source.read_bytes() == original
    content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    if mode == "base64":
        payloads = [
            base64.b64decode(value) for value in re.findall(r"data:image/png;base64,([A-Za-z0-9+/=]+)", content)
        ]
        assert payloads == [first, second]
    else:
        images = [item for item in result.artifacts if item.kind == "image"]
        assert [Path(item.staging_path).read_bytes() for item in images] == [first, second]
        assert images[0].suggested_name != images[1].suggested_name
        assert all(item.suggested_name in content for item in images)


def test_mhtml_binds_relative_content_location_without_rewriting_body_or_code(pipeline, tmp_path: Path) -> None:
    uri = "https://archive.example/docs/picture.png"
    source = tmp_path / "relative.mhtml"
    _archive(
        source,
        f'<p>{uri}</p><pre><code>{uri}?literal=1</code></pre><img src="picture.png"/>',
        [("picture.png", _test_png_bytes())],
    )
    output = tmp_path / "output"
    output.mkdir()
    result = _run_request(pipeline[1], source, "mhtml", output, image_link_style="markdown_embed")
    assert result.success, result.error
    content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert uri in content and uri + "?literal=1" in content
    assert len([item for item in result.artifacts if item.kind == "image"]) == 1


def test_mhtml_img_binding_does_not_replace_literal_exact_or_prefix_urls(pipeline, tmp_path: Path) -> None:
    uri = "https://archive.example/picture.png"
    source = tmp_path / "literal.mhtml"
    _archive(source, f'<p>{uri}</p><pre><code>{uri}?extra=1</code></pre><img src="{uri}"/>', [(uri, _test_png_bytes())])
    output = tmp_path / "output"
    output.mkdir()
    result = _run_request(pipeline[1], source, "mhtml", output)
    assert result.success, result.error
    content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert uri in content and uri + "?extra=1" in content


def test_html_local_file_uri_copies_real_image_and_keeps_source_bytes(pipeline, tmp_path: Path) -> None:
    image = tmp_path / "local image 中文.png"
    payload = _test_png_bytes()
    image.write_bytes(payload)
    source = tmp_path / "local.html"
    source.write_text(f'<html><body><img src="{image.as_uri()}"/></body></html>', encoding="utf-8")
    output = tmp_path / "output"
    output.mkdir()
    result = _run_request(pipeline[1], source, "html", output)
    assert result.success, result.error
    images = [item for item in result.artifacts if item.kind == "image"]
    assert len(images) == 1
    assert Path(images[0].staging_path).read_bytes() == payload
    assert image.read_bytes() == payload


def test_file_base_uri_resolves_relative_images(pipeline, tmp_path: Path) -> None:
    resources = tmp_path / "resources 中文"
    resources.mkdir()
    image = resources / "local image.png"
    image.write_bytes(_test_png_bytes())
    source = tmp_path / "base.html"
    source.write_text(
        f'<html><head><base href="{resources.as_uri()}/"/></head><body><img src="local%20image.png"/></body></html>',
        encoding="utf-8",
    )
    output = tmp_path / "output"
    output.mkdir()
    result = _run_request(pipeline[1], source, "html", output)
    assert result.success, result.error
    assert len([item for item in result.artifacts if item.kind == "image"]) == 1


def test_mhtml_base_href_resolves_to_the_exact_image(pipeline, tmp_path: Path) -> None:
    source = tmp_path / "base.mhtml"
    _archive(
        source,
        '<html><head><base href="../assets/"/></head><body><img src="picture.png"/></body></html>',
        [("../assets/picture.png", _test_png_bytes())],
    )
    output = tmp_path / "output"
    output.mkdir()
    result = _run_request(pipeline[1], source, "mhtml", output)
    assert result.success, result.error
    assert len([item for item in result.artifacts if item.kind == "image"]) == 1


@pytest.mark.parametrize("referenced", [False, True])
def test_mhtml_conflicting_identity_fails_only_if_referenced(pipeline, tmp_path: Path, referenced: bool) -> None:
    source = tmp_path / "conflicting.mhtml"
    _archive(
        source,
        '<img src="duplicate.png"/>' if referenced else "<p>Ordinary body</p>",
        [("duplicate.png", _test_png_bytes((200, 0, 0))), ("duplicate.png", _test_png_bytes((0, 0, 200)))],
    )
    output = tmp_path / "output"
    output.mkdir()
    result = _run_request(pipeline[1], source, "mhtml", output)
    assert result.success is not referenced
    if referenced:
        assert not result.artifacts
