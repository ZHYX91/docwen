"""Final Runtime results for MIME context, HTML base and authored image tokens."""

from __future__ import annotations

import base64
import hashlib
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pytest

from ._input_routes_support import _run_request, _test_png_bytes
from ._input_routes_support import pipeline as pipeline

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


def _image(location: str, data: bytes) -> MIMEImage:
    part = MIMEImage(data, _subtype="png")
    part["Content-Location"] = location
    return part


def _convert(pipeline, tmp_path: Path, source: Path, **options):
    output = tmp_path / "output"
    output.mkdir(exist_ok=True)
    original = source.read_bytes()
    result = _run_request(
        pipeline[1], source, source.suffix.lstrip("."), output, image_link_style="markdown_embed", **options
    )
    assert source.read_bytes() == original
    return result


def _markdown(result) -> str:
    assert result.success, result.error
    return Path(next(item.staging_path for item in result.artifacts if item.is_primary)).read_text(encoding="utf-8")


@pytest.mark.parametrize("absolute", [False, True], ids=["synthetic-base", "absolute-base"])
def test_resource_content_location_uses_its_own_mime_parent_not_html_uri(
    pipeline, tmp_path: Path, absolute: bool
) -> None:
    archive = MIMEMultipart("related", type="text/html")
    if absolute:
        archive["Content-Location"] = "https://archive.example/"
    body = MIMEText('<p>Expected blue</p><img src="assets/picture.png"/>', "html", "utf-8")
    body["Content-Location"] = "https://archive.example/docs/page.html" if absolute else "docs/page.html"
    archive.attach(body)
    archive.attach(_image("assets/picture.png", _test_png_bytes((240, 0, 0))))
    blue = _test_png_bytes((0, 0, 240))
    archive.attach(_image("docs/assets/picture.png", blue))
    source = tmp_path / "mime-parent.mhtml"
    source.write_bytes(archive.as_bytes())
    result = _convert(pipeline, tmp_path, source)
    _markdown(result)
    images = [item for item in result.artifacts if item.kind == "image"]
    assert len(images) == 1
    assert Path(images[0].staging_path).read_bytes() == blue


def test_related_start_selects_root_body_instead_of_first_html(pipeline, tmp_path: Path) -> None:
    archive = MIMEMultipart("related", type="text/html", start="<root@example.invalid>")
    for cid, text in [("frame@example.invalid", "Other frame"), ("root@example.invalid", "Expected root")]:
        body = MIMEText(f"<p>{text}</p>", "html", "utf-8")
        body["Content-ID"] = "<" + cid + ">"
        archive.attach(body)
    source = tmp_path / "selected-root.mhtml"
    source.write_bytes(archive.as_bytes())
    content = _markdown(_convert(pipeline, tmp_path, source))
    assert "Expected root" in content and "Other frame" not in content


@pytest.mark.parametrize("decoy", [False, True])
def test_relative_html_base_has_authority_over_sibling_and_companion(pipeline, tmp_path: Path, decoy: bool) -> None:
    (tmp_path / "assets").mkdir()
    red = _test_png_bytes((240, 0, 0))
    (tmp_path / "assets/picture.png").write_bytes(red)
    if decoy:
        (tmp_path / "picture.png").write_bytes(_test_png_bytes((0, 0, 240)))
    source = tmp_path / "relative-base.html"
    source.write_text(
        '<html><head><base href="assets/"/></head><body><img src="picture.png"/></body></html>', encoding="utf-8"
    )
    result = _convert(pipeline, tmp_path, source)
    _markdown(result)
    images = [item for item in result.artifacts if item.kind == "image"]
    assert len(images) == 1
    assert Path(images[0].staging_path).read_bytes() == red


@pytest.mark.parametrize("keep", [False, True])
def test_authored_body_and_code_image_token_remain_exact(pipeline, tmp_path: Path, keep: bool) -> None:
    source = tmp_path / "literal-token.html"
    source.write_text(
        '<p>DOCWENHTMLIMAGE1TOKEN</p><pre><code>DOCWENHTMLIMAGE1TOKEN</code></pre><img src="https://example.invalid/picture.png"/>',
        encoding="utf-8",
    )
    content = _markdown(_convert(pipeline, tmp_path, source, to_md_keep_images=keep))
    assert content.count("DOCWENHTMLIMAGE1TOKEN") == 2


@pytest.mark.parametrize("keep", [False, True])
def test_enex_literal_resource_token_is_preserved_in_final_runtime_text(pipeline, tmp_path: Path, keep: bool) -> None:
    image = _test_png_bytes()
    digest = hashlib.md5(image).hexdigest()
    literal = f"__DOCWEN_RES_{digest}__"
    enml = f'<en-note><p>{literal}</p><pre><code>{literal}</code></pre><en-media hash="{digest}" type="image/png"/></en-note>'
    source = tmp_path / "literal-resource.enex"
    source.write_text(
        f'<en-export><note><title>Literal resource</title><content><![CDATA[{enml}]]></content><resource><data encoding="base64">{base64.b64encode(image).decode()}</data><mime>image/png</mime></resource></note></en-export>',
        encoding="utf-8",
    )
    content = _markdown(_convert(pipeline, tmp_path, source, to_md_keep_images=keep))
    # markdownify escapes literal underscores in ordinary text, but leaves
    # code text exact. Compare the authored visible spelling in both contexts.
    assert content.replace("\\_", "_").count(literal) == 2


@pytest.mark.parametrize("ambiguous", [False, True])
def test_mhtml_basename_alias_is_used_only_when_unique(pipeline, tmp_path: Path, ambiguous: bool) -> None:
    archive = MIMEMultipart("related", type="text/html")
    archive.attach(MIMEText('<img src="picture.png"/>', "html", "utf-8"))
    red = _test_png_bytes((240, 0, 0))
    archive.attach(_image("https://cdn.example/assets/picture.png", red))
    if ambiguous:
        archive.attach(_image("https://other.example/picture.png", _test_png_bytes((0, 0, 240))))
    source = tmp_path / "alias.mhtml"
    source.write_bytes(archive.as_bytes())
    result = _convert(pipeline, tmp_path, source)
    if ambiguous:
        assert not result.success and not result.artifacts
        assert result.error is not None and len(result.error.message) < 200
    else:
        _markdown(result)
        image = next(item for item in result.artifacts if item.kind == "image")
        assert Path(image.staging_path).read_bytes() == red


def test_nested_mime_parent_context_selects_inner_image(pipeline, tmp_path: Path) -> None:
    archive = MIMEMultipart("related", type="multipart/related")
    archive["Content-Location"] = "https://archive.example/"
    nested = MIMEMultipart("related", type="text/html")
    nested["Content-Location"] = "frames/"
    body = MIMEText('<img src="assets/picture.png"/>', "html", "utf-8")
    body["Content-Location"] = "page.html"
    nested.attach(body)
    blue = _test_png_bytes((0, 0, 240))
    nested.attach(_image("assets/picture.png", blue))
    archive.attach(nested)
    archive.attach(_image("assets/picture.png", _test_png_bytes((240, 0, 0))))
    source = tmp_path / "nested.mhtml"
    source.write_bytes(archive.as_bytes())
    result = _convert(pipeline, tmp_path, source)
    _markdown(result)
    assert [Path(item.staging_path).read_bytes() for item in result.artifacts if item.kind == "image"] == [blue]


@pytest.mark.parametrize("invalid", ["missing", "duplicate", "unsupported", "first-plain", "type-mismatch"])
def test_invalid_declared_root_fails_without_selecting_other_html(pipeline, tmp_path: Path, invalid: str) -> None:
    archive = MIMEMultipart("related", type="text/html")
    body = MIMEText("<p>Unrelated HTML must not become the root</p>", "html", "utf-8")
    if invalid == "first-plain":
        archive.attach(MIMEText("Unsupported root", "plain"))
    else:
        archive.set_param("start", "<root@example.invalid>")
    body["Content-ID"] = (
        "<root@example.invalid>" if invalid in {"duplicate", "type-mismatch"} else "<frame@example.invalid>"
    )
    archive.attach(body)
    if invalid in {"duplicate", "unsupported"}:
        part = MIMEText("Selected unsupported root", "plain")
        part["Content-ID"] = "<root@example.invalid>"
        archive.attach(part)
    if invalid == "type-mismatch":
        archive.set_param("type", "multipart/alternative")
    source = tmp_path / "invalid-root.mhtml"
    source.write_bytes(archive.as_bytes())
    result = _convert(pipeline, tmp_path, source)
    assert not result.success and not result.artifacts
    assert result.error is not None and result.error.diagnostic_code == "HTML2MD-PARSE-ERROR"
    assert len(result.error.message) < 200 and "example.invalid" not in result.error.message


def test_related_alternative_root_selects_last_supported_html(pipeline, tmp_path: Path) -> None:
    archive = MIMEMultipart("related", type="multipart/alternative", start="<root@example.invalid>")
    archive.attach(MIMEText("<p>Unrelated frame</p>", "html", "utf-8"))
    alternative = MIMEMultipart("alternative")
    alternative["Content-ID"] = "<root@example.invalid>"
    alternative.attach(MIMEText("Fallback plain", "plain"))
    alternative.attach(MIMEText("<p>Less preferred HTML</p>", "html", "utf-8"))
    alternative.attach(MIMEText("<p>Preferred root HTML</p>", "html", "utf-8"))
    archive.attach(alternative)
    source = tmp_path / "alternative-root.mhtml"
    source.write_bytes(archive.as_bytes())
    content = _markdown(_convert(pipeline, tmp_path, source))
    assert "Preferred root HTML" in content
    assert "Unrelated frame" not in content and "Less preferred" not in content


def test_explicit_missing_base_never_uses_sibling_or_companion_decoys(pipeline, tmp_path: Path) -> None:
    source = tmp_path / "missing-base.html"
    decoy = _test_png_bytes((0, 0, 240))
    (tmp_path / "picture.png").write_bytes(decoy)
    companion = tmp_path / "missing-base_files"
    companion.mkdir()
    (companion / "picture.png").write_bytes(decoy)
    source.write_text(
        '<html><head><base href="absent/"/></head><body><img src="picture.png"/></body></html>', encoding="utf-8"
    )
    result = _convert(pipeline, tmp_path, source)
    _markdown(result)
    assert not [item for item in result.artifacts if item.kind == "image"]
    assert "HTML-LOCAL-IMAGE-UNAVAILABLE" in {item.code for item in result.diagnostics}


def test_missing_local_image_diagnostic_has_bounded_count_and_length(pipeline, tmp_path: Path) -> None:
    source = tmp_path / "missing-images.html"
    source.write_text(
        "<html><body>"
        + "".join(f'<img src="missing-{index}-' + "x" * 500 + '.png"/>' for index in range(20))
        + "</body></html>",
        encoding="utf-8",
    )
    result = _convert(pipeline, tmp_path, source)
    _markdown(result)
    diagnostics = [item for item in result.diagnostics if item.code == "HTML-LOCAL-IMAGE-UNAVAILABLE"]
    assert len(diagnostics) == 1
    assert "20 local image" in diagnostics[0].message and len(diagnostics[0].message) < 256
    assert "x" * 100 not in diagnostics[0].message


def test_omit_images_does_not_remove_authored_tokens(pipeline, tmp_path: Path) -> None:
    source = tmp_path / "omit-images.html"
    uri = "data:image/png;base64," + base64.b64encode(_test_png_bytes()).decode()
    source.write_text(
        f'<p>DOCWENHTMLIMAGE1TOKEN</p><pre><code>DOCWENHTMLIMAGE1TOKEN</code></pre><img src="{uri}"/>', encoding="utf-8"
    )
    result = _convert(pipeline, tmp_path, source, image_mode="omit")
    assert _markdown(result).count("DOCWENHTMLIMAGE1TOKEN") == 2
    assert not [item for item in result.artifacts if item.kind == "image"]
