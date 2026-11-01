"""Second review counterexamples through real MIME/ENEX/HTML Runtime outputs."""

from __future__ import annotations

import os
import xml.etree.ElementTree as ET
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from urllib.parse import urlsplit

import pytest

from ._input_routes_support import _run_request, _test_png_bytes
from ._input_routes_support import pipeline as pipeline

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


def _convert(pipeline, tmp_path: Path, source: Path, **options):
    output = tmp_path / "output"
    output.mkdir(exist_ok=True)
    original = source.read_bytes()
    result = _run_request(
        pipeline[1], source, source.suffix.lstrip("."), output, image_link_style="markdown_embed", **options
    )
    assert source.read_bytes() == original
    return result


def _content(result) -> str:
    assert result.success, result.error
    return Path(next(item.staging_path for item in result.artifacts if item.is_primary)).read_text(encoding="utf-8")


@pytest.mark.parametrize("dual", [False, True])
def test_content_id_header_is_literal_and_cid_url_decodes_once(pipeline, tmp_path: Path, dual: bool) -> None:
    archive = MIMEMultipart("related", type="text/html")
    sources = ["cid:part%252Fone@example.invalid"]
    if dual:
        sources.append("cid:part%2Fone@example.invalid")
    archive.attach(
        MIMEText(
            "<p>Exact CID identities</p>" + "".join(f'<img src="{source}"/>' for source in sources), "html", "utf-8"
        )
    )
    red = _test_png_bytes((240, 0, 0))
    blue = _test_png_bytes((0, 0, 240))
    for cid, payload in [("part%2Fone@example.invalid", red), *([("part/one@example.invalid", blue)] if dual else [])]:
        part = MIMEImage(payload, _subtype="png")
        part["Content-ID"] = "<" + cid + ">"
        archive.attach(part)
    source = tmp_path / "literal-cid.mhtml"
    source.write_bytes(archive.as_bytes())
    result = _convert(pipeline, tmp_path, source)
    _content(result)
    assert [Path(item.staging_path).read_bytes() for item in result.artifacts if item.kind == "image"] == (
        [red, blue] if dual else [red]
    )


def test_alternative_skips_valid_but_unsupported_nested_representation(pipeline, tmp_path: Path) -> None:
    archive = MIMEMultipart("related", type="multipart/alternative")
    alternative = MIMEMultipart("alternative")
    alternative.attach(MIMEText("<p>Supported earlier HTML</p>", "html", "utf-8"))
    later = MIMEMultipart("related", type="text/plain")
    later.attach(MIMEText("Unsupported plain representation", "plain", "utf-8"))
    alternative.attach(later)
    archive.attach(alternative)
    source = tmp_path / "fallback.mhtml"
    source.write_bytes(archive.as_bytes())
    content = _content(_convert(pipeline, tmp_path, source))
    assert "Supported earlier HTML" in content and "Unsupported plain" not in content


@pytest.mark.skipif(os.name != "nt", reason="Actual Windows file URI and drive-path conversion")
@pytest.mark.parametrize("resolved", [False, True], ids=["declared-drive", "resolved-drive"])
def test_windows_root_relative_uri_with_own_drive_delivers_real_png(pipeline, tmp_path: Path, resolved: bool) -> None:
    image = tmp_path / "actual image 中文.png"
    payload = _test_png_bytes()
    image.write_bytes(payload)
    source = tmp_path / "drive.html"
    reference = urlsplit((image.resolve() if resolved else image).as_uri()).path
    source.write_text(
        f'<html><head><base href="{source.as_uri()}"/></head><body><img src="{reference}"/></body></html>',
        encoding="utf-8",
    )
    result = _convert(pipeline, tmp_path, source)
    _content(result)
    assert [Path(item.staging_path).read_bytes() for item in result.artifacts if item.kind == "image"] == [payload]


def _write_enex(source: Path, enml: str) -> None:
    root = ET.Element("en-export")
    note = ET.SubElement(root, "note")
    ET.SubElement(note, "title").text = "Actual ENML nodes"
    ET.SubElement(note, "content").text = enml
    ET.ElementTree(root).write(source, encoding="utf-8", xml_declaration=True)


@pytest.mark.parametrize("mode", ["file", "omit"])
@pytest.mark.parametrize("pseudo", ["comment", "cdata"])
def test_pseudo_en_media_does_not_create_a_resource_dependency(
    pipeline, tmp_path: Path, mode: str, pseudo: str
) -> None:
    fake = '<en-media hash="00000000000000000000000000000000" type="image/png"/>'
    body = "<!-- " + fake + " -->" if pseudo == "comment" else "<pre><code><![CDATA[" + fake + "]]></code></pre>"
    source = tmp_path / "pseudo.enex"
    _write_enex(source, "<en-note>" + body + "<p>Visible author text</p></en-note>")
    result = _convert(pipeline, tmp_path, source, image_mode=mode)
    content = _content(result)
    assert "Visible author text" in content
    if pseudo == "cdata":
        assert fake in content
    assert not [item for item in result.artifacts if item.kind == "image"]


def test_omit_without_ocr_does_not_report_missing_images(pipeline, tmp_path: Path) -> None:
    source = tmp_path / "omit.html"
    source.write_text('<p>Preserved</p><img src="missing.png"/>', encoding="utf-8")
    result = _convert(pipeline, tmp_path, source, image_mode="omit", to_md_enable_ocr=False)
    assert "Preserved" in _content(result)
    assert not [item for item in result.diagnostics if item.code == "HTML-LOCAL-IMAGE-UNAVAILABLE"]


def test_wrong_cid_url_does_not_alias_a_percent_literal_header(pipeline, tmp_path: Path) -> None:
    archive = MIMEMultipart("related", type="text/html")
    archive.attach(MIMEText('<img src="cid:part%2Fone@example.invalid"/>', "html", "utf-8"))
    part = MIMEImage(_test_png_bytes(), _subtype="png")
    part["Content-ID"] = "<part%2Fone@example.invalid>"
    archive.attach(part)
    source = tmp_path / "different-cid.mhtml"
    source.write_bytes(archive.as_bytes())
    result = _convert(pipeline, tmp_path, source)
    _content(result)
    assert not [item for item in result.artifacts if item.kind == "image"]


@pytest.mark.parametrize("invalid", ["missing", "ambiguous", "type-mismatch", "empty"])
def test_invalid_later_alternative_does_not_hide_root_integrity_errors(pipeline, tmp_path: Path, invalid: str) -> None:
    alternative = MIMEMultipart("alternative")
    alternative.attach(MIMEText("<p>Earlier HTML must not hide malformed roots</p>", "html", "utf-8"))
    later = MIMEMultipart("related", type="text/plain", start="<chosen@example.invalid>")
    if invalid != "empty":
        body = MIMEText("Plain", "plain", "utf-8")
        body["Content-ID"] = "<other@example.invalid>" if invalid == "missing" else "<chosen@example.invalid>"
        later.attach(body)
    if invalid == "ambiguous":
        duplicate = MIMEText("Another plain", "plain", "utf-8")
        duplicate["Content-ID"] = "<chosen@example.invalid>"
        later.attach(duplicate)
    if invalid == "type-mismatch":
        later.set_param("type", "text/html")
    alternative.attach(later)
    archive = MIMEMultipart("related", type="multipart/alternative")
    archive.attach(alternative)
    source = tmp_path / "invalid-later.mhtml"
    source.write_bytes(archive.as_bytes())
    result = _convert(pipeline, tmp_path, source)
    assert not result.success and not result.artifacts
    assert result.error is not None and len(result.error.message) < 200


def test_attribute_pseudo_media_and_named_entities_are_author_text(pipeline, tmp_path: Path) -> None:
    source = tmp_path / "attributes.enex"
    _write_enex(
        source,
        '<en-note><div title="&lt;en-media hash=&quot;00000000000000000000000000000000&quot;/&gt;">A&nbsp;B &amp; C</div><pre><code><![CDATA[&nbsp; <en-media hash="fake"/>]]></code></pre></en-note>',
    )
    result = _convert(pipeline, tmp_path, source)
    content = _content(result)
    assert "A\xa0B & C" in content
    assert '&nbsp; <en-media hash="fake"/>' in content
    assert not [item for item in result.artifacts if item.kind == "image"]


@pytest.mark.parametrize("mode", ["file", "omit"])
def test_actual_missing_enml_resource_still_fails(pipeline, tmp_path: Path, mode: str) -> None:
    source = tmp_path / "actual-missing.enex"
    _write_enex(
        source, '<en-note><p>Visible</p><en-media hash="00000000000000000000000000000000" type="image/png"/></en-note>'
    )
    result = _convert(pipeline, tmp_path, source, image_mode=mode)
    assert not result.success and not result.artifacts
    assert result.error is not None and "missing or corrupt resources" in result.error.message


@pytest.mark.parametrize("kind", ["file", "network"])
def test_enml_external_entities_are_rejected_without_expansion(pipeline, tmp_path: Path, kind: str) -> None:
    secret = tmp_path / "entity.txt"
    secret.write_text("ENTITY FILE MUST NOT APPEAR", encoding="utf-8")
    target = secret.as_uri() if kind == "file" else "http://127.0.0.1:1/not-requested"
    source = tmp_path / "external.enex"
    _write_enex(
        source, f'<!DOCTYPE en-note [<!ENTITY external SYSTEM "{target}">]><en-note><p>&external;</p></en-note>'
    )
    result = _convert(pipeline, tmp_path, source)
    assert not result.success and not result.artifacts
    assert result.error is not None and len(result.error.message) < 200
    assert "ENTITY FILE MUST NOT APPEAR" not in result.error.message
    assert secret.read_text(encoding="utf-8") == "ENTITY FILE MUST NOT APPEAR"


def test_omit_without_ocr_never_reads_local_image_bytes(pipeline, tmp_path: Path, monkeypatch) -> None:
    image = tmp_path / "skip.png"
    image.write_bytes(_test_png_bytes())
    original = Path.read_bytes
    reads = []

    def read(path: Path) -> bytes:
        if path.resolve() == image.resolve():
            reads.append(path)
        return original(path)

    source = tmp_path / "skip.html"
    source.write_text('<p>Visible</p><img src="skip.png"/>', encoding="utf-8")
    monkeypatch.setattr(Path, "read_bytes", read)
    result = _convert(pipeline, tmp_path, source, image_mode="omit", to_md_enable_ocr=False)
    assert "Visible" in _content(result)
    assert not reads


def test_omit_with_ocr_retains_needed_data_and_truthful_missing_notice(pipeline, tmp_path: Path, monkeypatch) -> None:
    from docwen_core.text.ocr import OcrOutcome, OcrStatus

    image = tmp_path / "ocr.png"
    payload = _test_png_bytes()
    image.write_bytes(payload)
    observed = []

    def ocr(path: str, **_options):
        observed.append(Path(path).read_bytes())
        return OcrOutcome(OcrStatus.SUCCESS, text="Controlled OCR text")

    monkeypatch.setattr("docwen_plugin_markup.markdown_resources.run_ocr_outcome", ocr)
    source = tmp_path / "ocr.html"
    source.write_text('<p>Visible</p><img src="ocr.png"/><img src="missing.png"/>', encoding="utf-8")
    result = _convert(
        pipeline,
        tmp_path,
        source,
        image_mode="omit",
        to_md_enable_ocr=True,
        recognize_tables=False,
        ocr_placement="main_md",
    )
    assert "Controlled OCR text" in _content(result)
    assert observed == [payload]
    warnings = [item for item in result.diagnostics if item.code == "HTML-LOCAL-IMAGE-UNAVAILABLE"]
    assert len(warnings) == 1 and "OCR could not run" in warnings[0].message
    assert "links were retained" not in warnings[0].message
    assert not [item for item in result.artifacts if item.kind == "image"]
