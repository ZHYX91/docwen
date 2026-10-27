"""Relative MHTML references preserve the same query identity as absolute bases."""

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


@pytest.mark.parametrize("absolute", [False, True], ids=["synthetic", "absolute"])
@pytest.mark.parametrize("decoy", [False, True], ids=["only-correct", "with-decoy"])
@pytest.mark.parametrize(
    ("reference", "location"),
    [
        ("#v", "picture.png?rev=red"),
        ("#", "picture.png?rev=red"),
        ("?rev=blue#v", "picture.png?rev=blue"),
        ("?#v", "picture.png?"),
        ("?", "picture.png?"),
        ("other.png#v", "other.png"),
        ("./picture.png?rev=green", "picture.png?rev=green"),
        ("../other.png#v", "other.png"),
    ],
)
def test_fragment_reference_inherits_mhtml_base_query(
    pipeline, tmp_path: Path, absolute: bool, decoy: bool, reference: str, location: str
) -> None:
    correct = _test_png_bytes((240, 0, 0))
    wrong = _test_png_bytes((0, 0, 240))
    archive = MIMEMultipart("related", type="text/html")
    if absolute:
        archive["Content-Base"] = "https://archive.example/"
    archive.attach(
        MIMEText(f'<base href="picture.png?rev=red"><p>Expected red</p><img src="{reference}">', "html", "utf-8")
    )
    resources = [(location, correct)]
    if decoy:
        resources.append((location.partition("?")[0] if "?" in location else location + "?decoy", wrong))
    for location, data in resources:
        part = MIMEImage(data, _subtype="png")
        part["Content-Location"] = location
        archive.attach(part)
    source = tmp_path / "query.mhtml"
    source.write_bytes(archive.as_bytes())
    original = source.read_bytes()
    output = tmp_path / "output"
    output.mkdir()
    result = _run_request(pipeline[1], source, "mhtml", output, image_link_style="markdown_embed")
    assert result.success, result.error
    images = [item for item in result.artifacts if item.kind == "image"]
    assert [Path(item.staging_path).read_bytes() for item in images] == [correct]
    assert "Expected red" in Path(next(item.staging_path for item in result.artifacts if item.is_primary)).read_text(
        encoding="utf-8"
    )
    assert source.read_bytes() == original


@pytest.mark.parametrize("prefix", ["thismessage:/", "https://archive.example/"])
@pytest.mark.parametrize(
    ("reference", "suffix"),
    [
        ("", "dir/picture.png?rev=red"),
        ("#v", "dir/picture.png?rev=red#v"),
        ("#", "dir/picture.png?rev=red#"),
        ("?", "dir/picture.png?"),
        ("?#v", "dir/picture.png?#v"),
        ("?rev=blue", "dir/picture.png?rev=blue"),
        ("other.png", "dir/other.png"),
        ("../other.png#v", "other.png#v"),
    ],
)
def test_empty_query_and_fragment_presence_are_not_lost(prefix: str, reference: str, suffix: str) -> None:
    expected = prefix + suffix
    resolved = resolve_mhtml_uri(prefix + "dir/picture.png?rev=red#old", reference)
    assert resolved == expected
    assert mhtml_resource_identity(resolved) == expected.partition("#")[0]


@pytest.mark.parametrize("authority", ["docwen-mhtml.invalid", "archive.example"])
@pytest.mark.parametrize("own_base", [False, True], ids=["network-reference", "base-authority"])
def test_real_authority_never_aliases_the_internal_synthetic_carrier(
    pipeline, tmp_path: Path, authority: str, own_base: bool
) -> None:
    correct, wrong = _test_png_bytes((240, 0, 0)), _test_png_bytes((0, 0, 240))
    archive = MIMEMultipart("related", type="text/html")
    if own_base:
        archive["Content-Base"] = f"THISMESSAGE://{authority}/dir/"
    base = "../" if own_base else f"//{authority}/"
    archive.attach(MIMEText(f'<base href="{base}"><img src="p.png">', "html", "utf-8"))
    for location, data in [(f"THISMESSAGE://{authority}/p.png", correct), ("thismessage:/p.png", wrong)]:
        part = MIMEImage(data, _subtype="png")
        part["Content-Location"] = location
        archive.attach(part)
    source = tmp_path / "authority.mhtml"
    source.write_bytes(archive.as_bytes())
    output = tmp_path / "output"
    output.mkdir()
    result = _run_request(pipeline[1], source, "mhtml", output)
    assert result.success, result.error
    assert [Path(item.staging_path).read_bytes() for item in result.artifacts if item.kind == "image"] == [correct]


@pytest.mark.parametrize(
    ("base", "reference", "expected"),
    [
        ("THISMESSAGE:/dir/page?red", "#v", "thismessage:/dir/page?red#v"),
        ("THISMESSAGE://docwen-mhtml.invalid/dir/page?red", "#v", "thismessage://docwen-mhtml.invalid/dir/page?red#v"),
        ("thismessage://docwen-mhtml.invalid/dir/page", "../p.png", "thismessage://docwen-mhtml.invalid/p.png"),
        ("thismessage:/page", "//docwen-mhtml.invalid/p.png?", "thismessage://docwen-mhtml.invalid/p.png?"),
    ],
)
def test_synthetic_scheme_case_and_explicit_authorities_remain_distinct(
    base: str, reference: str, expected: str
) -> None:
    assert resolve_mhtml_uri(base, reference) == expected
