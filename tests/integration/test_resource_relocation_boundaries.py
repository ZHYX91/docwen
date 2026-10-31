"""Independent resource identities through the real finalizer and Bundle."""

from __future__ import annotations

import hashlib
import html
import posixpath
import re
from pathlib import Path, PurePosixPath
from typing import Any
from urllib.parse import unquote

import pytest
from docx import Document

from docwen_application.bundle_mapping import build_bundle_draft
from docwen_core.models import ArtifactManifest, OutputPolicy
from docwen_core.text.image_markdown import generate_image_markdown
from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter
from docwen_runtime.output.finalizer import OutputFinalizer

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


def _publish(tmp_path: Path, markdown: bytes, resources: list[tuple[str, bytes]]) -> tuple[Any, Path, bytes]:
    staging = tmp_path / "staging"
    staging.mkdir()
    source = tmp_path / "source.docx"
    Document().save(str(source))
    main = staging / "main.md"
    main.write_bytes(markdown)
    artifacts = [
        ArtifactManifest(
            artifact_id="main",
            kind="primary",
            staging_path=str(main),
            suggested_name="main.md",
            media_type="text/markdown",
            is_primary=True,
        )
    ]
    for index, (name, content) in enumerate(resources):
        path = staging / f"resource-{index}.bin"
        path.write_bytes(content)
        artifacts.append(
            ArtifactManifest(
                artifact_id=f"resource-{index}",
                kind="auxiliary",
                staging_path=str(path),
                suggested_name=name,
                media_type="application/octet-stream",
                metadata={"source_kind": "embedded_resource"},
            )
        )
    output = tmp_path / "output"
    result = OutputFinalizer().finalize(
        "resource-boundary", artifacts, OutputPolicy(output_dir=str(output)), input_path=str(source)
    )
    if not result.success:
        assert main.read_bytes() == markdown
        assert [(staging / f"resource-{index}.bin").read_bytes() for index in range(len(resources))] == [
            content for _, content in resources
        ]
        return result, output, b""
    bundle = ArtifactBundleCommitter().commit(
        task_id="resource-boundary",
        staging_root=str(output),
        draft=build_bundle_draft(
            profile="document_with_resources", output_media_type="text/markdown", artifacts=result.artifacts
        ),
    )
    for item in bundle.artifacts:
        data = (output / item.locator).read_bytes()
        assert item.size_bytes == len(data)
        assert item.sha256 == hashlib.sha256(data).hexdigest()
    primary = next(item for item in bundle.artifacts if item.artifact_id == "main")
    return bundle, output, (output / primary.locator).read_bytes()


def _resolve(bundle: Any, target: str) -> Any:
    primary = next(item for item in bundle.artifacts if item.artifact_id == "main")
    path = unquote(target.partition("#")[0].partition("?")[0])
    logical = posixpath.normpath(posixpath.join(str(PurePosixPath(primary.logical_path).parent), path))
    [artifact] = [item for item in bundle.artifacts if item.logical_path == logical]
    return artifact


def test_qualified_relocation_uses_shared_source_ownership_and_preserves_exact_other_bytes(tmp_path: Path) -> None:
    source = (
        '\ufeff---\r\nexample: "![yaml](assets/image.png)"\r\n---\r\n\r\n'
        "[real](assets/image.png?download=1#author%29fragment)\r\n"
        "[space](assets/my%20space%25.png#kept)\r\n"
        "[other](../unrelated/image.png)\r\n"
        "[[../unrelated/image.png|untouched]]\r\n"
        "![[assets/image.png|kept alias]]\r\n"
        'Visible <img src="assets/image.png" title=\'src="assets/image.png"\'>\r\n'
        'Visible <img src="../unrelated/image.png">\r\n'
        "``literal ` ![sample](assets/image.png)``\r\n"
        '[label `![sample](assets/image.png)`](assets/image.png "![title](assets/image.png)")\r\n'
        "\r\n<!-- ![comment](assets/image.png) -->\r\n"
        '\r\n<div>\r\n<img src="assets/image.png">\r\n![html body literal](assets/image.png)\r\n</div>\r\n'
        '\r\n<pre>\r\n<img src="assets/image.png">\r\n</pre>\r\n'
        "\r\n- ```markdown\r\n  ![list sample](assets/image.png)\r\n"
        "\r\n![after unclosed list fence](assets/image.png)\r\n"
        "\r\n````markdown\r\n![top sample](assets/image.png)\r\n```\r\n````\r\n"
    )
    expected = (
        source.replace(
            "[real](assets/image.png?download=1#author%29fragment)", "[real](image.png?download=1#author%29fragment)"
        )
        .replace("[space](assets/my%20space%25.png#kept)", "[space](my%20space%25.png#kept)")
        .replace("![[assets/image.png|kept alias]]", "![[image.png|kept alias]]")
        .replace('Visible <img src="assets/image.png"', 'Visible <img src="image.png"')
        .replace('`](assets/image.png "', '`](image.png "')
        .replace('<div>\r\n<img src="assets/image.png">', '<div>\r\n<img src="image.png">')
        .replace("![after unclosed list fence](assets/image.png)", "![after unclosed list fence](image.png)")
    )
    bundle, _, result = _publish(
        tmp_path, source.encode(), [("assets/image.png", b"actual image"), ("assets/my space%.png", b"space percent")]
    )
    assert result == expected.encode()
    assert (
        _resolve(bundle, "image.png?download=1#author%29fragment").sha256 == hashlib.sha256(b"actual image").hexdigest()
    )
    assert _resolve(bundle, "my%20space%25.png#kept").sha256 == hashlib.sha256(b"space percent").hexdigest()


@pytest.mark.parametrize("syntax", ["markdown", "wiki", "html"])
def test_ambiguous_bare_alias_rejects_publication_but_qualified_aliases_are_exact(tmp_path: Path, syntax: str) -> None:
    token = {"markdown": "[x](photo.png)", "wiki": "[[photo.png|x]]", "html": 'Text <img src="photo.png">'}[syntax]
    result, output, _ = _publish(
        tmp_path, (token + "\n").encode(), [("left/photo.png", b"left"), ("right/photo.png", b"right")]
    )
    assert not result.success
    assert result.error is not None
    assert "ambiguous producer-declared artifact reference" in result.error.message
    assert not output.exists() or not list(output.iterdir())


def test_qualified_aliases_with_same_basename_keep_two_independent_resources(tmp_path: Path) -> None:
    bundle, _, content = _publish(
        tmp_path,
        b"[left](left/photo.png)\n[right](right/photo.png)\n",
        [("left/photo.png", b"left"), ("right/photo.png", b"right")],
    )
    targets = re.findall(rb"\[(?:left|right)\]\(([^)]+)\)", content)
    assert [_resolve(bundle, value.decode()).sha256 for value in targets] == [
        hashlib.sha256(b"left").hexdigest(),
        hashlib.sha256(b"right").hexdigest(),
    ]


@pytest.mark.parametrize("description", ["", "Portrait | [front] & rear"])
def test_wiki_image_identity_survives_finalizer_filename_changes(tmp_path: Path, description: str) -> None:
    name = "assets/photo [front]|100%.png"
    token = generate_image_markdown(image_path=name, image_link_style="wiki_link", alt_text=description)
    bundle, _, output = _publish(tmp_path, (token + "\n").encode(), [(name, b"original photo bytes")])
    [body] = re.findall(r"(?<!!)\[\[([^\]\r\n]+)\]\]", output.decode())
    target, separator, alias = body.partition("|")
    assert separator == "|"
    assert html.unescape(alias) == description
    artifact = _resolve(bundle, target)
    assert artifact.sha256 == hashlib.sha256(b"original photo bytes").hexdigest()
    assert artifact.logical_path.endswith("photo [front]_100%.png")
