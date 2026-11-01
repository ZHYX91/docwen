"""Document-node link relocation keeps external and literal targets intact."""

from __future__ import annotations

import hashlib
from pathlib import Path

import pytest

from docwen_core.models import OutputPolicy
from docwen_runtime.output.finalizer import OutputFinalizer

from ._document_node_output_support import _artifact

pytestmark = pytest.mark.integration


def test_document_node_link_relocation_preserves_external_targets_and_integrity(tmp_path: Path) -> None:
    source = tmp_path / "report.docx"
    source.write_bytes(b"source")
    staging = tmp_path / "staging-links"
    staging.mkdir()
    main = staging / "main.md"
    child = staging / "child.md"
    image = staging / "image.png"
    hash_image = staging / "hash-image.png"
    colon_image = staging / "colon-image.png"
    spaced_image = staging / "space-image.png"
    main_original = (
        "[online](https://example.test/image.png#section)\n"
        '[angle-online](<https://example.test/image.png#section> "title")\n'
        '[angle-network](<//example.test/image.png#section> "title")\n'
        '[angle-local](<assets/my image.png#part> "local title")\n'
        '[angle-hash](<assets/image%23v1.png#part> "hash title")\n'
        "[invalid](https://[broken/image.png)\n"
        "[nfkc](https://／example.test/image.png)\n"
        "[different](https://example.test/other.png#other)\n"
        "[network](//example.test/image.png#network)\n"
        "[unc](\\\\server\\image.png#unc)\n"
        "[anchor](#section)\n"
        "[local](assets/image.png#local)\n"
        "[encoded-path](assets%2Fimage.png#encoded)\n"
        "[encoded-fragment](assets/image.png#part%29tail)\n"
        "![[assets/image.png#part%29wiki]]\n"
        "[[https://example.test/image.png#wiki-online]]\n"
        "[encoded-hash-name](assets/image%23v1.png#p)\n"
        "[encoded-colon-name](part%3Aone.png)\n"
        "[space-name](assets/my%20image.png#space)\n"
        '<a href="https://example.test/image.png#html-online">online</a>\n'
        '<a href="https://[broken/image.png">broken</a>\n'
        '<a href="https&colon;//example.test/image.png#part">entity</a>\n'
        '<a href="https&#58;//example.test/image.png#part">numeric</a>\n'
        '<a href="part%3Aone.png#part&amp;tail">local colon</a>\n'
        '<img src="assets/image.png#part%22tail">\n'
    )
    child_original = (
        "[local](assets/image.png#child%29tail)\n"
        '[angle-local](<assets/my image.png#part> "local title")\n'
        '[angle-online](<https://example.test/image.png#part> "title")\n'
        '<a href="https&colon;//example.test/image.png#part">entity</a>\n'
        '<img src="assets/image.png#child%22tail">\n'
        "![[assets/image.png#child%29wiki]]\n"
        "[hash](assets/image%23v1.png#child)\n"
    )
    main.write_bytes(main_original.encode("utf-8"))
    child.write_bytes(child_original.encode("utf-8"))
    image.write_bytes(b"png")
    hash_image.write_bytes(b"hash-png")
    colon_image.write_bytes(b"colon-png")
    spaced_image.write_bytes(b"space-png")
    output = tmp_path / "output-links"

    result = OutputFinalizer().finalize(
        "task.node.links",
        [
            _artifact(
                main,
                artifact_id="main",
                suggested_name="report.md",
                media_type="text/markdown",
                primary=True,
            ),
            _artifact(
                child,
                artifact_id="child",
                suggested_name="child.md",
                media_type="text/markdown",
            ),
            _artifact(
                image,
                artifact_id="image",
                suggested_name="assets/image.png",
                media_type="image/png",
            ),
            _artifact(
                hash_image,
                artifact_id="hash-image",
                suggested_name="assets/image#v1.png",
                media_type="image/png",
            ),
            _artifact(
                colon_image,
                artifact_id="colon-image",
                suggested_name="part:one.png",
                media_type="image/png",
            ),
            _artifact(
                spaced_image,
                artifact_id="space-image",
                suggested_name="assets/my image.png",
                media_type="image/png",
            ),
        ],
        OutputPolicy(output_dir=str(output), overwrite_mode="error"),
        input_path=str(source),
    )

    assert result.success is True, result.diagnostics
    expected_main = (
        "[online](https://example.test/image.png#section)\n"
        '[angle-online](<https://example.test/image.png#section> "title")\n'
        '[angle-network](<//example.test/image.png#section> "title")\n'
        '[angle-local](<my%20image.png#part> "local title")\n'
        '[angle-hash](<image%23v1.png#part> "hash title")\n'
        "[invalid](https://[broken/image.png)\n"
        "[nfkc](https://／example.test/image.png)\n"
        "[different](https://example.test/other.png#other)\n"
        "[network](//example.test/image.png#network)\n"
        "[unc](\\\\server\\image.png#unc)\n"
        "[anchor](#section)\n"
        "[local](image.png#local)\n"
        "[encoded-path](image.png#encoded)\n"
        "[encoded-fragment](image.png#part%29tail)\n"
        "![[image.png#part%29wiki]]\n"
        "[[https://example.test/image.png#wiki-online]]\n"
        "[encoded-hash-name](image%23v1.png#p)\n"
        "[encoded-colon-name](part_one.png)\n"
        "[space-name](my%20image.png#space)\n"
        '<a href="https://example.test/image.png#html-online">online</a>\n'
        '<a href="https://[broken/image.png">broken</a>\n'
        '<a href="https&colon;//example.test/image.png#part">entity</a>\n'
        '<a href="https&#58;//example.test/image.png#part">numeric</a>\n'
        '<a href="part_one.png#part&amp;tail">local colon</a>\n'
        '<img src="image.png#part%22tail">\n'
    ).encode()
    expected_child = (
        b"[local](../image.png#child%29tail)\n"
        b'[angle-local](<../my%20image.png#part> "local title")\n'
        b'[angle-online](<https://example.test/image.png#part> "title")\n'
        b'<a href="https&colon;//example.test/image.png#part">entity</a>\n'
        b'<img src="../image.png#child%22tail">\n'
        b"![[../image.png#child%29wiki]]\n"
        b"[hash](../image%23v1.png#child)\n"
    )
    artifacts = {artifact.artifact_id: artifact for artifact in result.artifacts}
    for artifact_id, expected in {"main": expected_main, "child": expected_child}.items():
        artifact = artifacts[artifact_id]
        published = Path(artifact.staging_path)
        assert published.read_bytes() == expected
        assert artifact.size_bytes == len(expected)
        assert artifact.sha256 == hashlib.sha256(expected).hexdigest()
    assert Path(artifacts["main"].staging_path).parent.parent == output
    assert Path(artifacts["child"].staging_path).parent.parent.parent == output


def test_document_node_link_relocation_keeps_code_samples_literal(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"source")
    staging = tmp_path / "staging-code-links"
    staging.mkdir()
    markdown = staging / "source.md"
    image = staging / "image.png"
    markdown.write_text(
        "![real](assets/image.png)\n"
        "`![inline](assets/image.png)`\n"
        "```markdown\n"
        "![fenced](assets/image.png)\n"
        "![[assets/image.png]]\n"
        '<img src="assets/image.png">\n'
        "```\n",
        encoding="utf-8",
    )
    image.write_bytes(b"png")
    output = tmp_path / "output-code-links"

    result = OutputFinalizer().finalize(
        "task.node.code-links",
        [
            _artifact(
                markdown,
                artifact_id="main",
                suggested_name="source.md",
                media_type="text/markdown",
                primary=True,
            ),
            _artifact(
                image,
                artifact_id="image",
                suggested_name="assets/image.png",
                media_type="image/png",
            ),
        ],
        OutputPolicy(output_dir=str(output), overwrite_mode="error"),
        input_path=str(source),
    )

    assert result.success is True, result.diagnostics
    primary = next(item for item in result.artifacts if item.artifact_id == "main")
    published = Path(primary.staging_path).read_text(encoding="utf-8")
    assert "![real](image.png)" in published
    assert "`![inline](assets/image.png)`" in published
    assert "![fenced](assets/image.png)" in published
    assert "![[assets/image.png]]" in published
    assert '<img src="assets/image.png">' in published
