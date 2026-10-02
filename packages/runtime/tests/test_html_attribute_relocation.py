"""HTML attribute entities retain their meaning through real publication."""

from pathlib import Path

import pytest

from docwen_core.models import ArtifactManifest, OutputPolicy
from docwen_runtime.output.finalizer import OutputFinalizer

pytestmark = pytest.mark.integration


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("assets/a&copycat.png#part&copycat", "a%26copycat.png#part&amp;copycat"),
        ("assets/a&copy=cat.png#part&copy=cat", "a%26copy=cat.png#part&amp;copy=cat"),
        ("assets/a&copy;cat.png#part&copy;cat", "a©cat.png#part©cat"),
        ("assets/a&#169;cat.png#part&#xA9;cat", "a©cat.png#part©cat"),
        ("https&colon;//example.test/a&copycat.png", "https&colon;//example.test/a&copycat.png"),
    ],
)
def test_attribute_entities_select_the_correct_real_asset(tmp_path: Path, raw: str, expected: str) -> None:
    main = tmp_path / "main.md"
    original = f'<a href="{raw}">link</a>\n'.encode()
    main.write_bytes(original)
    artifacts = [
        ArtifactManifest(
            artifact_id="main",
            kind="primary",
            staging_path=str(main),
            suggested_name="report.md",
            media_type="text/markdown",
            is_primary=True,
        )
    ]
    for index, name in enumerate(("a&copycat.png", "a&copy=cat.png", "a©cat.png")):
        image = tmp_path / name
        image.write_bytes(f"distinct-image-{index}".encode())
        artifacts.append(
            ArtifactManifest(
                artifact_id=f"image-{index}",
                kind="auxiliary",
                staging_path=str(image),
                suggested_name=f"assets/{name}",
                media_type="image/png",
            )
        )
    result = OutputFinalizer().finalize(
        "test.attribute.entities",
        artifacts,
        OutputPolicy(output_dir=str(tmp_path / "output"), overwrite_mode="error"),
        input_path=str(tmp_path / "source.docx"),
    )
    assert result.success, result.diagnostics
    published = next(item for item in result.artifacts if item.is_primary)
    assert Path(published.staging_path).read_text(encoding="utf-8") == f'<a href="{expected}">link</a>\n'
    assert main.read_bytes() == original
    for index in range(3):
        item = next(item for item in result.artifacts if item.artifact_id == f"image-{index}")
        assert Path(item.staging_path).read_bytes() == f"distinct-image-{index}".encode()
