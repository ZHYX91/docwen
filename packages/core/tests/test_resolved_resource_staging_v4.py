"""Range-local, path-independent v4 resource staging gates."""

from __future__ import annotations

import hashlib
from pathlib import Path

import pytest

from docwen_core.models.resolved_numbering import (
    ResolvedDocument,
    ResolvedEmbeddedResource,
    ResolvedResourceOccurrence,
)
from docwen_core.resolved_resource_staging import (
    ResolvedResourceStagingError,
    ResolvedTextEdit,
    ResolvedTextProjection,
    bind_resolved_document_resources,
)

pytestmark = pytest.mark.unit

_PNG = b"\x89PNG\r\n\x1a\nrequest-owned-test-bytes"


def _resource(resource_id: str, *, content: bytes = _PNG) -> ResolvedEmbeddedResource:
    return ResolvedEmbeddedResource(
        resource_id=resource_id,
        role="linked_resource",
        media_type="image/png",
        size_bytes=len(content),
        sha256=hashlib.sha256(content).hexdigest(),
        content=content,
    )


def _occurrence(source: str, token: str, resource_id: str, *, after: int = 0) -> ResolvedResourceOccurrence:
    start = source.index(token, after)
    return ResolvedResourceOccurrence(
        source_start=start,
        source_end=start + len(token),
        source_slice_sha256=hashlib.sha256(token.encode()).hexdigest(),
        authored_token=token,
        authored_locator="../../outside.png",
        resource_id=resource_id,
    )


def _document() -> ResolvedDocument:
    first = "![same.png](../../outside.png)"
    second = "![same.png](../../outside.png)"
    third = "![[../../outside.png|200x100]]"
    source = f"{first}\n{second}\n{third}\n"
    first_occurrence = _occurrence(source, first, "res-a")
    second_occurrence = _occurrence(
        source,
        second,
        "res-b",
        after=first_occurrence.source_end,
    )
    third_occurrence = _occurrence(
        source,
        third,
        "res-a",
        after=second_occurrence.source_end,
    )
    return ResolvedDocument(
        source,
        (),
        (),
        (first_occurrence, second_occurrence, third_occurrence),
        (),
        (_resource("res-a"), _resource("res-b")),
    )


def test_same_locator_binds_by_occurrence_without_source_path_lookup(tmp_path: Path) -> None:
    root = tmp_path / "private-resources"

    binding = bind_resolved_document_resources(_document(), root)

    assert binding.path_for("res-a").read_bytes() == _PNG
    assert binding.path_for("res-b").read_bytes() == _PNG
    lines = binding.rendered_markdown.splitlines()
    assert lines[0] == f"![same.png](<{(root / 'res-a.png').as_posix()}>)"
    assert lines[1] == f"![same.png](<{(root / 'res-b.png').as_posix()}>)"
    assert lines[2] == f"![](<{(root / 'res-a.png').as_posix()}> =200x100)"
    assert "../../outside.png" not in binding.rendered_markdown
    assert sorted(path.name for path in root.iterdir()) == ["res-a.png", "res-b.png"]
    projection = binding.text_projection
    assert projection.source_length == len(_document().authored_markdown)
    assert projection.result_length == len(binding.rendered_markdown)
    assert tuple(binding.rendered_markdown[item.result_start : item.result_end] for item in projection.edits) == tuple(
        item.replacement for item in projection.edits
    )
    assert projection.project_range(projection.source_length - 1, projection.source_length) == (
        projection.result_length - 1,
        projection.result_length,
    )
    first = projection.edits[0]
    with pytest.raises(ResolvedResourceStagingError, match="overlaps"):
        projection.project_range(first.source_start, first.source_end)


def test_unicode_range_binds_a_cross_folder_short_link_with_spaces(tmp_path: Path) -> None:
    token = "![[image with spaces.png]]"
    source = f"😀 {token}\n"
    start = source.index(token)
    occurrence = ResolvedResourceOccurrence(
        source_start=start,
        source_end=start + len(token),
        source_slice_sha256=hashlib.sha256(token.encode()).hexdigest(),
        authored_token=token,
        authored_locator="image with spaces.png",
        resource_id="image-1",
    )
    document = ResolvedDocument(source, (), (), (occurrence,), (), (_resource("image-1"),))
    root = tmp_path / "resolved-resources"

    binding = bind_resolved_document_resources(document, root)

    assert binding.rendered_markdown == f"😀 ![](<{(root / 'image-1.png').as_posix()}>)\n"
    assert binding.path_for("image-1").read_bytes() == _PNG


def test_staging_revalidates_bytes_before_publishing_directory(tmp_path: Path) -> None:
    document = _document()
    resource = document.resources[0]
    tampered = ResolvedEmbeddedResource(
        resource.resource_id,
        resource.role,
        resource.media_type,
        resource.size_bytes,
        resource.sha256,
        resource.content + b"tamper",
    )
    changed = ResolvedDocument(
        document.authored_markdown,
        document.targets,
        document.references,
        document.resource_occurrences,
        document.citations,
        (tampered, document.resources[1]),
    )
    root = tmp_path / "must-not-exist"

    with pytest.raises(ResolvedResourceStagingError, match="changed after admission"):
        bind_resolved_document_resources(changed, root)
    assert not root.exists()
    assert not list(tmp_path.glob(".must-not-exist.*"))


def test_text_projection_rejects_noncanonical_edit_coordinates() -> None:
    edit = ResolvedTextEdit(2, 4, "bound", 2, 7)

    with pytest.raises(ResolvedResourceStagingError, match="result length"):
        ResolvedTextProjection(6, 99, (edit,))
    with pytest.raises(ResolvedResourceStagingError, match="exact and ordered"):
        ResolvedTextProjection(6, 9, (ResolvedTextEdit(2, 4, "bound", 3, 8),))


def test_existing_directory_is_never_reused_or_overwritten(tmp_path: Path) -> None:
    root = tmp_path / "already-owned"
    root.mkdir()
    sentinel = root / "sentinel.txt"
    sentinel.write_text("preserve", encoding="utf-8")

    with pytest.raises(ResolvedResourceStagingError, match="not fresh"):
        bind_resolved_document_resources(_document(), root)
    assert sentinel.read_text(encoding="utf-8") == "preserve"
    assert list(root.iterdir()) == [sentinel]


def test_embedded_bibliography_stays_in_memory(tmp_path: Path) -> None:
    payload = b'{"schema":"docwen.semantic_bibliography.v1","entries":[]}'
    bibliography = ResolvedEmbeddedResource(
        resource_id="bibliography",
        role="bibliography",
        media_type="application/vnd.docwen.semantic-bibliography+json",
        size_bytes=len(payload),
        sha256=hashlib.sha256(payload).hexdigest(),
        content=payload,
    )
    document = ResolvedDocument("body\n", (), (), (), (), (bibliography,))
    root = tmp_path / "no-resource-directory"

    binding = bind_resolved_document_resources(document, root)

    assert binding.rendered_markdown == "body\n"
    assert binding.bibliography is not None
    assert binding.bibliography.entries == ()
    assert binding.linked_paths == ()
    assert not root.exists()
