"""Explicit clipboard image occurrences through Runtime, Finalizer and Bundle.

Generated PNGs and typed snapshots are real inputs. This does not exercise an
installed Office/Qt/Obsidian clipboard provider or assume its display behavior.
"""

from __future__ import annotations

import hashlib
import html
import posixpath
import re
from dataclasses import replace
from datetime import UTC, datetime
from io import BytesIO
from pathlib import Path, PurePosixPath
from typing import Any
from urllib.parse import unquote

import pytest
from PIL import Image

from docwen_application.bundle_mapping import build_bundle_draft
from docwen_core.models.clipboard_document import (
    ClipboardDocument,
    ClipboardImageRef,
    ClipboardParagraph,
    ClipboardResource,
    ClipboardText,
    clipboard_document_to_bytes,
)
from docwen_core.models.document_node import ConversionIdentity
from docwen_core.models.request import OutputPolicy
from docwen_gui.clipboard_inputs import ClipboardInputStore
from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter
from tests.integration.test_structured_resource_group_runtime import _builder, _source_ref

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


def _resource(resource_id: str, color: tuple[int, int, int]) -> tuple[ClipboardResource, bytes]:
    with Image.new("RGBA", (3, 2), (*color, 255)) as image, BytesIO() as buffer:
        rgba_sha256 = hashlib.sha256(image.tobytes()).hexdigest()
        image.save(buffer, format="PNG")
        data = buffer.getvalue()
    return (
        ClipboardResource(
            resource_id,
            f"resources/{resource_id}.png",
            "image/png",
            len(data),
            hashlib.sha256(data).hexdigest(),
            3,
            2,
            rgba_sha256,
        ),
        data,
    )


@pytest.mark.parametrize("style", ["wiki_link", "wiki_embed"])
@pytest.mark.parametrize("mode", ["file", "embed"])
@pytest.mark.parametrize("rename", [False, True])
def test_clipboard_wiki_occurrences_keep_alias_identity_after_finalization(
    tmp_path: Path, round_trip_runtime: Any, style: str, mode: str, rename: bool
) -> None:
    red, red_bytes = _resource("red", (210, 20, 20))
    blue, blue_bytes = _resource("blue", (20, 20, 210))
    occurrences = (
        (red, "First | [red] & value"),
        (blue, "Second blue"),
        (red, ""),
        (red, "Last &amp; [literal]"),
    )
    model = ClipboardDocument(
        tuple(
            ClipboardParagraph(
                (
                    ClipboardText(f"Before occurrence {index} "),
                    ClipboardImageRef(resource.resource_id, alt),
                    ClipboardText(f" After occurrence {index}."),
                )
            )
            for index, (resource, alt) in enumerate(occurrences)
        ),
        (red, blue),
    )
    payload = clipboard_document_to_bytes(model)
    # Keep independent originals available even after the Store releases its
    # managed copies, including in a retained failing QA directory.
    (tmp_path / "original.dwclip").write_bytes(payload)
    (tmp_path / "original-red.png").write_bytes(red_bytes)
    (tmp_path / "original-blue.png").write_bytes(blue_bytes)
    output = tmp_path / "published"
    base_root = "Clipboard Identity_20261029_120000_fromClipboardDocument"
    occupied = output / base_root
    if rename:
        occupied.mkdir(parents=True)
        (occupied / "existing.txt").write_bytes(b"existing output must survive")

    store = ClipboardInputStore(tmp_path / "managed")
    try:
        snapshot = store.create_bundle(
            payload,
            display_name_template="Clipboard Identity {index}.dwclip",
            preview="Generated image occurrences",
            resources=(
                (red.resource_id, red.logical_path, red.media_type, red_bytes),
                (blue.resource_id, blue.logical_path, blue.media_type, blue_bytes),
            ),
        )
        store.sync_visible([snapshot.main.path])
        request, _ = _builder(store, [_source_ref(snapshot)]).single(
            file_path=snapshot.main.path,
            target_format="md",
            action_name="",
            options={},
            output_policy=OutputPolicy(output_dir=str(output)),
        )
        request = replace(
            request,
            options={"image_mode": mode, "image_link_style": style},
            conversion_identity=ConversionIdentity.create(
                task_id=request.request_id,
                source_stem="Clipboard Identity",
                source_format="clipboard_document",
                created_at=datetime(2026, 10, 29, 12, 0, tzinfo=UTC),
            ),
        )
        result = round_trip_runtime.execute(request)
        assert result.success, result.error
        assert Path(snapshot.main.path).read_bytes() == payload
        assert {Path(item.path).read_bytes() for item in snapshot.resources} == {red_bytes, blue_bytes}
        bundle = ArtifactBundleCommitter().commit(
            task_id=request.request_id,
            staging_root=str(output),
            draft=build_bundle_draft(
                profile="document_with_resources", output_media_type="text/markdown", artifacts=result.artifacts
            ),
        )
        primary_id = next(entry.artifact_id for entry in bundle.entries if entry.preferred)
        primary = next(item for item in bundle.artifacts if item.artifact_id == primary_id)
        root_name = "Clipboard Identity_20261029_120000_001_fromClipboardDocument" if rename else base_root
        assert primary.logical_path == f"{root_name}/{root_name}.md"
        if rename:
            assert (occupied / "existing.txt").read_bytes() == b"existing output must survive"
            assert [item.name for item in occupied.iterdir()] == ["existing.txt"]
        for artifact in bundle.artifacts:
            content = (output / artifact.locator).read_bytes()
            assert artifact.size_bytes == len(content)
            assert artifact.sha256 == hashlib.sha256(content).hexdigest()
        markdown = (output / primary.locator).read_text(encoding="utf-8-sig")
        # Independent Wiki grammar: target is left of the first separator;
        # only the output spelling is inspected, never DocWen's Wiki parser.
        tokens = list(re.finditer(r"(!?)\[\[([^\]\r\n]*)\]\]", markdown))
        assert len(tokens) == len(occurrences)
        actual_artifacts = []
        for index, (token, (original, alt)) in enumerate(zip(tokens, occurrences, strict=True)):
            target, separator, alias = token.group(2).partition("|")
            if style == "wiki_link":
                assert token.group(1) == ""
                assert separator == "|"
                assert html.unescape(alias) == alt
            else:
                assert token.group(1) == "!"
                assert separator == alias == ""
            logical_path = posixpath.normpath(
                posixpath.join(str(PurePosixPath(primary.logical_path).parent), unquote(target))
            )
            [artifact] = [item for item in bundle.artifacts if item.logical_path == logical_path]
            assert artifact.sha256 == original.sha256
            assert (output / artifact.locator).read_bytes() == (red_bytes if original is red else blue_bytes)
            before, after = f"Before occurrence {index} ", f" After occurrence {index}."
            assert markdown.count(before) == markdown.count(after) == 1
            assert markdown.index(before) < token.start() < token.end() <= markdown.index(after)
            actual_artifacts.append(artifact.artifact_id)
        assert actual_artifacts[0] == actual_artifacts[2] == actual_artifacts[3]
        assert actual_artifacts[0] != actual_artifacts[1]
        assert (tmp_path / "original.dwclip").read_bytes() == payload
        assert (tmp_path / "original-red.png").read_bytes() == red_bytes
        assert (tmp_path / "original-blue.png").read_bytes() == blue_bytes
    finally:
        store.close()


def test_markup_assigned_resource_names_share_the_wiki_alias_contract(tmp_path: Path, round_trip_runtime: Any) -> None:
    from tests.integration.test_resource_identity_pipeline import _bundle, _enex_note, _target_artifact

    _, first = _resource("first", (180, 20, 20))
    _, second = _resource("second", (20, 20, 180))
    source = tmp_path / "literal-percent.enex"
    source.write_text(
        "<en-export>"
        + _enex_note(
            "Literal filenames",
            [(first, "photo%20one.png", "image/png"), (second, "photo one.png", "image/png")],
        )
        + "</en-export>",
        encoding="utf-8",
    )
    bundle, markdown, primary = _bundle(
        round_trip_runtime, source, tmp_path / "published", image_link_style="wiki_link"
    )
    tokens = re.findall(r"(?<!!)\[\[([^\]\r\n]+)\]\]", markdown)
    assert len(tokens) == 2
    expected = [("photo-20one.png", first), ("photo-one.png", second)]
    for token, (name, original) in zip(tokens, expected, strict=True):
        target, separator, alias = token.partition("|")
        assert separator == "|"
        assert html.unescape(alias) == name
        artifact = _target_artifact(bundle, primary, target)
        assert artifact.sha256 == hashlib.sha256(original).hexdigest()
