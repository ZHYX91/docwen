from __future__ import annotations

import base64
import re
from io import BytesIO
from types import SimpleNamespace

import pytest
from PIL import Image

pytestmark = pytest.mark.contract

from docwen_core.text.ocr import OcrOutcome, OcrStatus
from docwen_plugin_markup.markdown_resources import (
    MarkdownResource,
    MarkdownResourceWriter,
    normalize_resource_key,
)


class _Workspace:
    def __init__(self, tmp_path):
        self.tmp_path = tmp_path
        self.artifacts = []
        self.counter = 0

    def create_artifact_path(self, kind, suffix):
        self.counter += 1
        return str(self.tmp_path / f"artifact_{self.counter}{suffix}")

    def add_artifact(self, artifact):
        self.artifacts.append(artifact)


class _Progress:
    def __init__(self):
        self.diagnostics = []

    def report_diagnostic(self, level, message, code="", location=""):
        self.diagnostics.append(SimpleNamespace(level=level, message=message, code=code, location=location))


def _context(tmp_path, *, config_snapshot=None, ocr_blockquote_title=""):
    return SimpleNamespace(
        workspace=_Workspace(tmp_path),
        progress=_Progress(),
        request=SimpleNamespace(config_snapshot=config_snapshot or {}),
        ocr_blockquote_title=ocr_blockquote_title,
    )


def _successful_ocr(text: str) -> OcrOutcome:
    return OcrOutcome(OcrStatus.SUCCESS, text=text)


def _image_bytes(
    image_format: str,
    color: tuple[int, int, int] = (40, 80, 120),
) -> bytes:
    with BytesIO() as buffer:
        Image.new("RGB", (2, 2), color).save(buffer, format=image_format)
        return buffer.getvalue()


def _png_bytes(color: tuple[int, int, int] = (40, 80, 120)) -> bytes:
    return _image_bytes("PNG", color)


_SAMPLE_PDF_BYTES = b"%PDF-1.4\n% embedded-resource fixture\n"


def test_normalize_resource_key_strips_query_fragment_and_parent_segments() -> None:
    assert normalize_resource_key("../Images/pic.png?x=1#frag") == "images/pic.png"


def test_normalize_resource_key_restores_encoded_container_href() -> None:
    """F-G2-004: EPUB-style hrefs normalize after URL decoding and slash repair."""
    assert normalize_resource_key(r"../Images\space%20name.PNG?cache=1#page") == "images/space name.png"


def test_writer_creates_unique_artifacts_and_links(tmp_path) -> None:
    context = _context(tmp_path)
    first_image = _png_bytes((40, 80, 120))
    second_image = _png_bytes((120, 80, 40))
    resources = [
        MarkdownResource(
            source_key="Images/pic.png",
            suggested_name="pic.png",
            media_type="image/png",
            data=first_image,
        ),
        MarkdownResource(
            source_key="Other/pic.png",
            suggested_name="pic.png",
            media_type="image/png",
            data=second_image,
        ),
    ]

    written = MarkdownResourceWriter().write_all(
        context,
        resources,
        image_link_style="markdown_embed",
    )

    assert set(written) == {"images/pic.png", "pic.png", "other/pic.png"}
    assert written["images/pic.png"].suggested_name == "pic.png"
    assert written["other/pic.png"].suggested_name == "pic-2.png"
    assert written["images/pic.png"].markdown_link == "![pic.png](pic.png)"
    assert written["other/pic.png"].markdown_link == "![pic-2.png](pic-2.png)"
    assert context.workspace.artifacts[0].suggested_name == "pic.png"
    assert context.workspace.artifacts[1].suggested_name == "pic-2.png"
    assert (tmp_path / "artifact_1.png").read_bytes() == first_image
    assert (tmp_path / "artifact_2.png").read_bytes() == second_image


def test_writer_renders_base64_images_without_artifacts(tmp_path) -> None:
    context = _context(tmp_path)
    image_bytes = _png_bytes()
    resources = [
        MarkdownResource(
            source_key="Images/pic.png",
            suggested_name="pic.png",
            media_type="image/png",
            data=image_bytes,
        )
    ]

    written = MarkdownResourceWriter().write_all(
        context,
        resources,
        image_link_style="markdown_embed",
        image_mode="base64",
    )

    item = written["images/pic.png"]
    encoded = base64.b64encode(image_bytes).decode("ascii")
    assert item.markdown_link == f"![pic.png](data:image/png;base64,{encoded})"
    assert item.artifact is None
    assert item.artifacts == ()
    assert context.workspace.artifacts == []


def test_writer_base64_images_consume_compression_semantics(tmp_path) -> None:
    source = BytesIO()
    Image.effect_noise((512, 512), 100).convert("RGB").save(source, format="PNG")
    source_bytes = source.getvalue()
    context = _context(
        tmp_path,
        config_snapshot={
            "conversion": {
                "export": {
                    "base64_compress_enabled": True,
                    "base64_compress_threshold_kb": 100,
                }
            }
        },
    )
    written = MarkdownResourceWriter().write_all(
        context,
        [
            MarkdownResource(
                source_key="Images/noise.png",
                suggested_name="noise.png",
                media_type="image/png",
                data=source_bytes,
            )
        ],
        image_link_style="markdown_embed",
        image_mode="base64",
    )

    item = written["images/noise.png"]
    match = re.search(r"data:([^;]+);base64,([A-Za-z0-9+/=]+)", item.markdown_link)
    assert match is not None
    payload = base64.b64decode(match.group(2))
    assert match.group(1) == "image/jpeg"
    assert payload.startswith(b"\xff\xd8\xff")
    assert len(payload) <= 100 * 1024
    assert len(payload) < len(source_bytes)
    assert item.artifact is None
    assert context.workspace.artifacts == []


def test_writer_omits_images_without_artifacts(tmp_path) -> None:
    context = _context(tmp_path)
    resources = [
        MarkdownResource(
            source_key="Images/pic.png",
            suggested_name="pic.png",
            media_type="image/png",
            data=_png_bytes(),
        )
    ]

    written = MarkdownResourceWriter().write_all(
        context,
        resources,
        image_mode="omit",
    )

    item = written["images/pic.png"]
    assert item.markdown_link == "<!-- image omitted: pic.png -->"
    assert item.artifact is None
    assert item.artifacts == ()
    assert context.workspace.artifacts == []


def test_writer_renders_embed_images_as_local_references(tmp_path) -> None:
    context = _context(tmp_path)
    resources = [
        MarkdownResource(
            source_key="Images/pic.png",
            suggested_name="pic.png",
            media_type="image/png",
            data=_png_bytes(),
        )
    ]

    written = MarkdownResourceWriter().write_all(
        context,
        resources,
        image_link_style="markdown_embed",
        image_mode="embed",
    )

    assert written["images/pic.png"].markdown_link == "![pic.png](./pic.png)"
    assert context.workspace.artifacts[0].suggested_name == "pic.png"


def test_writer_renders_non_image_resources_as_normal_links(tmp_path) -> None:
    context = _context(tmp_path)
    resources = [
        MarkdownResource(
            source_key="Attachments/doc.pdf",
            suggested_name="doc.pdf",
            media_type="application/pdf",
            data=_SAMPLE_PDF_BYTES,
        )
    ]

    written = MarkdownResourceWriter().write_all(context, resources)

    assert written["attachments/doc.pdf"].markdown_link == "[doc](doc.pdf)"
    assert context.workspace.artifacts[0].kind == "auxiliary"
    assert (tmp_path / "artifact_1.pdf").read_bytes() == _SAMPLE_PDF_BYTES


def test_writer_renders_markdown_resources_with_md_file_link_style(tmp_path) -> None:
    context = _context(
        tmp_path,
        config_snapshot={"link": {"format": {"md_file_link_style": "wiki_link"}}},
    )
    resources = [
        MarkdownResource(
            source_key="Attachments/note.md",
            suggested_name="note.md",
            media_type="text/markdown",
            data=b"# Note",
        ),
        MarkdownResource(
            source_key="Attachments/doc.pdf",
            suggested_name="doc.pdf",
            media_type="application/pdf",
            data=_SAMPLE_PDF_BYTES,
        ),
    ]

    written = MarkdownResourceWriter().write_all(context, resources)

    assert written["attachments/note.md"].markdown_link == "[[note.md]]"
    assert written["attachments/doc.pdf"].markdown_link == "[doc](doc.pdf)"


@pytest.mark.parametrize(
    ("image_format", "expected_format", "expected_media_type"),
    [
        ("JPEG", "jpeg", "image/jpeg"),
        ("PNG", "png", "image/png"),
        ("GIF", "gif", "image/gif"),
        ("BMP", "bmp", "image/bmp"),
        ("TIFF", "tiff", "image/tiff"),
        ("WEBP", "webp", "image/webp"),
    ],
)
def test_writer_routes_images_and_mime_from_content_with_unrecognized_suffix(
    tmp_path,
    image_format: str,
    expected_format: str,
    expected_media_type: str,
) -> None:
    context = _context(tmp_path)
    resource = MarkdownResource(
        source_key=f"Images/{image_format.lower()}.resource",
        suggested_name=f"{image_format.lower()}.resource",
        media_type="application/octet-stream",
        data=_image_bytes(image_format),
    )

    written = MarkdownResourceWriter().write_all(
        context,
        [resource],
        image_link_style="markdown_embed",
    )

    item = written[normalize_resource_key(resource.source_key)]
    assert item.markdown_link == f"![{resource.suggested_name}]({resource.suggested_name})"
    assert item.artifact is not None
    assert item.artifact.kind == "image"
    assert item.artifact.media_type == expected_media_type
    assert item.artifact.metadata["detected_format"] == expected_format


@pytest.mark.parametrize(
    ("content", "expected_media_type"),
    [
        (b"Plain embedded text.\n", "text/plain"),
        (b"# Embedded Markdown\n\nBody.\n", "text/markdown"),
    ],
    ids=["plain-text", "markdown"],
)
def test_writer_routes_text_content_as_markdown_with_unrecognized_suffix(
    tmp_path,
    content: bytes,
    expected_media_type: str,
) -> None:
    context = _context(tmp_path)
    resource = MarkdownResource(
        source_key="Attachments/note.resource",
        suggested_name="note.resource",
        media_type="image/png",
        data=content,
    )

    written = MarkdownResourceWriter().write_all(context, [resource])

    item = written["attachments/note.resource"]
    assert item.markdown_link == "![[note.resource]]"
    assert item.artifact is not None
    assert item.artifact.kind == "auxiliary"
    assert item.artifact.media_type == expected_media_type


@pytest.mark.parametrize(
    "payload",
    [b"\x00\x01\x02\x03" * 8, b"\x89PNG\r\n\x1a\ntruncated-image"],
    ids=["fake-image", "corrupt-image"],
)
def test_writer_does_not_promote_invalid_content_from_name_or_declared_mime(
    tmp_path,
    payload: bytes,
) -> None:
    context = _context(tmp_path)
    resource = MarkdownResource(
        source_key="Images/disguised.png",
        suggested_name="disguised.png",
        media_type="image/png",
        data=payload,
    )

    written = MarkdownResourceWriter().write_all(context, [resource])

    item = written["images/disguised.png"]
    assert item.markdown_link == "[disguised](disguised.png)"
    assert item.artifact is not None
    assert item.artifact.kind == "auxiliary"
    assert item.artifact.media_type == "application/octet-stream"


def test_writer_inlines_image_ocr_in_main_markdown(tmp_path, monkeypatch) -> None:
    context = _context(tmp_path)
    monkeypatch.setattr(
        "docwen_plugin_markup.markdown_resources.run_ocr_outcome",
        lambda _path, **_kwargs: _successful_ocr("first line\nsecond line"),
    )
    resources = [
        MarkdownResource(
            source_key="Images/pic.png",
            suggested_name="pic.png",
            media_type="image/png",
            data=_png_bytes(),
        )
    ]

    written = MarkdownResourceWriter().write_all(
        context,
        resources,
        image_link_style="markdown_embed",
        enable_ocr=True,
        ocr_placement="main_md",
        source_format="html",
    )

    item = written["images/pic.png"]
    assert item.markdown_link == "![pic.png](pic.png)\n\n> first line\n> second line"
    assert [artifact.kind for artifact in item.artifacts] == ["image"]


def test_writer_passes_ocr_language_and_locale_to_core(tmp_path, monkeypatch) -> None:
    context = _context(tmp_path)
    calls: list[tuple[str, str, str]] = []

    def fake_ocr(
        _path: str,
        *,
        source_format: str,
        ocr_language: str | None = None,
        current_locale: str = "zh_CN",
    ) -> OcrOutcome:
        calls.append((source_format, ocr_language or "", current_locale))
        return _successful_ocr("bonjour")

    monkeypatch.setattr("docwen_plugin_markup.markdown_resources.run_ocr_outcome", fake_ocr)
    resources = [
        MarkdownResource(
            source_key="Images/pic.png",
            suggested_name="pic.png",
            media_type="image/png",
            data=_png_bytes(),
        )
    ]

    MarkdownResourceWriter().write_all(
        context,
        resources,
        enable_ocr=True,
        ocr_language="latin",
        current_locale="fr_FR",
    )

    assert calls == [("png", "latin", "fr_FR")]


def test_writer_creates_image_ocr_sidecar(tmp_path, monkeypatch) -> None:
    context = _context(tmp_path)
    monkeypatch.setattr(
        "docwen_plugin_markup.markdown_resources.run_ocr_outcome",
        lambda _path, **_kwargs: _successful_ocr("sidecar text"),
    )
    resources = [
        MarkdownResource(
            source_key="Images/pic.png",
            suggested_name="pic.png",
            media_type="image/png",
            data=_png_bytes(),
        )
    ]

    written = MarkdownResourceWriter().write_all(
        context,
        resources,
        image_link_style="markdown_embed",
        enable_ocr=True,
        ocr_placement="image_md",
        source_format="html",
    )

    item = written["images/pic.png"]
    assert item.markdown_link.endswith("pic_ocr.md]]")
    assert [artifact.kind for artifact in item.artifacts] == ["image", "auxiliary"]
    sidecar_artifact = item.artifacts[1]
    assert sidecar_artifact.suggested_name == "pic_ocr.md"
    sidecar_text = (tmp_path / "artifact_2.md").read_text(encoding="utf-8")
    assert "![pic.png](pic.png)" in sidecar_text
    assert "> sidecar text" in sidecar_text


def test_writer_can_ocr_without_registering_image_artifact(tmp_path, monkeypatch) -> None:
    context = _context(tmp_path)
    monkeypatch.setattr(
        "docwen_plugin_markup.markdown_resources.run_ocr_outcome",
        lambda _path, **_kwargs: _successful_ocr("ocr only"),
    )
    resources = [
        MarkdownResource(
            source_key="Images/pic.png",
            suggested_name="pic.png",
            media_type="image/png",
            data=_png_bytes(),
        )
    ]

    written = MarkdownResourceWriter().write_all(
        context,
        resources,
        enable_ocr=True,
        ocr_placement="main_md",
        keep_resource_artifacts=False,
    )

    item = written["images/pic.png"]
    assert item.markdown_link == "> ocr only"
    assert item.artifact is None
    assert item.artifacts == ()
    assert context.workspace.artifacts == []


def test_writer_nonempty_snapshot_owns_link_and_ocr_policy(tmp_path, monkeypatch) -> None:
    context = _context(
        tmp_path,
        config_snapshot={
            "link": {
                "format": {
                    "image_link_style": "markdown_embed",
                    "md_file_link_style": "markdown_link",
                }
            },
            "conversion": {
                "ocr_output": {
                    "show_blockquote_title": True,
                    "blockquote_title_override_by_locale": {"zh_CN": "SNAPSHOT TITLE"},
                }
            },
        },
        ocr_blockquote_title="SNAPSHOT TITLE",
    )
    monkeypatch.setattr(
        "docwen_plugin_markup.markdown_resources.run_ocr_outcome",
        lambda _path, **_kwargs: _successful_ocr("owned text"),
    )
    written = MarkdownResourceWriter().write_all(
        context,
        [
            MarkdownResource(
                source_key="Images/pic.png",
                suggested_name="pic.png",
                media_type="image/png",
                data=_png_bytes(),
            )
        ],
        enable_ocr=True,
        ocr_placement="main_md",
    )

    content = written["images/pic.png"].markdown_link
    assert content.startswith("![pic.png](pic.png)")
    assert "> **SNAPSHOT TITLE**" in content


@pytest.mark.parametrize(
    "status",
    [
        OcrStatus.UNAVAILABLE,
        OcrStatus.MODEL_MISSING,
        OcrStatus.INITIALIZATION_FAILED,
        OcrStatus.RECOGNITION_FAILED,
    ],
)
def test_writer_reports_typed_ocr_failure_and_continues_with_later_images(
    tmp_path,
    monkeypatch,
    status: OcrStatus,
) -> None:
    context = _context(tmp_path)
    outcomes = iter(
        [
            OcrOutcome(status, message="deterministic OCR failure"),
            _successful_ocr("later image text"),
        ]
    )
    monkeypatch.setattr(
        "docwen_plugin_markup.markdown_resources.run_ocr_outcome",
        lambda _path, **_kwargs: next(outcomes),
    )
    resources = [
        MarkdownResource(
            source_key="Images/first.png",
            suggested_name="first.png",
            media_type="image/png",
            data=_png_bytes((40, 80, 120)),
        ),
        MarkdownResource(
            source_key="Images/second.png",
            suggested_name="second.png",
            media_type="image/png",
            data=_png_bytes((120, 80, 40)),
        ),
    ]

    written = MarkdownResourceWriter().write_all(
        context,
        resources,
        image_link_style="markdown_embed",
        enable_ocr=True,
        ocr_placement="main_md",
        source_format="epub",
    )

    assert written["images/first.png"].markdown_link == "![first.png](first.png)"
    assert written["images/second.png"].markdown_link.endswith("> later image text")
    assert len(context.progress.diagnostics) == 2
    diagnostic = context.progress.diagnostics[0]
    assert diagnostic.level == "warning"
    assert diagnostic.code == "OCR-BEST-EFFORT"
    assert diagnostic.location == "first.png"
    assert f"status={status.value}" in diagnostic.message
    assert "epub image first.png" in diagnostic.message
    assert "deterministic OCR failure" not in diagnostic.message
    success_warning = context.progress.diagnostics[1]
    assert success_warning.code == "OCR-BEST-EFFORT"
    assert success_warning.location == "second.png"
    assert "status=success" in success_warning.message
    assert "later image text" in written["images/second.png"].markdown_link


def test_writer_warns_that_no_text_may_be_a_missed_best_effort_result(tmp_path, monkeypatch) -> None:
    context = _context(tmp_path)
    monkeypatch.setattr(
        "docwen_plugin_markup.markdown_resources.run_ocr_outcome",
        lambda _path, **_kwargs: OcrOutcome(OcrStatus.NO_TEXT),
    )

    written = MarkdownResourceWriter().write_all(
        context,
        [
            MarkdownResource(
                source_key="Images/blank.png",
                suggested_name="blank.png",
                media_type="image/png",
                data=_png_bytes((255, 255, 255)),
            )
        ],
        image_link_style="markdown_embed",
        enable_ocr=True,
        ocr_placement="main_md",
        source_format="html",
    )

    assert written["images/blank.png"].markdown_link == "![blank.png](blank.png)"
    assert len(context.progress.diagnostics) == 1
    diagnostic = context.progress.diagnostics[0]
    assert diagnostic.code == "OCR-BEST-EFFORT"
    assert diagnostic.location == "blank.png"
    assert "status=no_text" in diagnostic.message
    assert "may have been missed" in diagnostic.message
