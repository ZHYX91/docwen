"""Golden / semantic parity tests for ENEX/HTML/MHTML → MD input routes.

Coverage:
  ROUTE-ENEX-001  (enex → md)
  ROUTE-HTML-001  (html → md)
  ROUTE-MHTML-001 (mhtml → md)
  ROUTE-HTM-001   (htm → md)
  ROUTE-MHT-001   (mht → md)

All tests use the full runtime pipeline with MarkupPlugin.
"""

from __future__ import annotations

import base64
import hashlib
import json
import tempfile
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO
from pathlib import Path
from typing import Any

import pytest
import yaml
from PIL import Image

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.text.ocr import OcrOutcome, OcrStatus

pytestmark = pytest.mark.golden

_PROJECT_ROOT = Path(__file__).resolve().parents[4]

_HTML_OLD_SYSTEM_FIXTURE = (
    _PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_html_to_markdown_semantics.json"
)

_MHTML_OLD_SYSTEM_FIXTURE = (
    _PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_mhtml_to_markdown_semantics.json"
)

_ENEX_OLD_SYSTEM_FIXTURE = (
    _PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_enex_to_markdown_semantics.json"
)

_EPUB_OLD_SYSTEM_FIXTURE = (
    _PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_epub_to_markdown_semantics.json"
)


def _successful_ocr(text: str) -> OcrOutcome:
    return OcrOutcome(OcrStatus.SUCCESS, text=text)


def _test_png_bytes(color: tuple[int, int, int] = (40, 80, 120)) -> bytes:
    with BytesIO() as buffer:
        Image.new("RGB", (2, 2), color).save(buffer, format="PNG")
        return buffer.getvalue()


def _test_png_base64(color: tuple[int, int, int] = (40, 80, 120)) -> str:
    return base64.b64encode(_test_png_bytes(color)).decode("ascii")


@pytest.fixture
def pipeline():
    """Build the full runtime pipeline with MarkupPlugin."""
    from docwen_plugin_markup import MarkupPlugin
    from docwen_runtime.engine.route_resolver import RouteResolver
    from docwen_runtime.engine.task_manager import TaskManager
    from docwen_runtime.output.finalizer import OutputFinalizer
    from docwen_runtime.plugin_registry.registry import PluginRegistry
    from docwen_runtime.workspace.manager import WorkspaceManager

    plugin = MarkupPlugin()
    registry = PluginRegistry()
    registry.register(plugin)

    resolver = RouteResolver(registry)
    ws_root = tempfile.mkdtemp(prefix="docwen_markup_")
    ws_mgr = WorkspaceManager(root_dir=ws_root)
    finalizer = OutputFinalizer()
    task_mgr = TaskManager(registry, resolver, ws_mgr, finalizer)

    yield plugin, task_mgr, ws_mgr
    ws_mgr.cleanup_all()
    import shutil

    shutil.rmtree(ws_root, ignore_errors=True)


def _run_request(
    task_mgr,
    input_path,
    source_format,
    output_dir,
    *,
    config_snapshot: dict[str, Any] | None = None,
    _on_event=None,
    **options,
) -> Any:
    """Run a single conversion request through the task manager."""
    opts: dict = {"to_md_keep_images": True}
    opts.update(options)

    request = ConversionRequest(
        request_id="markup-route-test",
        input_refs=[
            FileRef(
                path=str(input_path),
                format=source_format,
                category="markup",
                size_bytes=Path(input_path).stat().st_size,
            )
        ],
        target_format="md",
        output_policy=OutputPolicy(output_dir=str(output_dir)),
        options=opts,
        config_snapshot=config_snapshot or {},
    )
    return task_mgr.execute_single(request, on_event=_on_event)


def _load_html_old_system_fixture() -> dict[str, Any]:
    return json.loads(_HTML_OLD_SYSTEM_FIXTURE.read_text(encoding="utf-8"))


def _load_mhtml_old_system_fixture() -> dict[str, Any]:
    return json.loads(_MHTML_OLD_SYSTEM_FIXTURE.read_text(encoding="utf-8"))


def _load_enex_old_system_fixture() -> dict[str, Any]:
    return json.loads(_ENEX_OLD_SYSTEM_FIXTURE.read_text(encoding="utf-8"))


def _load_epub_old_system_fixture() -> dict[str, Any]:
    return json.loads(_EPUB_OLD_SYSTEM_FIXTURE.read_text(encoding="utf-8"))


def _write_html_multi_resource_probe(path: Path, probe: dict[str, Any]) -> None:
    path.write_text(probe["body"], encoding="utf-8")
    resource_dir = path.parent / probe["resource_dirname"]
    resource_dir.mkdir()
    for index, image in enumerate(probe["images"]):
        color = (40 + index * 80, 80, 120)
        (resource_dir / image["filename"]).write_bytes(_test_png_bytes(color))


def _write_enex_multi_resource_probe(path: Path, probe: dict[str, Any]) -> None:
    body = str(probe["body_template"])
    for index, resource in enumerate(probe["resources"]):
        color = (40 + index * 80, 80, 120)
        payload = _test_png_bytes(color)
        body = body.replace(resource["data_base64"], base64.b64encode(payload).decode("ascii"))
        body = body.replace(resource["md5"], hashlib.md5(payload).hexdigest())
    path.write_text(body, encoding="utf-8")


def _write_epub_multi_resource_probe(path: Path, probe: dict[str, Any]) -> None:
    from ebooklib import epub

    book = epub.EpubBook()
    book.set_identifier(probe["identifier"])
    book.set_title(probe["title"])
    book.set_language(probe["language"])
    book.add_author(probe["author"])
    for index, image in enumerate(probe["images"]):
        color = (40 + index * 80, 80, 120)
        book.add_item(
            epub.EpubItem(
                uid=image["uid"],
                file_name=image["filename"],
                media_type=image["media_type"],
                content=_test_png_bytes(color),
            )
        )
    chapter = epub.EpubHtml(
        title=probe["chapter_title"],
        file_name=probe["chapter_filename"],
        lang=probe["language"],
    )
    chapter.content = probe["chapter_html"]
    book.add_item(chapter)
    book.toc = [epub.Link(probe["chapter_filename"], probe["chapter_title"], "probe-chapter")]
    book.spine = ["nav", chapter]
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    epub.write_epub(str(path), book)


def _assert_yaml_title(content: str, title_key: str, title_value: str) -> None:
    assert content.startswith("---\n"), content[:200]
    frontmatter = yaml.safe_load(content.split("---", 2)[1])
    assert frontmatter[title_key] == title_value
    assert frontmatter["aliases"] == [title_value]


def _document_node_root(path: Path, output_dir: Path) -> Path:
    relative = path.relative_to(output_dir)
    assert len(relative.parts) >= 2
    root = output_dir / relative.parts[0]
    assert root.is_dir()
    return root


def _assert_finalized_markdown_content(result: Any, output_dir: Path, workspace_root: str) -> str:
    assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
    artifact = next(item for item in result.artifacts if item.media_type == "text/markdown" and item.is_primary)
    assert artifact.media_type == "text/markdown"
    artifact_path = Path(artifact.staging_path)
    assert artifact_path.parent.parent == output_dir
    assert artifact_path.name == f"{artifact_path.parent.name}.md"
    assert artifact_path.exists()
    manifest = next(item for item in result.artifacts if item.media_type == "application/vnd.docwen.document-node+json")
    assert Path(manifest.staging_path) == artifact_path.parent / "docwen-node.json"
    content = artifact_path.read_text(encoding="utf-8")
    assert str(Path(workspace_root)) not in content
    return content


def _write_mhtml_from_fixture(path: Path, input_mhtml: dict[str, Any]) -> None:
    html_part = MIMEText(input_mhtml["html_body"], "html", "utf-8")
    html_part["Content-Location"] = input_mhtml["html_content_location"]

    msg = MIMEMultipart("related")
    msg["Subject"] = input_mhtml["subject"]
    msg["From"] = "sender@example.com"
    msg["To"] = "receiver@example.com"
    msg.attach(html_part)

    images = input_mhtml.get("images") or [input_mhtml["image"]]
    for index, image in enumerate(images):
        media_type = str(image["media_type"])
        subtype = media_type.split("/", 1)[1] if "/" in media_type else "png"
        color = (40 + index * 80, 80, 120)
        img_part = MIMEImage(_test_png_bytes(color), _subtype=subtype)
        img_part["Content-ID"] = f"<{image['content_id']}>"
        img_part["Content-Location"] = image["content_location"]
        img_part.add_header("Content-Disposition", "inline", filename=image["filename"])
        msg.attach(img_part)

    path.write_bytes(msg.as_bytes())


@pytest.fixture
def sample_enex_file(tmp_path: Path) -> Path:
    """Create a sample ENEX file with one note containing HTML content."""
    enex_content = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE en-export SYSTEM "http://xml.evernote.com/pub/evernote-export3.dtd">
<en-export export-date="20260101T000000Z" application="Evernote" version="10.x">
<note>
    <title>Test Note Title</title>
    <content>
        <![CDATA[<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE en-note SYSTEM "http://xml.evernote.com/pub/enml2.dtd">
<en-note>
    <h1>Welcome</h1>
    <p>This is a <b>bold</b> paragraph with <i>italic</i> text.</p>
    <ul>
        <li>Item One</li>
        <li>Item Two</li>
    </ul>
    <p>A plain paragraph.</p>
</en-note>
]]>
    </content>
    <created>20260101T000000Z</created>
    <updated>20260101T000000Z</updated>
</note>
</en-export>"""
    path = tmp_path / "test.enex"
    path.write_text(enex_content, encoding="utf-8")
    return path


@pytest.fixture
def sample_enex_with_resources(tmp_path: Path) -> Path:
    """Create an ENEX file with an embedded image resource."""
    tiny_png = _test_png_bytes()
    tiny_png_b64 = base64.b64encode(tiny_png).decode("ascii")
    tiny_png_md5 = hashlib.md5(tiny_png).hexdigest()
    enex_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE en-export SYSTEM "http://xml.evernote.com/pub/evernote-export3.dtd">
<en-export export-date="20260101T000000Z" application="Evernote" version="10.x">
<note>
    <title>Note with Image</title>
    <content>
        <![CDATA[<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE en-note SYSTEM "http://xml.evernote.com/pub/enml2.dtd">
<en-note>
    <p>Image below:</p>
    <en-media hash="{tiny_png_md5}" type="image/png" />
    <p>After image.</p>
</en-note>
]]>
    </content>
    <created>20260101T000000Z</created>
    <updated>20260101T000000Z</updated>
    <resource>
        <data encoding="base64">{tiny_png_b64}</data>
        <mime>image/png</mime>
        <resource-attributes>
            <file-name>test_image.png</file-name>
        </resource-attributes>
    </resource>
</note>
</en-export>"""
    path = tmp_path / "test_with_res.enex"
    path.write_text(enex_content, encoding="utf-8")
    return path


@pytest.fixture
def sample_enex_with_markdown_resource(tmp_path: Path) -> Path:
    """Create an ENEX file with an embedded Markdown resource."""
    resource_bytes = b"# Linked Note\n"
    resource_hash = hashlib.md5(resource_bytes).hexdigest()
    resource_b64 = base64.b64encode(resource_bytes).decode("ascii")
    enex_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE en-export SYSTEM "http://xml.evernote.com/pub/evernote-export3.dtd">
<en-export export-date="20260101T000000Z" application="Evernote" version="10.x">
<note>
    <title>Note with Markdown Attachment</title>
    <content>
        <![CDATA[<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE en-note SYSTEM "http://xml.evernote.com/pub/enml2.dtd">
<en-note>
    <p>Markdown attachment:</p>
    <en-media hash="{resource_hash}" type="text/markdown" />
</en-note>
]]>
    </content>
    <created>20260101T000000Z</created>
    <updated>20260101T000000Z</updated>
    <resource>
        <data encoding="base64">{resource_b64}</data>
        <mime>text/markdown</mime>
        <resource-attributes>
            <file-name>linked_note.md</file-name>
        </resource-attributes>
    </resource>
</note>
</en-export>"""
    path = tmp_path / "test_with_md_res.enex"
    path.write_text(enex_content, encoding="utf-8")
    return path


@pytest.fixture
def sample_html_file(tmp_path: Path) -> Path:
    """Create a sample HTML file with headings, paragraphs, lists, and links."""
    html_content = """<!DOCTYPE html>
<html>
<head><title>Test HTML Document</title></head>
<body>
    <h1>Main Heading</h1>
    <p>This is a <strong>bold</strong> paragraph with <em>italic</em> text.</p>
    <h2>Subheading</h2>
    <p>A plain paragraph with a <a href="https://example.com">link</a>.</p>
    <ul>
        <li>Item A</li>
        <li>Item B</li>
        <li>Item C</li>
    </ul>
    <ol>
        <li>First</li>
        <li>Second</li>
    </ol>
    <h3>Table Section</h3>
    <table>
        <tr><th>Name</th><th>Value</th></tr>
        <tr><td>Foo</td><td>10</td></tr>
        <tr><td>Bar</td><td>20</td></tr>
    </table>
    <blockquote>
        <p>A quoted paragraph.</p>
    </blockquote>
    <pre><code>print("Hello World")</code></pre>
</body>
</html>"""
    path = tmp_path / "test.html"
    path.write_text(html_content, encoding="utf-8")
    return path


@pytest.fixture
def sample_html_file_with_companion_image(tmp_path: Path) -> Path:
    """Create an HTML file whose image lives in the standard companion folder."""
    image_dir = tmp_path / "test_with_image_files"
    image_dir.mkdir()
    image_path = image_dir / "picture.png"
    image_path.write_bytes(_test_png_bytes())
    path = tmp_path / "test_with_image.html"
    path.write_text(
        "<html><body><h1>HTML Image</h1><p><img src='picture.png'></p></body></html>",
        encoding="utf-8",
    )
    return path


@pytest.fixture
def sample_html_file_with_data_uri_image(tmp_path: Path) -> Path:
    """Create an HTML file with an inline data URI image."""
    tiny_png_b64 = _test_png_base64()
    path = tmp_path / "test_data_uri.html"
    path.write_text(
        f"<html><body><h1>HTML Data URI</h1><p><img src='data:image/png;base64,{tiny_png_b64}'></p></body></html>",
        encoding="utf-8",
    )
    return path


@pytest.fixture
def sample_html_file_with_remote_image(tmp_path: Path) -> Path:
    """Create an HTML file with a remote image reference."""
    path = tmp_path / "test_remote_image.html"
    path.write_text(
        "<html><body><h1>HTML Remote Image</h1><p><img src='https://example.com/image.png'></p></body></html>",
        encoding="utf-8",
    )
    return path


@pytest.fixture
def sample_mhtml_file(tmp_path: Path) -> Path:
    """Create a sample MHTML file with HTML content and embedded image."""
    html_part = MIMEText(
        "<html><body><h1>MHTML Test</h1><p>This is from an MHTML file.</p>"
        "<p><img src='cid:embedded-img'></p></body></html>",
        "html",
    )
    html_part["Content-Location"] = "test.html"

    msg = MIMEMultipart("related")
    msg["Subject"] = "Test MHTML"
    msg["From"] = "sender@example.com"
    msg["To"] = "receiver@example.com"
    msg.attach(html_part)

    img_bytes = _test_png_bytes()
    img_part = MIMEImage(img_bytes, _subtype="png")
    img_part["Content-ID"] = "<embedded-img>"
    img_part["Content-Location"] = "embedded.png"
    msg.attach(img_part)

    path = tmp_path / "test.mhtml"
    path.write_bytes(msg.as_bytes())
    return path


@pytest.fixture
def sample_epub_with_image(tmp_path: Path) -> Path:
    """Create a minimal EPUB with one chapter image."""
    from ebooklib import epub

    img_bytes = _test_png_bytes()

    book = epub.EpubBook()
    book.set_identifier("docwen-epub-image-test")
    book.set_title("EPUB Image Test")
    book.set_language("en")

    image = epub.EpubItem(
        uid="pic",
        file_name="images/pic.png",
        media_type="image/png",
        content=img_bytes,
    )
    chapter = epub.EpubHtml(title="Chapter 1", file_name="chapter.xhtml", lang="en")
    chapter.content = "<html><body><h1>Chapter 1</h1><p><img src='images/pic.png' /></p></body></html>"

    book.add_item(chapter)
    book.add_item(image)
    book.toc = [chapter]
    book.spine = ["nav", chapter]
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())

    path = tmp_path / "with_image.epub"
    epub.write_epub(str(path), book)
    return path


__all__ = (
    "ConversionRequest",
    "FileRef",
    "MIMEBase",
    "MIMEImage",
    "MIMEMultipart",
    "MIMEText",
    "OcrOutcome",
    "OcrStatus",
    "OutputPolicy",
    "Path",
    "_assert_finalized_markdown_content",
    "_assert_yaml_title",
    "_document_node_root",
    "_load_enex_old_system_fixture",
    "_load_epub_old_system_fixture",
    "_load_html_old_system_fixture",
    "_load_mhtml_old_system_fixture",
    "_run_request",
    "_successful_ocr",
    "_test_png_base64",
    "_test_png_bytes",
    "_write_enex_multi_resource_probe",
    "_write_epub_multi_resource_probe",
    "_write_html_multi_resource_probe",
    "_write_mhtml_from_fixture",
    "base64",
    "encoders",
    "hashlib",
    "pipeline",
    "pytest",
    "pytestmark",
    "sample_enex_file",
    "sample_enex_with_markdown_resource",
    "sample_enex_with_resources",
    "sample_epub_with_image",
    "sample_html_file",
    "sample_html_file_with_companion_image",
    "sample_html_file_with_data_uri_image",
    "sample_html_file_with_remote_image",
    "sample_mhtml_file",
)
