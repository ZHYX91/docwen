"""Tests for docwen_core.links — data URI image handling and embedded image
processing.

Covers F-H2-006 and F-H2-008.
"""

from __future__ import annotations

import base64
import os
from pathlib import Path

import pytest

from docwen_core.links import (
    EmbeddedImageMode,
    format_image_placeholder,
    is_data_uri_image,
    process_embedded_image,
    resolve_data_uri_image_to_temp_file,
    split_alt_text_and_size,
)

pytestmark = pytest.mark.unit


# ═══════════════════════════════════════════════════════════════════════════
# is_data_uri_image  (F-H2-006)
# ═══════════════════════════════════════════════════════════════════════════


class TestIsDataUriImage:
    """Covers F-H2-006: ``is_data_uri_image`` detection."""

    def test_png_data_uri(self) -> None:
        assert is_data_uri_image("data:image/png;base64,iVBORw0KGgo=") is True

    def test_jpeg_data_uri(self) -> None:
        assert is_data_uri_image("data:image/jpeg;base64,/9j/4AAQ") is True

    def test_gif_data_uri(self) -> None:
        assert is_data_uri_image("data:image/gif;base64,R0lGODlh") is True

    def test_svg_data_uri(self) -> None:
        assert is_data_uri_image("data:image/svg+xml;base64,PHN2Zy") is True

    def test_plain_text_not_data_uri(self) -> None:
        assert is_data_uri_image("plain text") is False

    def test_file_path_not_data_uri(self) -> None:
        assert is_data_uri_image("/path/to/image.png") is False

    def test_http_url_not_data_uri(self) -> None:
        assert is_data_uri_image("https://example.com/image.png") is False

    def test_data_uri_no_base64(self) -> None:
        # data:image/... without ";base64," — not a base64 data URI
        assert is_data_uri_image("data:image/png,some-raw-data") is False

    def test_data_text_not_image(self) -> None:
        # data:text/... is not an image
        assert is_data_uri_image("data:text/plain;base64,SGVsbG8=") is False

    def test_empty_string(self) -> None:
        assert is_data_uri_image("") is False

    def test_data_uri_with_charset(self) -> None:
        # "data:image/png;charset=utf-8;base64,..." is still a data URI image
        assert is_data_uri_image("data:image/png;charset=utf-8;base64,abc") is True


# ═══════════════════════════════════════════════════════════════════════════
# resolve_data_uri_image_to_temp_file  (F-H2-006)
# ═══════════════════════════════════════════════════════════════════════════


_SAMPLE_PNG_BYTES = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
    b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\x0f"
    b"\x00\x00\x01\x01\x00\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82"
)


def _make_data_uri(image_bytes: bytes, mime: str = "image/png") -> str:
    payload = base64.b64encode(image_bytes).decode("ascii")
    return f"data:{mime};base64,{payload}"


class TestResolveDataUriImageToTempFile:
    """Covers F-H2-006: ``resolve_data_uri_image_to_temp_file``."""

    def test_resolve_png_to_temp_file(self) -> None:
        data_uri = _make_data_uri(_SAMPLE_PNG_BYTES)
        result = resolve_data_uri_image_to_temp_file(data_uri)
        assert result is not None
        assert os.path.isfile(result)
        written = Path(result).read_bytes()
        assert written == _SAMPLE_PNG_BYTES

    def test_resolve_jpeg_mime_to_jpg_ext(self) -> None:
        data_uri = _make_data_uri(b"\xff\xd8\xff\xe0\x00\x10JFIF", mime="image/jpeg")
        result = resolve_data_uri_image_to_temp_file(data_uri)
        assert result is not None
        assert result.endswith(".jpg")

    def test_resolve_with_custom_temp_dir(self, tmp_path: Path) -> None:
        data_uri = _make_data_uri(_SAMPLE_PNG_BYTES)
        result = resolve_data_uri_image_to_temp_file(data_uri, temp_dir=str(tmp_path))
        assert result is not None
        assert result.startswith(str(tmp_path))

    def test_returns_none_for_non_data_uri(self) -> None:
        assert resolve_data_uri_image_to_temp_file("not a data uri") is None

    def test_returns_none_for_missing_base64(self) -> None:
        assert resolve_data_uri_image_to_temp_file("data:image/png,raw-bytes") is None

    def test_returns_none_for_text_data_uri(self) -> None:
        assert resolve_data_uri_image_to_temp_file("data:text/plain;base64,SGVsbG8=") is None

    def test_returns_none_for_malformed_uri(self) -> None:
        assert resolve_data_uri_image_to_temp_file("data:image/png;base64") is None
        # missing comma → no split
        assert resolve_data_uri_image_to_temp_file("data:image/png;base64,") is not None
        # empty payload still decodes (zero bytes)

    def test_returns_none_for_invalid_base64(self) -> None:
        assert resolve_data_uri_image_to_temp_file("data:image/png;base64,!!!not-valid-base64!!!") is None

    def test_respected_max_size_limit(self) -> None:
        """Payload larger than 10 MB is rejected."""
        huge_payload = base64.b64encode(b"A" * (11 * 1024 * 1024)).decode("ascii")
        data_uri = f"data:image/png;base64,{huge_payload}"
        assert resolve_data_uri_image_to_temp_file(data_uri) is None

    def test_unknown_mime_subtype_gets_fallback_ext(self) -> None:
        data_uri = _make_data_uri(_SAMPLE_PNG_BYTES, mime="image/x-unknown-fmt")
        result = resolve_data_uri_image_to_temp_file(data_uri)
        assert result is not None
        # falls back to .xunknownfmt
        assert ".xunknownfmt" in result


# ═══════════════════════════════════════════════════════════════════════════
# format_image_placeholder  (F-H2-008)
# ═══════════════════════════════════════════════════════════════════════════


class TestFormatImagePlaceholder:
    """Covers F-H2-008: ``format_image_placeholder``."""

    def test_no_dimensions(self) -> None:
        assert format_image_placeholder("photo.png") == "{{IMAGE:photo.png}}"

    def test_with_width_and_height(self) -> None:
        assert format_image_placeholder("photo.png", width=200, height=150) == ("{{IMAGE:photo.png|200|150}}")

    def test_with_width_only(self) -> None:
        assert format_image_placeholder("photo.png", width=200) == ("{{IMAGE:photo.png|200|}}")

    def test_with_height_only(self) -> None:
        assert format_image_placeholder("photo.png", height=150) == ("{{IMAGE:photo.png||150}}")

    def test_with_width_zero(self) -> None:
        assert format_image_placeholder("x.png", width=0) == "{{IMAGE:x.png|0|}}"

    def test_relative_path(self) -> None:
        assert format_image_placeholder("./assets/img.png") == ("{{IMAGE:./assets/img.png}}")

    def test_absolute_path_placeholder(self) -> None:
        # Placeholder format is purely textual — absolute paths are accepted.
        p = format_image_placeholder("/abs/path/img.png", width=100, height=50)
        assert p == "{{IMAGE:/abs/path/img.png|100|50}}"


# ═══════════════════════════════════════════════════════════════════════════
# split_alt_text_and_size  (F-H2-008)
# ═══════════════════════════════════════════════════════════════════════════


class TestSplitAltTextAndSize:
    """Covers F-H2-008: ``split_alt_text_and_size``."""

    def test_none_input(self) -> None:
        assert split_alt_text_and_size(None) == (None, None, None)

    def test_empty_string(self) -> None:
        assert split_alt_text_and_size("") == (None, None, None)

    def test_plain_text_no_dimensions(self) -> None:
        assert split_alt_text_and_size("photo") == ("photo", None, None)

    def test_width_height_both(self) -> None:
        disp, w, h = split_alt_text_and_size("photo|200x150")
        assert disp == "photo"
        assert w == 200
        assert h == 150

    def test_width_only(self) -> None:
        disp, w, h = split_alt_text_and_size("photo|200")
        assert disp == "photo"
        assert w == 200
        assert h is None

    def test_width_with_empty_height(self) -> None:
        disp, w, h = split_alt_text_and_size("photo|200x")
        assert disp == "photo"
        assert w == 200
        assert h is None

    def test_trailing_pipe_ignored(self) -> None:
        # "photo|" — trailing pipe with no content is ignored
        assert split_alt_text_and_size("photo|") == ("photo|", None, None)

    def test_empty_size_section(self) -> None:
        assert split_alt_text_and_size("photo||") == ("photo||", None, None)

    def test_non_numeric_size(self) -> None:
        # "photo|large" — non-numeric, treated as regular text
        assert split_alt_text_and_size("photo|large") == ("photo|large", None, None)

    def test_multiple_pipes_last_wins(self) -> None:
        # Only the last pipe segment is examined for dimensions
        disp, w, h = split_alt_text_and_size("a|b|200x150")
        assert disp == "a|b"
        assert w == 200
        assert h == 150

    def test_surrounded_by_spaces(self) -> None:
        disp, w, h = split_alt_text_and_size("  hello  |  300x200  ")
        assert disp == "  hello  "
        assert w == 300
        assert h == 200

    def test_zero_dimensions(self) -> None:
        disp, w, h = split_alt_text_and_size("empty|0x0")
        assert disp == "empty"
        assert w == 0
        assert h == 0

    def test_negative_dimension_is_not_digit(self) -> None:
        # "-100" is not .isdigit()
        assert split_alt_text_and_size("bad|-100") == ("bad|-100", None, None)


# ═══════════════════════════════════════════════════════════════════════════
# process_embedded_image  (F-H2-008)
# ═══════════════════════════════════════════════════════════════════════════


class TestProcessEmbeddedImage:
    """Covers F-H2-008: ``process_embedded_image`` mode dispatch."""

    _IMG = "photo.png"
    _LINK = "![[photo.png]]"

    def test_mode_embed_without_dimensions(self) -> None:
        result = process_embedded_image(self._IMG, self._LINK, mode=EmbeddedImageMode.EMBED)
        assert result == "{{IMAGE:photo.png}}"

    def test_mode_embed_with_dimensions(self) -> None:
        result = process_embedded_image(
            self._IMG,
            self._LINK,
            mode="embed",
            width=300,
            height=200,
        )
        assert result == "{{IMAGE:photo.png|300|200}}"

    def test_mode_keep(self) -> None:
        result = process_embedded_image(self._IMG, self._LINK, mode=EmbeddedImageMode.KEEP)
        assert result == "![[photo.png]]"

    def test_mode_keep_preserves_link_text_exactly(self) -> None:
        markdown_link = "![alt text](photo.png)"
        result = process_embedded_image(self._IMG, markdown_link, mode="keep")
        assert result == markdown_link

    def test_mode_extract_text_with_display_text(self) -> None:
        result = process_embedded_image(
            self._IMG,
            self._LINK,
            mode=EmbeddedImageMode.EXTRACT_TEXT,
            display_text="My Photo",
        )
        assert result == "My Photo"

    def test_mode_extract_text_fallback_to_filename(self) -> None:
        result = process_embedded_image(self._IMG, self._LINK, mode="extract_text")
        assert result == "photo.png"

    def test_mode_remove(self) -> None:
        result = process_embedded_image(self._IMG, self._LINK, mode=EmbeddedImageMode.REMOVE)
        assert result == ""

    def test_mode_remove_ignores_display_text(self) -> None:
        result = process_embedded_image(
            self._IMG,
            self._LINK,
            mode="remove",
            display_text="visible text",
        )
        assert result == ""

    def test_str_mode_accepted(self) -> None:
        """Modes can be passed as plain strings, not just enum values."""
        result = process_embedded_image(self._IMG, self._LINK, mode="embed")
        assert result == "{{IMAGE:photo.png}}"

    def test_unknown_mode_falls_back_to_embed(self) -> None:
        # Invalid mode string raises ValueError via StrEnum constructor.
        with pytest.raises(ValueError):
            process_embedded_image(self._IMG, self._LINK, mode="bogus")

    def test_display_text_in_keep_mode_is_ignored(self) -> None:
        result = process_embedded_image(
            self._IMG,
            self._LINK,
            mode="keep",
            display_text="ignored",
        )
        assert result == self._LINK

    def test_mode_extract_text_empty_display_text_falls_back(self) -> None:
        result = process_embedded_image(
            self._IMG,
            self._LINK,
            mode="extract_text",
            display_text="",
        )
        assert result == "photo.png"


# ═══════════════════════════════════════════════════════════════════════════
# User-path integration tests  (F-H2-006, F-H2-008)
# ═══════════════════════════════════════════════════════════════════════════


class TestUserPathDataUriToEmbeddedImage:
    """End-to-end path: detect data URI → materialise → embed placeholder."""

    def test_full_data_uri_to_placeholder_pipeline(self) -> None:
        """Simulate the real flow: a data URI image is detected, written to
        a temp file, then wrapped in a placeholder."""
        data_uri = _make_data_uri(_SAMPLE_PNG_BYTES)

        # Step 1: detect
        assert is_data_uri_image(data_uri) is True

        # Step 2: resolve to temp file
        temp_path = resolve_data_uri_image_to_temp_file(data_uri)
        assert temp_path is not None
        try:
            assert os.path.isfile(temp_path)
            assert os.path.getsize(temp_path) > 0

            # Step 3: wrap in placeholder
            filename = Path(temp_path).name
            placeholder = format_image_placeholder(filename)
            assert placeholder.startswith("{{IMAGE:")
            assert filename in placeholder
        finally:
            # Clean up temp file
            Path(temp_path).unlink(missing_ok=True)


class TestEmbeddedImageModeEnum:
    """Verify EmbeddedImageMode values match the cardinal four modes."""

    def test_all_modes_present(self) -> None:
        modes = set(EmbeddedImageMode)
        assert modes == {"embed", "keep", "extract_text", "remove"}

    def test_mode_values_match_strings(self) -> None:
        assert EmbeddedImageMode.EMBED.value == "embed"
        assert EmbeddedImageMode.KEEP.value == "keep"
        assert EmbeddedImageMode.EXTRACT_TEXT.value == "extract_text"
        assert EmbeddedImageMode.REMOVE.value == "remove"
