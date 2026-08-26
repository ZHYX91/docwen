"""Tests for ``docwen_core.text.numbering`` and ``docwen_core.paths``.

Covers:
- F-I2b-011: ``number_to_chinese_upper`` — Chinese uppercase numerals
- F-I2b-014: ``number_to_letter_upper``    — Latin uppercase letters
- F-I2b-010..F-I2b-019: all 8 numbering styles (normal, edge, out-of-range)
- Path normalisation, safe join, directory creation, input stem/name helpers
- F-I2a-020 / F-I2a-021: ``sanitize_filename`` / ``sanitize_for_wiki_link``
  already resolved in ``docwen_core.markdown_utils`` (verified here with
  import + smoke tests to confirm continued availability).
"""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from docwen_core.markdown_utils import sanitize_filename, sanitize_for_wiki_link
from docwen_core.paths import (
    IGNORED_INPUT_DIRECTORY_NAMES,
    ensure_dir_exists,
    input_name,
    input_stem,
    normalize_path,
    safe_join_path,
    scan_input_directory,
)
from docwen_core.text.numbering import (
    number_to_arabic_full,
    number_to_chinese,
    number_to_chinese_upper,
    number_to_circled,
    number_to_letter_lower,
    number_to_letter_upper,
    number_to_roman_lower,
    number_to_roman_upper,
)

pytestmark = pytest.mark.unit


# ═══════════════════════════════════════════════════════════════════════════
# F-I2b-011: number_to_chinese_upper
# ═══════════════════════════════════════════════════════════════════════════


class TestNumberToChineseUpper:
    """Chinese uppercase (financial) numeral conversion."""

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "壹"),
            (5, "伍"),
            (10, "拾"),
            (11, "拾壹"),
            (20, "贰拾"),
            (21, "贰拾壹"),
            (23, "贰拾叁"),
            (50, "伍拾"),
            (99, "玖拾玖"),
        ],
    )
    def test_normal_range(self, num: int, expected: str) -> None:
        assert number_to_chinese_upper(num) == expected

    @pytest.mark.parametrize("num", [0, -1, -5, -100])
    def test_non_positive_returns_empty(self, num: int) -> None:
        assert number_to_chinese_upper(num) == ""

    def test_over_99_returns_arabic_string(self) -> None:
        assert number_to_chinese_upper(100) == "100"
        assert number_to_chinese_upper(999) == "999"


# ═══════════════════════════════════════════════════════════════════════════
# F-I2b-014: number_to_letter_upper
# ═══════════════════════════════════════════════════════════════════════════


class TestNumberToLetter:
    """Latin letter conversion (bijective base-26)."""

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "A"),
            (26, "Z"),
            (27, "AA"),
            (28, "AB"),
            (52, "AZ"),
            (53, "BA"),
            (702, "ZZ"),
        ],
    )
    def test_upper(self, num: int, expected: str) -> None:
        assert number_to_letter_upper(num) == expected

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "a"),
            (26, "z"),
            (27, "aa"),
            (28, "ab"),
        ],
    )
    def test_lower(self, num: int, expected: str) -> None:
        assert number_to_letter_lower(num) == expected

    @pytest.mark.parametrize("num", [0, -1, -10])
    def test_non_positive_returns_empty(self, num: int) -> None:
        assert number_to_letter_upper(num) == ""
        assert number_to_letter_lower(num) == ""


# ═══════════════════════════════════════════════════════════════════════════
# number_to_chinese (lowercase)
# ═══════════════════════════════════════════════════════════════════════════


class TestNumberToChinese:
    """Lowercase Chinese numeral conversion."""

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "一"),
            (5, "五"),
            (9, "九"),
            (10, "十"),
            (11, "十一"),
            (19, "十九"),
            (20, "二十"),
            (21, "二十一"),
            (55, "五十五"),
            (99, "九十九"),
        ],
    )
    def test_normal_range(self, num: int, expected: str) -> None:
        assert number_to_chinese(num) == expected

    def test_non_positive_returns_empty(self) -> None:
        assert number_to_chinese(0) == ""
        assert number_to_chinese(-1) == ""

    def test_over_99_returns_arabic(self) -> None:
        assert number_to_chinese(100) == "100"


# ═══════════════════════════════════════════════════════════════════════════
# number_to_circled
# ═══════════════════════════════════════════════════════════════════════════


class TestNumberToCircled:
    """Circled number conversion."""

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "①"),
            (10, "⑩"),
            (20, "⑳"),
            (30, "㉚"),
            (40, "㊵"),
            (50, "㊿"),
        ],
    )
    def test_normal_range(self, num: int, expected: str) -> None:
        assert number_to_circled(num) == expected

    @pytest.mark.parametrize("num", [0, -1, 51, 100])
    def test_out_of_range_returns_parenthesized(self, num: int) -> None:
        assert number_to_circled(num) == f"({num})"


# ═══════════════════════════════════════════════════════════════════════════
# number_to_arabic_full
# ═══════════════════════════════════════════════════════════════════════════


class TestNumberToArabicFull:
    """Full-width Arabic numeral conversion."""

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (0, "０"),
            (1, "１"),
            (9, "９"),
            (10, "１０"),
            (99, "９９"),
            (123, "１２３"),
            (1234567890, "１２３４５６７８９０"),
        ],
    )
    def test_conversions(self, num: int, expected: str) -> None:
        assert number_to_arabic_full(num) == expected


# ═══════════════════════════════════════════════════════════════════════════
# number_to_roman_upper / lower
# ═══════════════════════════════════════════════════════════════════════════


class TestNumberToRoman:
    """Roman numeral conversion."""

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "I"),
            (4, "IV"),
            (9, "IX"),
            (14, "XIV"),
            (40, "XL"),
            (90, "XC"),
            (400, "CD"),
            (900, "CM"),
            (1994, "MCMXCIV"),
            (3999, "MMMCMXCIX"),
        ],
    )
    def test_upper(self, num: int, expected: str) -> None:
        assert number_to_roman_upper(num) == expected

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "i"),
            (4, "iv"),
            (1994, "mcmxciv"),
        ],
    )
    def test_lower(self, num: int, expected: str) -> None:
        assert number_to_roman_lower(num) == expected

    @pytest.mark.parametrize("num", [0, -1, 4000, 5000])
    def test_out_of_range_returns_arabic(self, num: int) -> None:
        assert number_to_roman_upper(num) == str(num)
        assert number_to_roman_lower(num) == str(num).lower()


# ═══════════════════════════════════════════════════════════════════════════
# Path normalisation / safe join
# ═══════════════════════════════════════════════════════════════════════════


class TestNormalizePath:
    """Path normalisation behaviour."""

    def test_expands_user(self) -> None:
        result = normalize_path("~/foo")
        assert result == normalize_path(str(Path.home() / "foo"))

    def test_resolves_relative(self, tmp_path: Path) -> None:
        d = tmp_path / "sub"
        d.mkdir()
        result = normalize_path(str(d / ".." / "sub" / "."))
        assert result.endswith("sub")

    def test_lower_cases_on_windows(self) -> None:
        """On Windows the result is lower-cased; on other platforms it is
        returned as-is."""
        path = "/HOME/USER/Docs"
        result = normalize_path(path)
        if os.name == "nt":
            assert result == result.lower()
        # On Unix the path may or may not resolve depending on existence;
        # just verify it is a string.
        assert isinstance(result, str)


class TestSafeJoinPath:
    """Safe path joining with traversal protection."""

    def test_normal_join(self, tmp_path: Path) -> None:
        base = str(tmp_path)
        joined = safe_join_path(base, "subdir", "file.txt")
        assert joined == normalize_path(str(tmp_path / "subdir" / "file.txt"))

    def test_traversal_blocked(self, tmp_path: Path) -> None:
        base = str(tmp_path)
        with pytest.raises(ValueError, match="Path traversal blocked"):
            safe_join_path(base, "../../../etc/passwd")

    def test_absolute_path_inside_base(self, tmp_path: Path) -> None:
        base = str(tmp_path)
        sub = tmp_path / "legal"
        sub.mkdir()
        # Joining with an absolute path that is inside base should work
        joined = safe_join_path(base, str(sub))
        assert joined == normalize_path(str(sub.resolve()))


class TestEnsureDirExists:
    """Directory creation utility."""

    def test_creates_missing_dir(self, tmp_path: Path) -> None:
        target = tmp_path / "a" / "b" / "c"
        result = ensure_dir_exists(str(target))
        assert target.is_dir()
        assert isinstance(result, str)

    def test_existing_dir_noop(self, tmp_path: Path) -> None:
        result1 = ensure_dir_exists(str(tmp_path))
        result2 = ensure_dir_exists(str(tmp_path))
        assert result1 == result2

    def test_creates_parents(self, tmp_path: Path) -> None:
        deep = tmp_path / "x" / "y" / "z"
        ensure_dir_exists(str(deep))
        assert deep.is_dir()
        assert (tmp_path / "x").is_dir()


class TestInputDirectoryScan:
    def test_prunes_tool_directories_case_insensitively(self, tmp_path: Path) -> None:
        kept = tmp_path / "documents" / "keep.docx"
        kept.parent.mkdir()
        kept.write_text("keep", encoding="utf-8")
        for index, name in enumerate(sorted(IGNORED_INPUT_DIRECTORY_NAMES)):
            actual_name = name.upper() if index % 2 else name
            hidden = tmp_path / actual_name / "hidden.docx"
            hidden.parent.mkdir()
            hidden.write_text("hidden", encoding="utf-8")

        scan = scan_input_directory(tmp_path)

        assert scan.files == (kept,)
        assert scan.unreadable_paths == ()
        assert scan.truncated is False

    def test_limit_reports_truncation_without_overcollection(self, tmp_path: Path) -> None:
        for index in range(3):
            (tmp_path / f"{index}.txt").write_text("x", encoding="utf-8")

        scan = scan_input_directory(tmp_path, limit=2)

        assert [path.name for path in scan.files] == ["0.txt", "1.txt"]
        assert scan.truncated is True

    def test_directory_junctions_are_pruned_without_following(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        junction = tmp_path / "junction"
        junction.mkdir()
        (junction / "outside.docx").write_text("outside", encoding="utf-8")
        original_is_junction = Path.is_junction
        monkeypatch.setattr(
            Path,
            "is_junction",
            lambda path: path == junction or original_is_junction(path),
        )

        scan = scan_input_directory(tmp_path)

        assert scan.files == ()

    def test_walk_errors_are_returned_as_data(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        from docwen_core import paths

        blocked = tmp_path / "blocked"

        def fake_walk(root, *, topdown, onerror, followlinks):
            del root, topdown, followlinks
            onerror(PermissionError(13, "denied", str(blocked)))
            return iter(())

        monkeypatch.setattr(paths.os, "walk", fake_walk)

        scan = scan_input_directory(tmp_path)

        assert scan.files == ()
        assert scan.unreadable_paths == (blocked,)


class TestInputStemName:
    """Filename stem / name helpers."""

    def test_stem(self) -> None:
        assert input_stem("/a/b/财务报告_2024.pdf") == "财务报告_2024"
        assert input_stem("image.png") == "image"
        assert input_stem(Path("/tmp/file.tar.gz")) == "file.tar"

    def test_name(self) -> None:
        assert input_name("/a/b/财务报告_2024.pdf") == "财务报告_2024.pdf"
        assert input_name("image.png") == "image.png"
        assert input_name(Path("/tmp/file.tar.gz")) == "file.tar.gz"


# ═══════════════════════════════════════════════════════════════════════════
# F-I2a-020 / F-I2a-021 — sanity: sanitize_filename / sanitize_for_wiki_link
#   These are ALREADY implemented in ``docwen_core.markdown_utils`` and
#   tested in ``test_yaml_markdown_utils.py``.  The tests below serve as
#   **confirmation** that they are importable and behave correctly — this is
#   the evidence these findings are closed.
# ═══════════════════════════════════════════════════════════════════════════


class TestSanitizeFilename:
    """F-I2a-020 closure evidence: ``sanitize_filename`` available in core."""

    def test_replaces_illegal_chars(self) -> None:
        assert sanitize_filename("a/b\\c:d?e") == "a_b_c_d_e"

    def test_compresses_whitespace(self) -> None:
        assert sanitize_filename("a   b\t c") == "a b c"

    def test_strips_dots_and_spaces(self) -> None:
        assert sanitize_filename(" .name. ") == "name"

    def test_chinese_passthrough(self) -> None:
        assert sanitize_filename("财务报告 (定稿)") == "财务报告 (定稿)"

    def test_preserves_valid_chars(self) -> None:
        assert sanitize_filename("hello_world-123") == "hello_world-123"


class TestSanitizeForWikiLink:
    """F-I2a-021 closure evidence: ``sanitize_for_wiki_link`` available in core."""

    def test_strips_sensitive_chars(self) -> None:
        result = sanitize_for_wiki_link("file[1] | section#sub^id")
        assert "[1]" not in result
        assert "#sub^id" not in result
        # | is FS-illegal → replaced with _ by sanitize_filename first
        assert "file1 _ sectionsubid" in result or "file1__sectionsubid" in result

    def test_preserves_normal_chars(self) -> None:
        assert sanitize_for_wiki_link("hello world.png") == "hello world.png"

    def test_also_sanitizes_filesystem_illegal(self) -> None:
        assert sanitize_for_wiki_link("a/b:c.png") == "a_b_c.png"
