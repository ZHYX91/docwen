"""No-hardcoded-English scan for GUI widgets.

Scans GUI source files for visible English string patterns that should use i18n
``t()`` calls instead, following the current GUI and configuration contracts.

Policy: visible GUI strings must go through i18n.  Explicit allowlist for
technical tokens (format names, log levels, env vars) that are NOT translated.
"""

from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

GUI_SRC = Path(__file__).resolve().parents[1] / "src" / "docwen_gui"

# ── Allowlisted words / patterns ─────────────────────────────────────────

# These patterns are NOT considered hardcoded-English violations:
# - Technical format names, log levels, env vars
# - Qt property values (object names, class names)
# - Test files
# - Python keywords and stdlib identifiers
ALLOWED_PATTERNS = re.compile(
    r"(PDF|DOCX?|ODT|RTF|WPS|XLSX?|ODS|CSV|TSV|"
    r"PNG|JPE?G|GIF|BMP|TIF{1,2}|WebP|HEIC|HEIF|"
    r"OFD|XPS|EPUB|ENEX|PPTX?|MHT|MHTML|"
    r"HTML|HTM|Markdown|YAML|JSON|TOML|"
    r"Qt\.|QWidget|Q[A-Z]\w+|"
    r"DEBUG|INFO|WARNING|ERROR|CRITICAL|"
    r"Ctrl\+|Esc|Enter|Return|"
    r"KB|MB|GB|DPI|RGB|OCR|"
    r"setObjectName|setProperty|setWindowTitle|"
    r"QSS|CSS|XML|API|IPC|UI|UX|"
    r"regex|callback|fixture|stub|mock)"
)


def _collect_hardcoded_strings(file_path: Path) -> list[tuple[int, str, str]]:
    """Collect hardcoded English strings in *file_path*.

    Returns list of (line_number, string_value, context_hint).
    """
    issues: list[tuple[int, str, str]] = []
    source = file_path.read_text(encoding="utf-8")
    lines = source.splitlines()

    # Parse AST to find string literals
    try:
        tree = ast.parse(source, filename=str(file_path))
    except SyntaxError:
        return issues
    _attach_parents(tree)

    for node in ast.walk(tree):
        # String literals in function calls
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            line_num = node.lineno
            line = lines[line_num - 1] if line_num <= len(lines) else ""
            text = node.value.strip()
            if not text or len(text) < 4:
                continue  # too short to be a meaningful UI label

            # Skip if matches allowlist
            if ALLOWED_PATTERNS.search(text):
                continue

            # Skip technical-looking strings (paths, patterns, etc.)
            if re.match(r"^[/\\]|[{}()<>*?]", text):
                continue

            if _is_docstring(node):
                continue

            if not _is_visible_ui_context(node):
                continue

            if _is_in_non_ui_context(node, line):
                continue

            # Check if this is an English word/phrase likely to be UI text
            # (contains at least one ASCII letter and looks like a readable phrase)
            if not re.search(r"[A-Za-z]{3,}", text):
                continue

            # Skip if t() or _t() call on same line (i18n already used)
            if re.search(r"\bt\(", line):
                continue

            # Skip if _t( helper pattern: text after _t( is the fallback
            if re.search(r"_t\(", line):
                continue

            # Skip if there's a t() call or _t() call nearby (same line)
            if "t(" in line or "_t(" in line or "t(" in line:
                continue

            issues.append((line_num, node.value, text[:80]))

    return issues


def _attach_parents(tree: ast.AST) -> None:
    """Attach parent links so docstrings can be identified reliably."""
    for parent in ast.walk(tree):
        for child in ast.iter_child_nodes(parent):
            vars(child)["_parent"] = parent


def _is_in_non_ui_context(node: ast.AST, line_text: str) -> bool:
    """Check if string literal is inside a non-UI call on the same line.

    Simple heuristic: if the source line contains setObjectName, setProperty,
    or findChild calls, the literal is likely a config/property value, not UI text.
    """
    return any(token in line_text for token in ("setObjectName", "setProperty", "findChild", "QKeySequence"))


_VISIBLE_CONSTRUCTORS = frozenset(
    {
        "QLabel",
        "QPushButton",
        "QToolButton",
        "QCheckBox",
        "QRadioButton",
        "QGroupBox",
        "MessageBox",
        "MessageBoxBase",
    }
)

_VISIBLE_METHODS = frozenset(
    {
        "setText",
        "setWindowTitle",
        "setToolTip",
        "setAccessibleName",
        "setAccessibleDescription",
        "setPlaceholderText",
        "add_settings_card",
        "add_form_row",
        "create_settings_toggle",
        "show_status",
        "error",
        "warn",
        "warning",
        "info",
        "confirm",
        "notify",
    }
)


def _call_name(call: ast.Call) -> str:
    func = call.func
    if isinstance(func, ast.Name):
        return func.id
    if isinstance(func, ast.Attribute):
        return func.attr
    return ""


def _is_logger_call(call: ast.Call) -> bool:
    func = call.func
    if not isinstance(func, ast.Attribute):
        return False
    value = func.value
    if isinstance(value, ast.Name):
        return value.id in {"logger", "log"}
    if isinstance(value, ast.Call):
        inner = value.func
        return (
            isinstance(inner, ast.Attribute)
            and inner.attr == "getLogger"
            and isinstance(inner.value, ast.Name)
            and inner.value.id == "logging"
        )
    return False


def _is_visible_ui_context(node: ast.AST) -> bool:
    """Return True when a literal is passed to a likely visible UI API."""
    if _is_dynamic_schema_ui_literal(node):
        return True

    current: ast.AST | None = node
    while current is not None:
        parent = getattr(current, "_parent", None)
        if isinstance(parent, ast.keyword) and parent.arg in {
            "object_name",
            "objectName",
            "name",
            "key",
            "route_key",
        }:
            return False
        if isinstance(parent, ast.Call):
            if _is_logger_call(parent):
                return False
            for keyword in parent.keywords:
                if keyword.value is current and keyword.arg in {
                    "object_name",
                    "objectName",
                    "name",
                    "key",
                    "route_key",
                }:
                    return False
            name = _call_name(parent)
            if name == "notify":
                return not (parent.args and parent.args[0] is current)
            if name in _VISIBLE_CONSTRUCTORS or name in _VISIBLE_METHODS:
                return True
            if name == "addItem":
                if parent.args and parent.args[0] is current:
                    value = getattr(current, "value", "")
                    return not (isinstance(value, str) and re.fullmatch(r"[a-z][a-z0-9_]*", value))
                return len(parent.args) > 1 and parent.args[1] is current
            return False
        current = parent
    return False


_VISIBLE_SCHEMA_KEYS = frozenset({"title", "description", "text", "tooltip", "label"})


def _is_dynamic_schema_ui_literal(node: ast.AST) -> bool:
    """Return True for visible strings inside DynamicSettingsTab schema data."""
    if not _has_schema_assignment_ancestor(node):
        return False

    parent = getattr(node, "_parent", None)
    if isinstance(parent, ast.Dict):
        for key_node, value_node in zip(parent.keys, parent.values, strict=False):
            if value_node is not node:
                continue
            return (
                isinstance(key_node, ast.Constant)
                and isinstance(key_node.value, str)
                and key_node.value in _VISIBLE_SCHEMA_KEYS
            )

    if isinstance(parent, ast.Tuple) and parent.elts and parent.elts[0] is node:
        return _is_schema_items_tuple(parent)

    return False


def _has_schema_assignment_ancestor(node: ast.AST) -> bool:
    current: ast.AST | None = node
    while current is not None:
        if isinstance(current, ast.Assign):
            return any(isinstance(target, ast.Name) and target.id == "schema" for target in current.targets)
        if isinstance(current, ast.AnnAssign):
            return isinstance(current.target, ast.Name) and current.target.id == "schema"
        current = getattr(current, "_parent", None)
    return False


def _is_schema_items_tuple(tuple_node: ast.Tuple) -> bool:
    current: ast.AST | None = tuple_node
    while current is not None:
        parent = getattr(current, "_parent", None)
        if isinstance(parent, ast.Dict):
            for key_node, value_node in zip(parent.keys, parent.values, strict=False):
                if value_node is current:
                    return isinstance(key_node, ast.Constant) and key_node.value == "items"
            return False
        current = parent
    return False


def _is_docstring(node: ast.AST) -> bool:
    """Check if a string literal is a Python docstring (not UI text)."""
    parent = getattr(node, "_parent", None)
    grandparent = getattr(parent, "_parent", None)
    return (
        isinstance(parent, ast.Expr)
        and isinstance(grandparent, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef))
        and bool(grandparent.body)
        and grandparent.body[0] is parent
    )


# ── Tests ───────────────────────────────────────────────────────────────


# Files that are expected to have SOME hardcoded English
# (fixture data, test patterns, etc.)
_ALLOWED_FILES = frozenset(
    {
        "__init__.py",
        "__main__.py",
        "test_",
        "conftest.py",
        "qt_bridge/",
        "_optimization_filter.py",
        "file_types.py",  # format→category mapping, not UI
    }
)


def _is_allowed_file(path: Path) -> bool:
    normalized = path.as_posix()
    return any(exc in path.name or exc in normalized for exc in _ALLOWED_FILES)


class TestNoHardcodedEnglish:
    @pytest.mark.parametrize(
        "py_file",
        sorted(p for p in GUI_SRC.rglob("*.py") if "tests" not in str(p) and not _is_allowed_file(p)),
        ids=lambda p: str(p.relative_to(GUI_SRC.parent)),
    )
    def test_no_hardcoded_english(self, py_file: Path) -> None:
        issues = _collect_hardcoded_strings(py_file)
        if issues:
            rel = py_file.relative_to(GUI_SRC.parent.parent)
            msg = f"{rel}: {len(issues)} potential hardcoded English strings:\n"
            for line_num, _value, snippet in issues[:10]:
                msg += f"  L{line_num}: {snippet}\n"
            if len(issues) > 10:
                msg += f"  ... and {len(issues) - 10} more\n"
            raise AssertionError(msg)

    def test_detector_reports_visible_untranslated_literal(self, tmp_path: Path) -> None:
        source = tmp_path / "visible_literal.py"
        source.write_text('label = QLabel("Visible English label")\n', encoding="utf-8")

        assert _collect_hardcoded_strings(source) == [(1, "Visible English label", "Visible English label")]

    def test_detector_ignores_docstrings_and_i18n_fallbacks(self, tmp_path: Path) -> None:
        source = tmp_path / "translated_literal.py"
        source.write_text(
            '"""English module docstring."""\nlabel = QLabel(t("labels.visible", "Translated fallback"))\n',
            encoding="utf-8",
        )

        assert _collect_hardcoded_strings(source) == []
