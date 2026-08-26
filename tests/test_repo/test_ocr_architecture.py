"""Repository guardrails for OCR architecture."""

from __future__ import annotations

import ast
import importlib.metadata
import sys
import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
PYTHON_REQUIREMENT = ">=3.12,<3.13"
ONNXRUNTIME_LINUX_VERSION = "1.23.2"
ONNXRUNTIME_OTHER_VERSION = "1.24.4"


def test_ocr_runtime_and_python_release_boundary_are_exact() -> None:
    manifests = [PROJECT_ROOT / "pyproject.toml", *sorted((PROJECT_ROOT / "packages").glob("**/pyproject.toml"))]
    for manifest in manifests:
        project = tomllib.loads(manifest.read_text(encoding="utf-8"))["project"]
        assert project["requires-python"] == PYTHON_REQUIREMENT, manifest.relative_to(PROJECT_ROOT)

    root_project = tomllib.loads((PROJECT_ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]
    assert f"onnxruntime=={ONNXRUNTIME_LINUX_VERSION}; sys_platform == 'linux'" in root_project["dependencies"]
    assert f"onnxruntime=={ONNXRUNTIME_OTHER_VERSION}; sys_platform != 'linux'" in root_project["dependencies"]

    lock = tomllib.loads((PROJECT_ROOT / "uv.lock").read_text(encoding="utf-8"))
    locked_onnxruntime = [package for package in lock["package"] if package["name"] == "onnxruntime"]
    assert {package["version"] for package in locked_onnxruntime} == {
        ONNXRUNTIME_LINUX_VERSION,
        ONNXRUNTIME_OTHER_VERSION,
    }
    expected_installed = ONNXRUNTIME_LINUX_VERSION if sys.platform == "linux" else ONNXRUNTIME_OTHER_VERSION
    assert importlib.metadata.version("onnxruntime") == expected_installed


def test_no_shared_plugin_ocr_wrapper() -> None:
    removed_wrapper = PROJECT_ROOT / "packages" / "plugins" / "_ocr_utils.py"

    assert not removed_wrapper.exists(), (
        "Do not reintroduce packages/plugins/_ocr_utils.py. Shared OCR behavior belongs in docwen_core.text.ocr."
    )


@pytest.mark.parametrize(
    "relative_path",
    [
        "packages/plugins/layout/src/docwen_plugin_layout/_ocr_utils.py",
        "packages/plugins/image/src/docwen_plugin_image/_ocr_utils.py",
        "packages/plugins/optimizers/invoice_cn/src/docwen_plugin_optimizer_invoice_cn/_ocr_utils.py",
    ],
)
def test_no_plugin_local_ocr_wrappers(relative_path: str) -> None:
    assert not (PROJECT_ROOT / relative_path).exists(), (
        "Do not reintroduce plugin-local _ocr_utils wrappers. Plugins should "
        "call docwen_core.text.ocr.run_ocr_outcome directly."
    )


@pytest.mark.parametrize(
    "relative_path",
    [
        "packages/plugins/layout/src/docwen_plugin_layout",
        "packages/plugins/image/src/docwen_plugin_image",
        "packages/plugins/optimizers/invoice_cn/src/docwen_plugin_optimizer_invoice_cn",
    ],
)
def test_plugins_do_not_instantiate_rapidocr(relative_path: str) -> None:
    plugin_root = PROJECT_ROOT / relative_path
    sources = "\n".join(path.read_text(encoding="utf-8") for path in plugin_root.rglob("*.py"))

    assert "RapidOCR(" not in sources


def test_core_ocr_has_single_cache_source_without_test_hook_state() -> None:
    source = (PROJECT_ROOT / "packages/core/src/docwen_core/text/ocr.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    module_names = {
        target.id
        for node in tree.body
        if isinstance(node, (ast.Assign, ast.AnnAssign))
        for target in ([node.target] if isinstance(node, ast.AnnAssign) else node.targets)
        if isinstance(target, ast.Name)
    }

    assert "_ocr_instances" in module_names
    assert "_ocr_instance" not in module_names
    assert "test hook" not in source.lower()


def test_plugin_typed_ocr_calls_pass_language_and_locale() -> None:
    plugin_root = PROJECT_ROOT / "packages" / "plugins"
    offenders: list[str] = []
    callers: set[str] = set()
    typed_call_names = {"run_ocr_outcome"}
    for path in plugin_root.rglob("*.py"):
        if "src" not in path.relative_to(plugin_root).parts:
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            func = node.func
            if not isinstance(func, ast.Name) or func.id not in typed_call_names:
                continue
            rel = path.relative_to(PROJECT_ROOT).as_posix()
            callers.add(rel)
            keyword_values = {keyword.arg: keyword.value for keyword in node.keywords if keyword.arg}
            if any(
                not isinstance(keyword_values.get(name), ast.Name) or keyword_values[name].id != name
                for name in ("ocr_language", "current_locale")
            ):
                offenders.append(f"{rel}:{node.lineno}")

    assert callers == {
        "packages/plugins/document/src/docwen_plugin_document/to_markdown/converter.py",
        "packages/plugins/image/src/docwen_plugin_image/to_markdown/converter.py",
        "packages/plugins/layout/src/docwen_plugin_layout/preprocess.py",
        "packages/plugins/layout/src/docwen_plugin_layout/to_markdown/converter.py",
        "packages/plugins/markup/src/docwen_plugin_markup/markdown_resources.py",
        "packages/plugins/optimizers/gongwen/src/docwen_plugin_optimizer_gongwen/pipeline.py",
        "packages/plugins/optimizers/invoice_cn/src/docwen_plugin_optimizer_invoice_cn/invoice_cn/image_parser.py",
        "packages/plugins/presentation/src/docwen_plugin_presentation/pptx_md/converter.py",
        "packages/plugins/spreadsheet/src/docwen_plugin_spreadsheet/to_markdown/converter.py",
    }
    assert offenders == []


def test_plugins_do_not_preflight_typed_ocr_with_ocr_available() -> None:
    plugin_root = PROJECT_ROOT / "packages" / "plugins"
    callers: list[str] = []
    for path in plugin_root.rglob("*.py"):
        if "src" not in path.relative_to(plugin_root).parts:
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            func = node.func
            if not isinstance(func, ast.Name) or func.id != "ocr_available":
                continue
            rel = path.relative_to(PROJECT_ROOT).as_posix()
            callers.append(f"{rel}:{node.lineno}")

    assert callers == []


def test_ocr_consuming_manifests_declare_ocr_language_option() -> None:
    manifest_paths = [
        PROJECT_ROOT / "packages/plugins/document/src/docwen_plugin_document/manifest.py",
        PROJECT_ROOT / "packages/plugins/spreadsheet/src/docwen_plugin_spreadsheet/manifest.py",
        PROJECT_ROOT / "packages/plugins/layout/src/docwen_plugin_layout/manifest.py",
        PROJECT_ROOT / "packages/plugins/image/src/docwen_plugin_image/manifest.py",
        PROJECT_ROOT / "packages/plugins/markup/src/docwen_plugin_markup/manifest.py",
        PROJECT_ROOT / "packages/plugins/presentation/src/docwen_plugin_presentation/manifest.py",
        PROJECT_ROOT / "packages/plugins/optimizers/gongwen/src/docwen_plugin_optimizer_gongwen/manifest.py",
        PROJECT_ROOT / "packages/plugins/optimizers/invoice_cn/src/docwen_plugin_optimizer_invoice_cn/manifest.py",
    ]

    for manifest_path in manifest_paths:
        source = manifest_path.read_text(encoding="utf-8")
        tree = ast.parse(source, filename=str(manifest_path))
        ocr_language_nodes = [
            node
            for node in ast.walk(tree)
            if isinstance(node, ast.Dict)
            and any(isinstance(key, ast.Constant) and key.value == "ocr_language" for key in node.keys)
        ]
        assert ocr_language_nodes, f"{manifest_path} does not declare ocr_language"

        text = ast.get_source_segment(source, ocr_language_nodes[0]) or source
        assert '"auto"' in text
        assert '"japanese"' in text
        assert "OCR multi-language support (ocr_language)" not in text
        assert '"x-docwen-status": "reserved"' not in text
