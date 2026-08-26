"""Semantic tests for MD heading numbering operations."""

from __future__ import annotations

import hashlib
import json
import tomllib
from pathlib import Path

import pytest
from tests.support.numbering import repository_numbering_registry

from docwen_plugin_markdown.numbering.converter import MdNumberingProcessor
from docwen_runtime.config import build_heading_cleanup_rules
from docwen_runtime.numbering.registry import NumberingSchemeInfo

from .conftest import make_context, write_temp_md

_PROJECT_ROOT = Path(__file__).resolve().parents[4]
_MD_NUMBERING_OLD_SYSTEM_FIXTURE = (
    _PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_md_numbering_semantics.json"
)
_DEFAULT_CLEANUP_RULES = build_heading_cleanup_rules(
    {
        "numbering": {
            "cleanup": tomllib.loads(
                (_PROJECT_ROOT / "configs" / "numbering" / "cleanup.toml").read_text(encoding="utf-8")
            )
        }
    }
)


def _load_md_numbering_old_system_fixture() -> dict:
    return json.loads(_MD_NUMBERING_OLD_SYSTEM_FIXTURE.read_text(encoding="utf-8"))


class _SingleSchemeRegistry:
    def __init__(self, scheme_id: str, scheme_config: dict, *, enabled: bool = True) -> None:
        self._scheme_id = scheme_id
        self._scheme_config = scheme_config
        self._enabled = enabled

    def get_scheme(self, scheme_id: str) -> NumberingSchemeInfo:
        if scheme_id != self._scheme_id:
            raise LookupError(scheme_id)
        return NumberingSchemeInfo(
            scheme_id=self._scheme_id,
            name="Custom roman-letter probe",
            description="Focused old/current custom numbering scheme probe",
            enabled=self._enabled,
            is_system=False,
            locales=(),
            levels={key: value["format"] for key, value in self._scheme_config.items()},
        )


class TestMdNumberingRemoval:
    """Tests for removing heading numbering from Markdown."""

    @pytest.mark.contract
    def test_removes_hierarchical_numbering(self):
        """Hierarchical numbering (1. 1.1) is removed from headings."""
        md_content = "## 1. Introduction\n\n## 2. Background\n\nSome text.\n\n### 2.1 Details\n"
        md_path = write_temp_md(md_content)
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={
                "remove_numbering": True,
                "add_numbering": False,
            },
            heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
        )

        processor = MdNumberingProcessor()
        result = processor.convert(ctx)

        assert result.success, f"Conversion failed: {result.error.message if result.error else 'unknown'}"
        assert len(result.artifacts) == 1
        artifact = result.artifacts[0]
        assert artifact.kind == "primary"
        assert artifact.media_type == "text/markdown"
        assert artifact.suggested_name.endswith(".md")

        # Read output and verify numbering is removed
        output_content = Path(artifact.staging_path).read_text(encoding="utf-8")
        assert "Introduction" in output_content
        assert "Background" in output_content
        assert "Details" in output_content
        # The numbering patterns should be gone
        assert "1. Introduction" not in output_content
        assert "1." not in output_content.split("Introduction")[0]

    @pytest.mark.contract
    def test_removes_bare_hierarchical_root_numbering(self):
        """Bare numeric roots emitted by hierarchical schemes are removable."""
        md_content = "# 1 Title\n\n## 1 Section\n\n## 1.1 Details\n"
        md_path = write_temp_md(md_content)
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={"remove_numbering": True, "add_numbering": False},
            heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
        )

        result = MdNumberingProcessor().convert(ctx)

        assert result.success
        output_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert output_content == "# Title\n\n## Section\n\n## Details\n"

    @pytest.mark.contract
    def test_long_fence_requires_matching_strict_closer(self):
        """Short fences and info strings never expose headings inside code."""
        md_content = (
            "````markdown\n"
            "# 1. Inside\n"
            "```not-a-closing-fence\n"
            "# 2. Still inside\n"
            "```\n"
            "# 3. Also inside\n"
            "````\n"
            "# 4. Outside\n"
            "    # 5. Indented code\n"
            "\t# 6. Tab-indented code\n"
        )
        md_path = write_temp_md(md_content)
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={"remove_numbering": True, "add_numbering": False},
            heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
        )

        result = MdNumberingProcessor().convert(ctx)

        assert result.success
        output_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "# 1. Inside" in output_content
        assert "# 2. Still inside" in output_content
        assert "# 3. Also inside" in output_content
        assert "# Outside" in output_content
        assert "# 4. Outside" not in output_content
        assert "    # 5. Indented code" in output_content
        assert "\t# 6. Tab-indented code" in output_content

    @pytest.mark.contract
    def test_removes_gongwen_numbering(self):
        """Chinese gongwen numbering is removed from headings."""
        md_content = "## 一、概述\n\n## 二、方法\n\nText."
        md_path = write_temp_md(md_content)
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={
                "remove_numbering": True,
                "add_numbering": False,
            },
            heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
        )

        processor = MdNumberingProcessor()
        result = processor.convert(ctx)

        assert result.success
        output_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "概述" in output_content
        assert "方法" in output_content

    @pytest.mark.contract
    def test_removes_legal_numbering(self):
        """Legal-style numbering (第一条) is removed from headings."""
        md_content = "## 第一条 总则\n\n## 第二条 定义\n\nText."
        md_path = write_temp_md(md_content)
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={
                "remove_numbering": True,
                "add_numbering": False,
            },
            heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
        )

        processor = MdNumberingProcessor()
        result = processor.convert(ctx)

        assert result.success
        output_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "总则" in output_content
        assert "定义" in output_content

    @pytest.mark.contract
    def test_request_cleanup_rules_are_the_only_rules_used(self):
        from docwen_core.text.heading_numbering import (
            compile_clean_rules_from_data,
        )

        request_rules = compile_clean_rules_from_data(
            [{"id": "request", "enabled": True, "pattern": r"^REQ:\s*", "level": 1}]
        )
        md_path = write_temp_md("# REQ: Request title\n# GLOBAL: Global title\n")
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={"remove_numbering": True, "add_numbering": False},
            heading_cleanup_rules=request_rules,
        )

        result = MdNumberingProcessor().convert(ctx)

        assert result.success
        output_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "# Request title" in output_content
        assert "# GLOBAL: Global title" in output_content


class TestMdNumberingAddition:
    """Tests for adding heading numbering to Markdown."""

    @pytest.mark.contract
    def test_adds_hierarchical_numbering(self):
        """Hierarchical numbering is added to untitled headings."""
        md_content = "# Title\n\n## Section A\n\n### Sub A1\n\n## Section B\n"
        md_path = write_temp_md(md_content)
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={
                "remove_numbering": False,
                "add_numbering": True,
                "numbering_scheme": "hierarchical_standard",
            },
            numbering_registry=repository_numbering_registry(),
            heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
        )

        processor = MdNumberingProcessor()
        result = processor.convert(ctx)

        assert result.success
        output_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        # The headings should now have 1, 1.1, 1.1.1, 1.2 prefixes
        # (format from TOML: "{1.arabic_half} " — no trailing dot)
        assert "1 Title" in output_content
        assert "1.1 Section A" in output_content
        assert "1.1.1 Sub A1" in output_content
        assert "1.2 Section B" in output_content

    @pytest.mark.contract
    def test_adds_gongwen_numbering(self):
        """Gongwen numbering scheme is applied to headings."""
        md_content = "# Title\n\n## Section A\n\n### Sub A1\n"
        md_path = write_temp_md(md_content)
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={
                "remove_numbering": False,
                "add_numbering": True,
                "numbering_scheme": "gongwen_standard",
            },
            numbering_registry=repository_numbering_registry(),
            heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
        )

        processor = MdNumberingProcessor()
        result = processor.convert(ctx)

        assert result.success
        output_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        # Should have gongwen style numbering
        assert "一、" in output_content
        assert "（一）" in output_content
        assert "1." in output_content  # Level 3

    @pytest.mark.contract
    def test_removes_then_adds(self):
        """Combined remove+add: old numbering is stripped, new scheme applied."""
        md_content = "## 1. Old Title\n\n## 2. Old Second\n"
        md_path = write_temp_md(md_content)
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={
                "remove_numbering": True,
                "add_numbering": True,
                "numbering_scheme": "gongwen_standard",
            },
            numbering_registry=repository_numbering_registry(),
            heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
        )

        processor = MdNumberingProcessor()
        result = processor.convert(ctx)

        assert result.success
        output_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        # Old "1." should be gone, gongwen numbering applied.
        # The headings are level 2 (##), so they get level_2 format: （一）, （二）
        # (TOML: "（{2.chinese_lower}）" — fullwidth parentheses, no extra space)
        assert "（一）Old Title" in output_content
        assert "（二）Old Second" in output_content


class TestMdNumberingValidation:
    """Tests for input validation in numbering processing."""

    @pytest.mark.contract
    def test_no_operation_selected_fails(self):
        """When neither remove nor add is set, the operation fails."""
        md_path = write_temp_md("# Title\n")
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={
                "remove_numbering": False,
                "add_numbering": False,
            },
        )

        processor = MdNumberingProcessor()
        result = processor.convert(ctx)

        assert not result.success
        assert result.error is not None
        assert result.error.error_type == "invalid_input"
        assert "MDNUM-NO-OPERATION" in result.error.diagnostic_code

    @pytest.mark.contract
    def test_numbering_result_metadata(self):
        """Numbering result includes operation metadata."""
        md_path = write_temp_md("# Title\n\n## Sub\n")
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={
                "remove_numbering": True,
                "add_numbering": False,
            },
            heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
        )

        processor = MdNumberingProcessor()
        result = processor.convert(ctx)

        assert result.success
        metadata = result.artifacts[0].metadata
        assert "remove_numbering" in metadata.get("operations", [])
        assert "original_length" in metadata
        assert "processed_length" in metadata

    @pytest.mark.parametrize(
        ("scheme", "registry", "error_type", "diagnostic_code"),
        [
            ("", repository_numbering_registry(), "invalid_input", "NUMBERING-SCHEME-REQUIRED"),
            (
                "gongwen_standard",
                None,
                "capability_unavailable",
                "NUMBERING-REGISTRY-UNAVAILABLE",
            ),
            (
                "missing_scheme",
                repository_numbering_registry(),
                "resource_not_found",
                "NUMBERING-SCHEME-NOT-FOUND",
            ),
            (
                "disabled",
                _SingleSchemeRegistry(
                    "disabled",
                    {"level_1": {"format": "{1.arabic_half} "}},
                    enabled=False,
                ),
                "capability_unavailable",
                "NUMBERING-SCHEME-DISABLED",
            ),
            (
                "empty",
                _SingleSchemeRegistry("empty", {}),
                "invalid_input",
                "NUMBERING-SCHEME-NO-LEVELS",
            ),
        ],
    )
    @pytest.mark.contract
    def test_add_numbering_rejects_unusable_exact_scheme(
        self,
        scheme: str,
        registry: object,
        error_type: str,
        diagnostic_code: str,
    ) -> None:
        md_path = write_temp_md("# Title\n")
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={
                "remove_numbering": False,
                "add_numbering": True,
                "numbering_scheme": scheme,
            },
            numbering_registry=registry,
        )

        result = MdNumberingProcessor().convert(ctx)

        assert not result.success
        assert result.error is not None
        assert result.error.error_type == error_type
        assert result.error.diagnostic_code == diagnostic_code
        assert result.artifacts == []


class TestOldSystemMdNumberingFixture:
    """Tests for the focused old-system Markdown numbering semantic fixture."""

    @pytest.mark.contract
    def test_md_numbering_matches_old_system_semantic_fixture(self):
        """Current processor matches old Tk/PySide6 remove/add numbering semantics."""
        fixture = _load_md_numbering_old_system_fixture()
        processor = MdNumberingProcessor()

        for case_name, case in fixture["cases"].items():
            if case_name == "invalid_scheme_falls_back_gongwen":
                continue
            options = case["options"]
            md_path = write_temp_md(fixture["input_markdown"])
            ctx, _workspace = make_context(
                md_path,
                target_format="md",
                action_name="process_md_numbering",
                options={
                    "remove_numbering": options["remove_numbering"],
                    "add_numbering": options["add_numbering"],
                    "numbering_scheme": options["numbering_scheme"],
                },
                numbering_registry=repository_numbering_registry(),
                heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
            )

            result = processor.convert(ctx)

            assert result.success, f"Conversion failed: {result.error.message if result.error else 'unknown'}"
            output_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert output_content == case["expected_markdown"]

    @pytest.mark.contract
    def test_unknown_numbering_scheme_is_rejected_without_old_fallback(self):
        """The 0.9 contract rejects an unknown exact scheme."""
        fixture = _load_md_numbering_old_system_fixture()
        case = fixture["cases"]["invalid_scheme_falls_back_gongwen"]
        md_path = write_temp_md(fixture["input_markdown"])
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={
                "remove_numbering": case["options"]["remove_numbering"],
                "add_numbering": case["options"]["add_numbering"],
                "numbering_scheme": case["options"]["numbering_scheme"],
            },
            numbering_registry=repository_numbering_registry(),
            heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
        )

        result = MdNumberingProcessor().convert(ctx)

        assert not result.success
        assert result.error is not None
        assert result.error.error_type == "resource_not_found"
        assert result.error.diagnostic_code == "NUMBERING-SCHEME-NOT-FOUND"
        assert result.artifacts == []

    @pytest.mark.contract
    def test_builtin_numbering_scheme_matrix_matches_old_systems(self):
        """Legal and H2-start built-in schemes match old-system projections."""
        fixture = _load_md_numbering_old_system_fixture()
        scope = fixture["scheme_matrix_probe"]
        processor = MdNumberingProcessor()

        for case_name, case in scope["cases"].items():
            options = case["options"]
            md_path = write_temp_md(scope["input_markdown"])
            ctx, _workspace = make_context(
                md_path,
                target_format="md",
                action_name="process_md_numbering",
                options={
                    "remove_numbering": options["remove_numbering"],
                    "add_numbering": options["add_numbering"],
                    "numbering_scheme": options["numbering_scheme"],
                },
                numbering_registry=repository_numbering_registry(),
                heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
            )

            result = processor.convert(ctx)

            assert result.success, f"{case_name} failed: {result.error.message if result.error else 'unknown'}"
            output_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
            assert output_content == case["expected_markdown"]

        expected_cases = list(scope["cases"])
        for project_name in ("docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"):
            assert scope["projects"][project_name]["matches_expected_cases"] == expected_cases

    @pytest.mark.contract
    def test_custom_numbering_scheme_matches_old_system_projection(self):
        """User-editable custom schemes match old-system formatter projection."""
        fixture = _load_md_numbering_old_system_fixture()
        scope = fixture["custom_scheme_probe"]
        options = scope["options"]
        md_path = write_temp_md(scope["input_markdown"])
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={
                "remove_numbering": options["remove_numbering"],
                "add_numbering": options["add_numbering"],
                "numbering_scheme": options["numbering_scheme"],
            },
            numbering_registry=_SingleSchemeRegistry(scope["scheme_id"], scope["scheme_config"]),
            heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
        )

        result = MdNumberingProcessor().convert(ctx)

        assert result.success, f"Conversion failed: {result.error.message if result.error else 'unknown'}"
        output_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert output_content == scope["expected_markdown"]
        assert result.artifacts[0].metadata["scheme"] == scope["scheme_id"]

        for project_name in ("docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"):
            assert scope["projects"][project_name]["matches_expected"] is True

    @pytest.mark.contract
    def test_malformed_numbering_scheme_matches_old_system_projection(self):
        """Malformed custom scheme placeholders match old-system fallback projection."""
        fixture = _load_md_numbering_old_system_fixture()
        scope = fixture["malformed_scheme_probe"]
        options = scope["options"]
        md_path = write_temp_md(scope["input_markdown"])
        ctx, _workspace = make_context(
            md_path,
            target_format="md",
            action_name="process_md_numbering",
            options={
                "remove_numbering": options["remove_numbering"],
                "add_numbering": options["add_numbering"],
                "numbering_scheme": options["numbering_scheme"],
            },
            numbering_registry=_SingleSchemeRegistry(scope["scheme_id"], scope["scheme_config"]),
            heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
        )

        result = MdNumberingProcessor().convert(ctx)

        assert result.success, f"Conversion failed: {result.error.message if result.error else 'unknown'}"
        output_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert output_content == scope["expected_markdown"]
        assert scope["projects"]["docwen-current"]["pre_fix_projection"] != scope["expected_markdown"]
        assert "{10.arabic_half}" not in output_content

        for project_name in ("docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"):
            assert scope["projects"][project_name]["matches_expected"] is True

    @pytest.mark.contract
    def test_broader_malformed_custom_document_matches_three_project_artifacts(self):
        """The frozen broader malformed/custom file is byte-identical across projects."""
        fixture = _load_md_numbering_old_system_fixture()
        scope = fixture["broader_malformed_custom_document_probe"]
        input_path = _PROJECT_ROOT / scope["input_fixture"]
        ctx, _workspace = make_context(
            str(input_path),
            target_format="md",
            action_name="process_md_numbering",
            options=scope["options"],
            numbering_registry=_SingleSchemeRegistry(scope["scheme_id"], scope["scheme_config"]),
            heading_cleanup_rules=_DEFAULT_CLEANUP_RULES,
        )

        result = MdNumberingProcessor().convert(ctx)

        assert result.success, f"Conversion failed: {result.error.message if result.error else 'unknown'}"
        output_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert output_content == scope["expected_markdown"]
        assert 'title: "# 元数据中的伪标题"' in output_content
        assert "```md\n# 一、代码块一级标题\n## （一）代码块二级标题\n```" in output_content
        assert "### X10/1 跳级三级标题" in output_content
        assert "###### [a] 六级标题" in output_content
        for project_name in ("docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"):
            assert scope["projects"][project_name]["matches_expected"] is True

    @pytest.mark.integration
    def test_md_numbering_old_system_fixture_finalizes_through_runtime(self, tmp_path: Path):
        """Runtime finalizer places the old-system numbering fixture output."""
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_plugin_markdown.plugin import MarkdownPlugin
        from docwen_runtime.config.loader import ConfigLoader
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        fixture = _load_md_numbering_old_system_fixture()
        case = fixture["cases"]["remove_add_gongwen"]
        input_file = tmp_path / "numbering-probe.md"
        input_file.write_text(fixture["input_markdown"], encoding="utf-8")
        output_dir = tmp_path / "out"
        output_dir.mkdir()

        registry = PluginRegistry()
        registry.register(MarkdownPlugin())
        task_mgr = TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / "workspace")),
            OutputFinalizer(),
            numbering_registry=repository_numbering_registry(),
        )
        request = ConversionRequest(
            request_id="md-numbering-runtime-old-system-fixture",
            input_refs=[FileRef(path=str(input_file), format="markdown", category="markdown")],
            target_format="md",
            action_name="process_md_numbering",
            options={
                "remove_numbering": case["options"]["remove_numbering"],
                "add_numbering": case["options"]["add_numbering"],
                "numbering_scheme": case["options"]["numbering_scheme"],
            },
            output_policy=OutputPolicy(output_dir=str(output_dir)),
            config_snapshot=ConfigLoader(
                base_dir=_PROJECT_ROOT / "configs",
                user_dir=tmp_path / "request-config",
            ).config.as_dict(),
        )

        result = task_mgr.execute_single(request)

        assert result.success, f"Conversion failed: {result.error.message if result.error else 'unknown'}"
        [primary] = [artifact for artifact in result.artifacts if artifact.is_primary]
        [manifest] = [
            artifact for artifact in result.artifacts if artifact.metadata.get("document_node_role") == "manifest"
        ]

        node_root = Path(result.metrics.extra["document_node_root"])
        primary_path = Path(primary.staging_path)
        manifest_path = Path(manifest.staging_path)
        assert node_root.parent == output_dir
        assert primary_path == node_root / f"{node_root.name}.md"
        assert manifest_path == node_root / "docwen-node.json"
        assert primary_path.read_text(encoding="utf-8") == case["expected_markdown"]

        document_node = json.loads(manifest_path.read_text(encoding="utf-8"))
        assert document_node["schema"] == "docwen.document_node.v1"
        assert document_node["task_id"] == request.request_id
        assert document_node["node_name"] == node_root.name
        assert document_node["source"] == {
            "name": input_file.name,
            "format": input_file.suffix.lstrip("."),
            "sha256": hashlib.sha256(input_file.read_bytes()).hexdigest(),
        }
        [primary_record] = document_node["artifacts"]
        primary_bytes = primary_path.read_bytes()
        assert primary_record == {
            "artifact_id": primary.artifact_id,
            "kind": primary.kind,
            "logical_path": f"{node_root.name}/{node_root.name}.md",
            "media_type": "text/markdown",
            "role": "primary",
            "size_bytes": len(primary_bytes),
            "sha256": hashlib.sha256(primary_bytes).hexdigest(),
        }
