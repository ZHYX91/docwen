"""SmartDoc converter tests."""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any

import pytest

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy

pytestmark = [pytest.mark.golden, pytest.mark.contract]


# ═══════════════════════════════════════════════════════════════════════════
# Fixtures
# ═══════════════════════════════════════════════════════════════════════════


@pytest.fixture
def pipeline():
    """Build the full runtime pipeline with the real DocumentPlugin."""
    from docwen_plugin_document import DocumentPlugin
    from docwen_runtime.engine.route_resolver import RouteResolver
    from docwen_runtime.engine.task_manager import TaskManager
    from docwen_runtime.output.finalizer import OutputFinalizer
    from docwen_runtime.plugin_registry.registry import PluginRegistry
    from docwen_runtime.workspace.manager import WorkspaceManager

    plugin = DocumentPlugin()
    registry = PluginRegistry()
    registry.register(plugin)

    resolver = RouteResolver(registry)
    ws_root = tempfile.mkdtemp(prefix="docwen_smartconv_")
    ws_mgr = WorkspaceManager(root_dir=ws_root)
    finalizer = OutputFinalizer()
    task_mgr = TaskManager(registry, resolver, ws_mgr, finalizer)

    yield plugin, task_mgr, ws_mgr
    ws_mgr.cleanup_all()
    import shutil

    shutil.rmtree(ws_root, ignore_errors=True)


def _run_smartconv(task_mgr, source_format, target_format, output_dir, *, config_snapshot=None) -> Any:
    """Run a SmartConverter request via the pipeline."""
    import uuid

    dummy_content = b"dummy binary content"
    dummy_path = Path(output_dir) / f"dummy.{source_format}"
    dummy_path.write_bytes(dummy_content)

    request = ConversionRequest(
        request_id=f"smartconv-{uuid.uuid4().hex[:8]}",
        input_refs=[
            FileRef(
                path=str(dummy_path),
                format=source_format,
                category="document",
                size_bytes=len(dummy_content),
            )
        ],
        target_format=target_format,
        output_policy=OutputPolicy(output_dir=str(output_dir)),
        config_snapshot=config_snapshot or {},
    )
    return task_mgr.execute_single(request)


ALL_SMARTCONV_PAIRS: list[tuple[str, str]] = [
    ("docx", "doc"),
    ("docx", "odt"),
    ("docx", "rtf"),
    ("docx", "wps"),
    ("doc", "docx"),
    ("odt", "docx"),
    ("rtf", "docx"),
    ("wps", "docx"),
    ("doc", "odt"),
    ("doc", "rtf"),
    ("doc", "wps"),
    ("odt", "doc"),
    ("odt", "rtf"),
    ("odt", "wps"),
    ("rtf", "doc"),
    ("rtf", "odt"),
    ("rtf", "wps"),
    ("wps", "doc"),
    ("wps", "odt"),
    ("wps", "rtf"),
]


class TestSmartConverterRoutes:
    def test_docx_to_rtf_honors_configured_word_priority(self, pipeline, tmp_path, monkeypatch) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_document.to_document import converter as smartdoc_module

        calls: list[dict[str, Any]] = []

        def fake_convert(input_path, output_path, **kwargs):
            calls.append(kwargs)
            Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="Microsoft Word")

        monkeypatch.setattr(smartdoc_module, "convert_with_backend_priority", fake_convert)

        _plugin, task_mgr, _ws_mgr = pipeline
        result = _run_smartconv(
            task_mgr,
            "docx",
            "rtf",
            tmp_path,
            config_snapshot={
                "software": {"default_priority": {"word_processors": ["msoffice_word", "libreoffice", "wps_writer"]}}
            },
        )

        assert result.success is True
        assert calls[0]["source_format"] == "docx"
        assert calls[0]["backend_priority"] == ["msoffice_word", "libreoffice", "wps_writer"]
        assert set(calls[0]["com_candidates"]) == {"wps_writer", "msoffice_word"}

    def test_odt_route_honors_odt_priority_without_wps(self, pipeline, tmp_path, monkeypatch) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_document.to_document import converter as smartdoc_module

        calls: list[dict[str, Any]] = []

        def fake_convert(input_path, output_path, **kwargs):
            calls.append(kwargs)
            Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="LibreOffice")

        monkeypatch.setattr(smartdoc_module, "convert_with_backend_priority", fake_convert)

        _plugin, task_mgr, _ws_mgr = pipeline
        result = _run_smartconv(
            task_mgr,
            "odt",
            "docx",
            tmp_path,
            config_snapshot={"software": {"special_conversions": {"odt": ["libreoffice", "msoffice_word"]}}},
        )

        assert result.success is True
        assert calls[0]["source_format"] == "odt"
        assert calls[0]["backend_priority"] == ["libreoffice", "msoffice_word"]
        assert set(calls[0]["com_candidates"]) == {"msoffice_word"}

    def test_two_hop_route_resolves_priority_per_leg(self, pipeline, tmp_path, monkeypatch) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_document.to_document import converter as smartdoc_module

        calls: list[dict[str, Any]] = []

        def fake_convert(input_path, output_path, **kwargs):
            calls.append(kwargs)
            Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(smartdoc_module, "convert_with_backend_priority", fake_convert)

        _plugin, task_mgr, _ws_mgr = pipeline
        result = _run_smartconv(
            task_mgr,
            "odt",
            "rtf",
            tmp_path,
            config_snapshot={
                "software": {
                    "default_priority": {"word_processors": ["msoffice_word", "wps_writer", "libreoffice"]},
                    "special_conversions": {"odt": ["libreoffice", "msoffice_word"]},
                }
            },
        )

        assert result.success is True
        assert [call["source_format"] for call in calls] == ["odt", "docx"]
        assert [call["backend_priority"] for call in calls] == [
            ["libreoffice", "msoffice_word"],
            ["msoffice_word", "wps_writer", "libreoffice"],
        ]
        assert [set(call["com_candidates"]) for call in calls] == [
            {"msoffice_word"},
            {"wps_writer", "msoffice_word"},
        ]

    def test_two_hop_route_does_not_apply_direct_docx_best_effort_preference(
        self, pipeline, tmp_path, monkeypatch
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_document.to_document import converter as smartdoc_module

        calls: list[dict[str, Any]] = []

        def fake_convert(input_path, output_path, **kwargs):
            calls.append(kwargs)
            Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(smartdoc_module, "convert_with_backend_priority", fake_convert)

        _plugin, task_mgr, _ws_mgr = pipeline
        output_dir = tmp_path / "doc_rtf_two_hop"
        output_dir.mkdir()
        result = _run_smartconv(task_mgr, "doc", "rtf", output_dir)

        assert result.success is True
        assert [call["backend_priority"] for call in calls] == [
            ["wps_writer", "msoffice_word", "libreoffice"],
            ["wps_writer", "msoffice_word", "libreoffice"],
        ]

    @pytest.mark.parametrize("source_fmt,target_fmt", ALL_SMARTCONV_PAIRS)
    def test_routes_succeed_when_bridge_available(
        self, pipeline, tmp_path, monkeypatch, source_fmt, target_fmt
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_document.to_document import converter as smartdoc_module

        def fake_convert(input_path, output_path, **kwargs):
            Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(smartdoc_module, "convert_with_backend_priority", fake_convert)

        _plugin, task_mgr, _ws_mgr = pipeline
        output_dir = tmp_path / f"{source_fmt}_{target_fmt}"
        output_dir.mkdir()
        result = _run_smartconv(task_mgr, source_fmt, target_fmt, output_dir)

        assert result.success is True
        assert result.error is None
        assert result.artifacts
        assert result.artifacts[0].suggested_name.endswith(f".{target_fmt}")
        assert result.diagnostics[0].code == "DOCX-SMARTDOC-OK"

    @pytest.mark.parametrize(
        ("target_fmt", "loss_terms"),
        [
            ("doc", ("fields", "revisions/comments", "inline-object identities", "layout")),
            ("rtf", ("fields", "revisions/comments", "inline-object identities", "layout")),
            ("odt", ("paragraphs", "tables", "fields", "revisions", "shapes", "sections", "pagination")),
        ],
    )
    def test_docx_legacy_targets_deliver_with_typed_best_effort_warning(
        self, pipeline, tmp_path, monkeypatch, target_fmt, loss_terms
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_document.to_document import converter as smartdoc_module

        def fake_convert(input_path, output_path, **kwargs):
            candidates = kwargs["com_candidates"]
            assert candidates["msoffice_word"].suppress_new_revisions is True
            Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="Microsoft Word")

        monkeypatch.setattr(smartdoc_module, "convert_with_backend_priority", fake_convert)

        _plugin, task_mgr, _ws_mgr = pipeline
        output_dir = tmp_path / f"docx_{target_fmt}_best_effort"
        output_dir.mkdir()
        result = _run_smartconv(task_mgr, "docx", target_fmt, output_dir)

        warnings = [diagnostic for diagnostic in result.diagnostics if diagnostic.level == "warning"]
        assert result.success is True
        assert result.artifacts
        assert len(warnings) == 1
        assert warnings[0].code == "DOCX-SMARTDOC-BEST-EFFORT-LOSS"
        assert target_fmt.upper() in warnings[0].message
        assert "Microsoft Word" in warnings[0].message
        assert "source file was not modified" in warnings[0].message
        assert "review" in warnings[0].message.lower()
        assert "lossless" not in warnings[0].message.lower()
        assert all(term in warnings[0].message for term in loss_terms)
        assert (output_dir / "dummy.docx").read_bytes() == b"dummy binary content"

    @pytest.mark.parametrize(
        ("source_fmt", "target_fmt"),
        [("docx", "wps"), ("doc", "docx"), ("doc", "rtf")],
    )
    def test_best_effort_warning_is_limited_to_selected_docx_outbound_targets(
        self, pipeline, tmp_path, monkeypatch, source_fmt, target_fmt
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_document.to_document import converter as smartdoc_module

        def fake_convert(input_path, output_path, **kwargs):
            Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(smartdoc_module, "convert_with_backend_priority", fake_convert)

        _plugin, task_mgr, _ws_mgr = pipeline
        output_dir = tmp_path / f"{source_fmt}_{target_fmt}_no_best_effort_warning"
        output_dir.mkdir()
        result = _run_smartconv(task_mgr, source_fmt, target_fmt, output_dir)

        assert result.success is True
        assert not [
            diagnostic for diagnostic in result.diagnostics if diagnostic.code == "DOCX-SMARTDOC-BEST-EFFORT-LOSS"
        ]

    def test_dependency_missing_is_surface_cleanly(self, pipeline, tmp_path, monkeypatch) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_document.to_document import converter as smartdoc_module

        def fake_fail(input_path, output_path, **kwargs):
            return BridgeResult(False, message="Install LibreOffice.")

        monkeypatch.setattr(smartdoc_module, "convert_with_backend_priority", fake_fail)

        _plugin, task_mgr, _ws_mgr = pipeline
        output_dir = tmp_path / "missing_backend"
        output_dir.mkdir()
        result = _run_smartconv(task_mgr, "docx", "doc", output_dir)

        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "dependency_missing"
        assert result.error.diagnostic_code == "DOCX-SMARTDOC-BACKEND"

    @pytest.mark.parametrize("source_fmt,target_fmt", [("docx", "odt"), ("odt", "docx")])
    def test_odt_routes_skip_wps_writer_candidate(
        self, pipeline, tmp_path, monkeypatch, source_fmt, target_fmt
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_document.to_document import converter as smartdoc_module

        calls: list[dict[str, Any]] = []

        def fake_convert(input_path, output_path, **kwargs):
            calls.append(kwargs)
            Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(smartdoc_module, "convert_with_backend_priority", fake_convert)

        _plugin, task_mgr, _ws_mgr = pipeline
        output_dir = tmp_path / f"{source_fmt}_{target_fmt}"
        output_dir.mkdir()
        result = _run_smartconv(task_mgr, source_fmt, target_fmt, output_dir)

        assert result.success is True
        assert len(calls) == 1
        assert calls[0]["libreoffice_format"] == target_fmt
        assert set(calls[0]["com_candidates"]) == {"msoffice_word"}

    def test_docx_to_rtf_prefers_word_then_keeps_configured_fallback_order(
        self, pipeline, tmp_path, monkeypatch
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_document.to_document import converter as smartdoc_module

        calls: list[dict[str, Any]] = []

        def fake_convert(input_path, output_path, **kwargs):
            calls.append(kwargs)
            Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(smartdoc_module, "convert_with_backend_priority", fake_convert)

        _plugin, task_mgr, _ws_mgr = pipeline
        output_dir = tmp_path / "docx_rtf"
        output_dir.mkdir()
        result = _run_smartconv(task_mgr, "docx", "rtf", output_dir)

        assert result.success is True
        assert len(calls) == 1
        assert calls[0]["libreoffice_format"] == "rtf"
        assert calls[0]["backend_priority"] == ["msoffice_word", "wps_writer", "libreoffice"]
        assert set(calls[0]["com_candidates"]) == {"wps_writer", "msoffice_word"}

    def test_docx_to_rtf_does_not_add_excluded_word_backend(self, pipeline, tmp_path, monkeypatch) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_document.to_document import converter as smartdoc_module

        calls: list[dict[str, Any]] = []

        def fake_convert(input_path, output_path, **kwargs):
            calls.append(kwargs)
            Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(smartdoc_module, "convert_with_backend_priority", fake_convert)

        _plugin, task_mgr, _ws_mgr = pipeline
        output_dir = tmp_path / "docx_rtf_word_excluded"
        output_dir.mkdir()
        result = _run_smartconv(
            task_mgr,
            "docx",
            "rtf",
            output_dir,
            config_snapshot={"software": {"default_priority": {"word_processors": ["libreoffice", "wps_writer"]}}},
        )

        assert result.success is True
        assert calls[0]["backend_priority"] == ["libreoffice", "wps_writer"]


class TestManifestCoverage:
    def test_all_20_routes_declared_in_manifest(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        manifest_pairs = {(route.source_format, route.target_format) for route in plugin.manifest.routes}
        missing = [f"{src}->{tgt}" for src, tgt in ALL_SMARTCONV_PAIRS if (src, tgt) not in manifest_pairs]
        assert not missing

    def test_can_handle_all_20_routes(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        missing = [f"{src}->{tgt}" for src, tgt in ALL_SMARTCONV_PAIRS if not plugin.can_handle(src, tgt)]
        assert not missing

    def test_manifest_exactly_22_routes(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        assert len(plugin.manifest.routes) == 22

    def test_extra_external_bridge_routes_entry(self) -> None:
        from docwen_plugin_document import DocumentPlugin

        plugin = DocumentPlugin()
        extra = plugin.manifest.extra

        assert "external_bridge_routes" in extra
        routes = extra["external_bridge_routes"]
        assert isinstance(routes, list)
        assert any("smartconverter" in str(item).lower() for item in routes)


class TestSmartConverterDependencyBoundary:
    def test_to_document_does_not_import_runtime(self) -> None:
        import sys

        forbidden = {
            "docwen_runtime",
            "docwen_application",
            "docwen_gui",
            "docwen_cli",
            "docwen_bundle",
        }

        module_names = [
            "docwen_plugin_document.to_document",
            "docwen_plugin_document.to_document.converter",
        ]

        for mod_name in module_names:
            mod = sys.modules.get(mod_name)
            if mod is None:
                continue
            for attr_name in dir(mod):
                val = getattr(mod, attr_name, None)
                if val is None:
                    continue
                val_mod = getattr(val, "__module__", "")
                if val_mod:
                    for forbidden_pkg in forbidden:
                        assert not val_mod.startswith(forbidden_pkg), (
                            f"{mod_name}.{attr_name} depends on {val_mod} (forbidden: {forbidden_pkg})"
                        )
