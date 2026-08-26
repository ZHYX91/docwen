from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def test_create_runtime_port_builds_with_no_default_plugins(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_bundle.runtime_factory as runtime_factory

    monkeypatch.setattr(runtime_factory, "_DEFAULT_PLUGIN_IMPORTS", [])

    port = runtime_factory.create_runtime_port()

    assert port is not None


def test_runtime_workspace_is_bound_to_explicit_governed_root(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    import docwen_bundle.runtime_factory as runtime_factory

    governed_root = tmp_path / ".workspace"
    (governed_root / "temp").mkdir(parents=True)
    (governed_root / "README.md").write_text("# DocWen 本地工作区\n", encoding="utf-8")
    monkeypatch.setenv("DOCWEN_WORKSPACE_ROOT", str(governed_root))

    assert runtime_factory._runtime_workspace_root() == governed_root / "temp" / "runtime"


def test_runtime_workspace_rejects_invalid_explicit_root(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    import docwen_bundle.runtime_factory as runtime_factory

    monkeypatch.setenv("DOCWEN_WORKSPACE_ROOT", str(tmp_path))

    with pytest.raises(RuntimeError, match="invalid governed DocWen workspace"):
        runtime_factory._runtime_workspace_root()


def test_create_runtime_port_wires_config_loader_to_adapter(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_bundle.runtime_factory as runtime_factory
    import docwen_runtime.adapters as adapters_module
    from docwen_core.models.task import TaskEvent
    from docwen_runtime.config import ConfigLoader

    monkeypatch.setattr(runtime_factory, "_DEFAULT_PLUGIN_IMPORTS", [])

    captured: dict[str, object] = {}

    def event_callback(_event: TaskEvent) -> None:
        pass

    class _Adapter:
        def __init__(
            self,
            *,
            task_manager,
            event_callback=None,
            config_loader=None,
            capability_provider=None,
            output_manifest_writer=None,
        ) -> None:
            captured["task_manager"] = task_manager
            captured["event_callback"] = event_callback
            captured["config_loader"] = config_loader
            captured["capability_provider"] = capability_provider
            captured["output_manifest_writer"] = output_manifest_writer

    monkeypatch.setattr(adapters_module, "RuntimePortAdapter", _Adapter)

    config_loader = ConfigLoader()
    port = runtime_factory.create_runtime_port(
        config_loader=config_loader,
        event_callback=event_callback,
    )

    assert isinstance(port, _Adapter)
    assert captured["event_callback"] is event_callback
    assert captured["config_loader"] is config_loader
    assert callable(captured["capability_provider"])
    assert captured["output_manifest_writer"] is not None


def test_runtime_factory_exposes_successful_empty_capability_projection(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_bundle.runtime_factory as runtime_factory

    monkeypatch.setattr(runtime_factory, "_DEFAULT_PLUGIN_IMPORTS", [])
    port = runtime_factory.create_runtime_port()
    try:
        projection = port.describe_capabilities()
    finally:
        port.shutdown()

    assert projection["runtime"]["state"] == "available"
    assert projection["sources"] == []
    assert projection["counts"]["routes"] == 0


def test_factory_runtime_port_direct_execute_cannot_bypass_file_admission(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    import docwen_bundle.runtime_factory as runtime_factory
    from docwen_core.detection import FileAdmissionError
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest

    monkeypatch.setattr(runtime_factory, "_DEFAULT_PLUGIN_IMPORTS", [])
    source = tmp_path / "disguised.docx"
    source.write_bytes(b"%PDF-1.4\ncontent\n")
    request = ConversionRequest(
        request_id="factory-direct-admission",
        input_refs=[FileRef(path=str(source), format="docx", category="document")],
        target_format="md",
    )
    port = runtime_factory.create_runtime_port()
    try:
        with pytest.raises(FileAdmissionError) as exc_info:
            port.execute(request)
    finally:
        port.shutdown()

    assert exc_info.value.error_type == "file_format_confirmation_required"


def test_create_runtime_port_fails_when_required_default_plugin_is_unavailable(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_bundle.runtime_factory as runtime_factory

    monkeypatch.setattr(runtime_factory, "_DEFAULT_PLUGIN_IMPORTS", ["required_plugin"])

    def fail_registration(import_path: str, _registry: object) -> None:
        raise ModuleNotFoundError(import_path)

    monkeypatch.setattr(runtime_factory, "_register_plugin", fail_registration)

    with pytest.raises(RuntimeError, match="Failed to load required default plugins: required_plugin"):
        runtime_factory.create_runtime_port()


def test_create_runtime_port_keeps_unavailable_extra_plugins_optional(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_bundle.runtime_factory as runtime_factory

    monkeypatch.setattr(runtime_factory, "_DEFAULT_PLUGIN_IMPORTS", [])

    def fail_registration(import_path: str, _registry: object) -> None:
        raise ModuleNotFoundError(import_path)

    monkeypatch.setattr(runtime_factory, "_register_plugin", fail_registration)

    port = runtime_factory.create_runtime_port(extra_plugins=["optional_plugin"])

    assert port is not None


@pytest.mark.parametrize("missing_export", ["PLUGIN_MANIFEST", "PLUGIN_CLASS"])
def test_register_plugin_rejects_modules_missing_required_exports(
    monkeypatch: pytest.MonkeyPatch,
    missing_export: str,
) -> None:
    import importlib
    from types import SimpleNamespace

    import docwen_bundle.runtime_factory as runtime_factory
    from docwen_runtime.plugin_registry.registry import PluginRegistry

    exports = {"PLUGIN_MANIFEST": object(), "PLUGIN_CLASS": object()}
    del exports[missing_export]
    monkeypatch.setattr(importlib, "import_module", lambda _import_path: SimpleNamespace(**exports))

    with pytest.raises(AttributeError, match=missing_export):
        runtime_factory._register_plugin("incomplete_plugin", PluginRegistry())
