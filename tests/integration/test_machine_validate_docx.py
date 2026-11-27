"""Existing-DOCX admission and cancellation through the composed Machine service."""

from __future__ import annotations

import hashlib
from dataclasses import replace
from pathlib import Path

import pytest
from docx import Document

from docwen_application.controller import ApplicationController
from docwen_application.conversion_contracts import (
    DOCX_MEDIA_TYPE,
    ConversionPlanRequest,
    ConversionServiceError,
    LocalInputHandle,
    StagingOutputTarget,
)
from docwen_application.conversion_service import ConversionService
from docwen_bundle.config_port import ConfigPortAdapter
from docwen_bundle.runtime_factory import create_runtime_port
from docwen_runtime.config import ConfigLoader
from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]
RULES = ("enable_symbol_pairing", "enable_symbol_correction", "enable_typos_rule", "enable_sensitive_word")


@pytest.fixture
def machine(tmp_path, monkeypatch):
    monkeypatch.setenv("DOCWEN_RUNTIME_ROOT", str(tmp_path / "runtime"))
    loader = ConfigLoader(user_dir=tmp_path / "config", runtime_overrides={"logger": {"console_enable": False}})
    config = ConfigPortAdapter(loader)
    controller = ApplicationController(runtime_port=create_runtime_port(config_loader=loader), config_port=config)
    controller.start()
    try:
        yield ConversionService(controller, ArtifactBundleCommitter()), config
    finally:
        controller.stop()


@pytest.fixture
def request_docx(tmp_path):
    source = tmp_path / "original.docx"
    doc = Document()
    doc.add_paragraph("A １ B")
    doc.save(source)
    original = source.read_bytes()
    staging = tmp_path / "output"
    staging.mkdir()
    handle = LocalInputHandle(
        "doc",
        str(source),
        DOCX_MEDIA_TYPE,
        len(original),
        hashlib.sha256(original).hexdigest(),
        "document",
        "source",
        source.name,
    )
    yield ConversionPlanRequest("validate.docx", (handle,), StagingOutputTarget(str(staging)), {})
    assert source.read_bytes() == original


def test_machine_plans_and_freezes_inherited_docx_rules(machine, request_docx):
    service, config = machine
    assert config.set_many({f"proofread.engine.{key}": False for key in RULES})
    with pytest.raises(ConversionServiceError) as disabled:
        service.plan(request_docx)
    assert (disabled.value.category, disabled.value.code) == ("invalid_request", "proofread_no_rules_enabled")
    assert not list(Path(request_docx.output.staging_root).iterdir())

    request = replace(request_docx, options={"enable_symbol_correction": True})
    plan = service.plan(request)
    expected = {key: key == "enable_symbol_correction" for key in RULES}
    assert plan.effective_options == expected
    assert config.set("proofread.engine.enable_symbol_pairing", True)
    with pytest.raises(ConversionServiceError) as changed:
        service.accept(plan.plan_id)
    assert changed.value.code == "plan_options_changed"
    assert not list(Path(request.output.staging_root).iterdir())

    plan = service.plan(request_docx)
    task = service.accept(plan.plan_id)
    # Disable the only enabled rule after acceptance. This accepted plan still runs.
    assert config.set("proofread.engine.enable_symbol_pairing", False)
    outcome = service.execute_accepted(task)
    assert outcome.state == "completed" and outcome.bundle is not None
    [artifact] = outcome.bundle.artifacts
    assert list(Document(Path(request.output.staging_root) / artifact.locator).comments) == []


def test_missing_docx_dependency_rejects_plan_and_acceptance(machine, request_docx, monkeypatch):
    from docwen_runtime import capabilities

    service, _config = machine
    plan = service.plan(request_docx)
    installed = capabilities._module_available
    monkeypatch.setattr(capabilities, "_module_available", lambda module: module != "docx" and installed(module))
    capability = next(c for c in service.list_capabilities() if c.capability_id == "validate.docx")
    assert capability.availability == "unavailable"
    assert capability.dependencies == ({"dependency_id": "python.docx", "required": True, "available": False},)
    for admit in (lambda: service.plan(request_docx), lambda: service.accept(plan.plan_id)):
        with pytest.raises(ConversionServiceError) as missing:
            admit()
        assert missing.value.code == "capability_unavailable"
        assert missing.value.details["missing_required_dependencies"] == ["python.docx"]
    assert not list(Path(request_docx.output.staging_root).iterdir())


def test_cancellation_during_docx_review_delivers_no_bundle(machine, request_docx, monkeypatch):
    from docwen_plugin_proofread import docx_validator

    service, _config = machine
    task_id = service.accept(service.plan(request_docx).plan_id)
    iterate = docx_validator._iter_main_story_paragraphs
    cancellations = []

    def cancel_during_traversal(container, check_cancelled):
        cancellations.append(service.cancel(task_id))
        yield from iterate(container, check_cancelled)

    monkeypatch.setattr(docx_validator, "_iter_main_story_paragraphs", cancel_during_traversal)
    outcome = service.execute_accepted(task_id)
    assert cancellations == ["cancellation_requested"]
    assert outcome.state == "cancelled" and outcome.bundle is None
    assert not list(Path(request_docx.output.staging_root).iterdir())


def test_untrusted_proofread_config_rejects_admission_and_recovers(machine, request_docx, monkeypatch):
    service, config = machine
    plan = service.plan(request_docx)

    def unavailable(_key, _default):
        raise RuntimeError("private configuration path")

    with monkeypatch.context() as patch:
        patch.setattr(config, "get", unavailable)
        for admit in (lambda: service.plan(request_docx), lambda: service.accept(plan.plan_id)):
            with pytest.raises(ConversionServiceError) as error:
                admit()
            assert (error.value.category, error.value.code) == ("unavailable", "proofread_config_unavailable")
            assert "private configuration path" not in str(error.value)
    assert not list(Path(request_docx.output.staging_root).iterdir())
    assert service.plan(request_docx).capability_id == "validate.docx"


def test_failed_docx_save_keeps_private_workspace_out_of_machine_result(machine, request_docx, monkeypatch):
    from docx.document import Document as DocxDocument

    def fail_save(_doc, staging_path):
        raise OSError(f"could not write private workspace {staging_path}")

    monkeypatch.setattr(DocxDocument, "save", fail_save)
    service, _config = machine
    outcome = service.execute_accepted(service.accept(service.plan(request_docx).plan_id))
    assert outcome.state == "failed" and outcome.bundle is None
    assert outcome.error is not None and outcome.error.diagnostic_code == "PROOFREAD-ERROR"
    assert "private workspace" not in outcome.error.message
    assert all("private workspace" not in item.message for item in outcome.diagnostics)
    assert not list(Path(request_docx.output.staging_root).iterdir())
