"""Public source-only post-render proofreading admission contracts."""

from __future__ import annotations

from itertools import product

import pytest

from docwen_application.conversion_capabilities import CAPABILITY_BY_ID
from docwen_application.conversion_options import resolve_conversion_options
from docwen_application.conversion_requests import build_conversion_request
from docwen_core.models.request import POSTPROCESS_PROOFREAD_OPTION

from ._conversion_service_support import (
    MARKDOWN_MEDIA_TYPE,
    MARKDOWN_SOURCE_TO_DOCX_CAPABILITY_ID,
    MARKDOWN_TO_DOCX_CAPABILITY_ID,
    ConversionService,
    ConversionServiceError,
    _Committer,
    _Controller,
    _request,
)

pytestmark = pytest.mark.unit

RULES = ("enable_symbol_pairing", "enable_symbol_correction", "enable_typos_rule", "enable_sensitive_word")


def test_public_schema_is_closed_source_only_without_defaults():
    source = CAPABILITY_BY_ID[MARKDOWN_SOURCE_TO_DOCX_CAPABILITY_ID]
    schema = source.options_schema["properties"]["proofread"]
    assert schema == {
        "type": "object",
        "properties": {key: {"type": "boolean"} for key in RULES},
        "required": [],
        "additionalProperties": False,
    }
    assert "proofread" not in CAPABILITY_BY_ID[MARKDOWN_TO_DOCX_CAPABILITY_ID].options_schema["properties"]
    assert all(POSTPROCESS_PROOFREAD_OPTION not in b.options_schema["properties"] for b in CAPABILITY_BY_ID.values())


@pytest.mark.parametrize(
    "checks", [None, {}, *[dict(zip(RULES, values, strict=True)) for values in product((False, True), repeat=4)]]
)
def test_plan_preserves_public_options_and_maps_only_at_runtime(tmp_path, monkeypatch, checks):
    source = tmp_path / "note.md"
    source.write_text("# Note\n", encoding="utf-8")
    staging = tmp_path / "output"
    staging.mkdir()
    options = {} if checks is None else {"proofread": checks}
    request = _request(
        source,
        staging,
        capability_id=MARKDOWN_SOURCE_TO_DOCX_CAPABILITY_ID,
        media_type=MARKDOWN_MEDIA_TYPE,
        options=options,
    )
    controller = _Controller()
    monkeypatch.setattr(
        controller, "_ROUTE_IDS", (*controller._ROUTE_IDS, "docwen_plugin_proofread:docx:docx:validate")
    )
    service = ConversionService(controller, _Committer())
    plan = service.plan(request)
    assert POSTPROCESS_PROOFREAD_OPTION not in plan.effective_options
    assert plan.effective_options.get("proofread") == checks
    runtime = build_conversion_request("task", request, CAPABILITY_BY_ID[request.capability_id], plan.effective_options)
    assert "proofread" not in runtime.options
    assert (POSTPROCESS_PROOFREAD_OPTION in runtime.options) == (checks is not None)
    if checks is not None:
        assert runtime.options[POSTPROCESS_PROOFREAD_OPTION] == checks


@pytest.mark.parametrize(
    "value",
    [
        True,
        False,
        None,
        [],
        "all",
        1,
        {"enabled": True},
        {"skip_code_blocks": True},
        *[{rule: value} for rule in RULES for value in (0, 1, "true", None, [])],
    ],
)
def test_invalid_postprocess_payload_rejected(value):
    with pytest.raises(ConversionServiceError) as error:
        resolve_conversion_options({"proofread": value}, CAPABILITY_BY_ID[MARKDOWN_SOURCE_TO_DOCX_CAPABILITY_ID])
    assert error.value.code == "option_value_invalid"


@pytest.mark.parametrize("key", ["proofread", POSTPROCESS_PROOFREAD_OPTION, *RULES])
def test_frozen_resolved_port_rejects_proofreading(key):
    with pytest.raises(ConversionServiceError) as error:
        resolve_conversion_options({key: {}}, CAPABILITY_BY_ID[MARKDOWN_TO_DOCX_CAPABILITY_ID])
    assert error.value.code == "unsupported_options"


@pytest.mark.parametrize("checks", [{}, {"enable_typos_rule": True}])
def test_missing_validator_fails_at_plan_without_affecting_old_conversion(tmp_path, checks):
    source = tmp_path / "note.md"
    source.write_text("# Note\n", encoding="utf-8")
    staging = tmp_path / "output"
    staging.mkdir()
    service = ConversionService(_Controller(), _Committer())
    request = _request(
        source,
        staging,
        capability_id=MARKDOWN_SOURCE_TO_DOCX_CAPABILITY_ID,
        media_type=MARKDOWN_MEDIA_TYPE,
        options={"proofread": checks},
    )
    with pytest.raises(ConversionServiceError) as error:
        service.plan(request)
    assert error.value.code == "proofread_postprocess_unavailable"
    from dataclasses import replace

    service.plan(replace(request, options={}))
    service.plan(replace(request, options={"proofread": dict.fromkeys(RULES, False)}))
    assert not list(staging.iterdir())


def test_normative_discovery_fixture_matches_runtime_schema():
    import json
    from pathlib import Path

    fixture = (
        Path(__file__).resolve().parents[3]
        / "contracts/fixtures/valid/machine.source-proofread-capability.response.json"
    )
    [capability] = json.loads(fixture.read_text(encoding="utf-8"))["result"]["capabilities"]
    binding = CAPABILITY_BY_ID[MARKDOWN_SOURCE_TO_DOCX_CAPABILITY_ID]
    assert capability["options_schema"] == binding.options_schema
    assert capability["input_shape"] == binding.input_shape.to_dict()
    assert capability["output_shape"] == binding.output_shape.to_dict()


@pytest.mark.parametrize("key", [POSTPROCESS_PROOFREAD_OPTION, *RULES])
def test_source_port_rejects_private_or_flat_rule_options(key):
    with pytest.raises(ConversionServiceError) as error:
        resolve_conversion_options({key: True}, CAPABILITY_BY_ID[MARKDOWN_SOURCE_TO_DOCX_CAPABILITY_ID])
    assert error.value.code == "unsupported_options"


def test_validator_route_is_rechecked_at_acceptance(tmp_path, monkeypatch):
    source = tmp_path / "note.md"
    source.write_text("# Note\n", encoding="utf-8")
    staging = tmp_path / "output"
    staging.mkdir()
    controller = _Controller()
    monkeypatch.setattr(
        controller, "_ROUTE_IDS", (*controller._ROUTE_IDS, "docwen_plugin_proofread:docx:docx:validate")
    )
    service = ConversionService(controller, _Committer())
    request = _request(
        source,
        staging,
        capability_id=MARKDOWN_SOURCE_TO_DOCX_CAPABILITY_ID,
        media_type=MARKDOWN_MEDIA_TYPE,
        options={"proofread": {}},
    )
    plan = service.plan(request)
    controller._ROUTE_IDS = _Controller._ROUTE_IDS
    with pytest.raises(ConversionServiceError) as error:
        service.accept(plan.plan_id)
    assert error.value.code == "proofread_postprocess_unavailable"
    assert not controller.requests and not list(staging.iterdir())
