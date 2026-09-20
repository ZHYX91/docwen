"""Planning owns its options; caller-visible dictionaries cannot change execution."""

from __future__ import annotations

from typing import Any

import pytest

from docwen_application.conversion_contracts import MARKDOWN_TO_DOCX_CAPABILITY_ID, MARKDOWN_TO_XLSX_CAPABILITY_ID
from docwen_application.conversion_service import ConversionService

from ._conversion_service_support import ConversionServiceError, Path, _Committer, _Controller, _request

pytestmark = pytest.mark.integration


@pytest.mark.parametrize(
    "extensions",
    [
        {"input": {"structural_tables": "true"}},
        {"input": {"structural_tables": 1}},
        {"input": {"unknown_extension": True}},
        {"input": []},
        {"sideways": {"structural_tables": True}},
    ],
)
def test_invalid_nested_options_fail_during_planning(tmp_path: Path, extensions: dict[str, Any]) -> None:
    source = tmp_path / "source.md"
    source.write_text("# Heading\n", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    controller = _Controller()
    service = ConversionService(controller, _Committer())
    with pytest.raises(ConversionServiceError) as error:
        service.plan(
            _request(
                source,
                staging,
                capability_id=MARKDOWN_TO_XLSX_CAPABILITY_ID,
                options={"markdown_extensions": extensions},
            )
        )
    assert error.value.code == "option_value_invalid"
    assert error.value.details["option_key"].startswith("markdown_extensions")
    assert not controller.requests and not list(staging.iterdir())


def test_public_capability_schema_cannot_change_later_plans() -> None:
    service = ConversionService(_Controller(), _Committer())
    capability = next(c for c in service.list_capabilities() if c.capability_id == MARKDOWN_TO_DOCX_CAPABILITY_ID)
    capability.options_schema["properties"]["locale"]["enum"].append("invented_locale")
    capability.limitations[0]["code"] = "caller_changed_code"
    fresh = next(c for c in service.list_capabilities() if c.capability_id == MARKDOWN_TO_DOCX_CAPABILITY_ID)
    assert "invented_locale" not in fresh.options_schema["properties"]["locale"]["enum"]
    assert fresh.limitations[0]["code"] == "resolved_document.provider_owned_semantics"


def test_request_and_returned_plan_mutations_cannot_change_accepted_options(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("# Heading\n", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    request = _request(
        source,
        staging,
        capability_id=MARKDOWN_TO_XLSX_CAPABILITY_ID,
        options={"markdown_extensions": {"input": {"structural_tables": True}}},
    )
    controller = _Controller()
    service = ConversionService(controller, _Committer())
    plan = service.plan(request)
    request.options["markdown_extensions"]["input"]["structural_tables"] = False
    plan.effective_options["markdown_extensions"]["input"]["structural_tables"] = False
    outcome = service.execute_accepted(service.accept(plan.plan_id))
    assert outcome.state == "completed"
    assert controller.requests[-1].options["markdown_extensions"] == {"input": {"structural_tables": True}}
