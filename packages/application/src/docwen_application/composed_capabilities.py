"""Promote application document chains to executable Machine capabilities."""

from __future__ import annotations

from collections.abc import Mapping
from copy import deepcopy
from dataclasses import replace
from typing import Any

from docwen_application.controller import CapabilityUnavailableError
from docwen_application.conversion_capabilities import CAPABILITY_BINDINGS, CAPABILITY_BY_ID, CapabilityBinding
from docwen_application.conversion_contracts import DOCX_TO_MARKDOWN_CAPABILITY_ID
from docwen_application.conversion_routes import resolve_conversion_route_plan
from docwen_application.optimization_catalog import OptimizationCatalog
from docwen_core.formats import get_media_type

_DOCUMENT_SOURCES = ("docx", "doc", "wps", "rtf", "odt")


def composed_capability_bindings(
    catalog: OptimizationCatalog,
    raw_routes: Mapping[str, dict[str, Any]],
) -> tuple[CapabilityBinding, ...]:
    """Bind supported Bundle shapes to real optimizer declarations.

    Document-to-Markdown optimizers use the document/resources Bundle mapping.
    Other optimizer families remain discoverable runtime resources, but do not
    acquire an invented Machine output contract by looking like a base route.
    """

    base = CAPABILITY_BY_ID[DOCX_TO_MARKDOWN_CAPABILITY_ID]
    bindings = list(CAPABILITY_BINDINGS)
    bindings.extend(
        replace(
            base,
            capability_id=f"convert.{source}.to_markdown",
            input_format=source,
            input_media_type=get_media_type(source),
        )
        for source in _DOCUMENT_SOURCES[1:]
    )
    for resource in catalog.resources:
        for source in _DOCUMENT_SOURCES:
            plan = resolve_conversion_route_plan(
                catalog.runtime_catalog,
                source_format=source,
                source_category="document",
                target_format="md",
                action_name=resource.action_name,
            )
            if plan is None or not any(binding.route_id == plan.final_route.id for binding in resource.bindings):
                continue
            route = raw_routes[plan.final_route.id]
            schema = route.get("options_schema")
            if (
                not isinstance(schema, dict)
                or schema.get("type") != "object"
                or schema.get("additionalProperties") is not False
                or not isinstance(schema.get("properties"), dict)
                or set(schema["properties"]) != set(plan.final_route.options)
                or not all(isinstance(value, dict) for value in schema["properties"].values())
            ):
                raise CapabilityUnavailableError(
                    f"Optimization route has no complete options contract: {plan.final_route.id}"
                )
            options_schema = deepcopy(schema)
            properties = options_schema["properties"]
            for internal, public in (
                ("to_md_enable_ocr", "recognize_text"),
                ("to_md_keep_images", "preserve_resources"),
            ):
                if internal in properties:
                    if public in properties:
                        raise CapabilityUnavailableError(
                            f"Optimization route has conflicting option names: {plan.final_route.id}"
                        )
                    properties[public] = properties.pop(internal)
                    options_schema["required"] = [
                        public if item == internal else item for item in options_schema.get("required", [])
                    ]
            bindings.append(
                replace(
                    base,
                    capability_id=f"optimize.{resource.id}.{source}.to_markdown",
                    operation="transform",
                    optimization_id=resource.id,
                    input_format=source,
                    input_media_type=get_media_type(source),
                    runtime_route_id=plan.final_route.id,
                    action_name=resource.action_name,
                    options_schema=options_schema,
                    output_shape=replace(
                        base.output_shape, relation_types=(*base.output_shape.relation_types, "attachment_of")
                    ),
                )
            )
    return tuple(bindings)
