"""Resolve and validate request options against a selected capability contract."""

from __future__ import annotations

import math
from copy import deepcopy
from typing import Any

from docwen_application.conversion_capabilities import (
    CapabilityBinding,
)
from docwen_application.conversion_contracts import (
    ConversionServiceError,
)


def resolve_conversion_options(
    provided: dict[str, Any],
    binding: CapabilityBinding,
) -> dict[str, Any]:
    schema = binding.options_schema
    properties = schema.get("properties", {})
    if not isinstance(properties, dict):
        raise ConversionServiceError(
            "internal",
            "capability_options_schema_invalid",
            "capability has an invalid options schema",
        )
    unknown = sorted(set(provided) - set(properties))
    if unknown:
        raise ConversionServiceError(
            "invalid_request",
            "unsupported_options",
            "capability does not accept one or more caller-defined options",
            details={"option_keys": unknown},
        )
    effective = {
        key: property_schema["default"]
        for key, property_schema in properties.items()
        if isinstance(property_schema, dict) and "default" in property_schema
    }
    effective.update(binding.effective_options)
    effective.update(provided)
    missing = [key for key in schema.get("required", []) if key not in effective]
    if missing:
        raise ConversionServiceError(
            "invalid_request",
            "required_options_missing",
            "capability requires one or more options",
            details={"option_keys": sorted(missing)},
        )
    for key, value in effective.items():
        property_schema = properties.get(key)
        if property_schema is None:
            continue
        _validate_option_value(key, value, property_schema)
    return deepcopy(effective)


def _validate_option_value(key: str, value: Any, schema: dict[str, Any]) -> None:
    raw_expected_type = schema.get("type")
    expected_type = raw_expected_type if isinstance(raw_expected_type, str) else ""
    valid_type = {
        "boolean": isinstance(value, bool),
        "string": isinstance(value, str),
        "integer": isinstance(value, int) and not isinstance(value, bool),
        "number": isinstance(value, (int, float)) and not isinstance(value, bool),
        "array": isinstance(value, list),
        "object": isinstance(value, dict),
    }.get(expected_type, True)
    invalid = not valid_type
    if not invalid and "enum" in schema:
        invalid = value not in schema["enum"]
    if not invalid and isinstance(value, (int, float)) and not isinstance(value, bool):
        minimum = schema.get("minimum")
        maximum = schema.get("maximum")
        invalid = (
            (isinstance(value, float) and not math.isfinite(value))
            or (minimum is not None and value < minimum)
            or (maximum is not None and value > maximum)
        )
    if not invalid and isinstance(value, dict):
        properties = schema.get("properties", {})
        additional = schema.get("additionalProperties", True)
        if any(name not in value for name in schema.get("required", [])):
            invalid = True
        else:
            for name, item in value.items():
                item_schema = properties.get(name, additional)
                if item_schema is False:
                    invalid = True
                    break
                if isinstance(item_schema, dict):
                    _validate_option_value(f"{key}.{name}", item, item_schema)
    if not invalid and isinstance(value, list) and isinstance(schema.get("items"), dict):
        minimum_items = schema.get("minItems")
        if isinstance(minimum_items, int) and len(value) < minimum_items:
            invalid = True
        if schema.get("uniqueItems") is True and any(value[index] in value[:index] for index in range(len(value))):
            invalid = True
        try:
            for index, item in enumerate(value):
                _validate_option_value(f"{key}[{index}]", item, schema["items"])
        except ConversionServiceError:
            invalid = True
    if invalid:
        raise ConversionServiceError(
            "invalid_request",
            "option_value_invalid",
            f"option has a value outside its capability contract: {key}",
            details={"option_key": key},
        )
