"""Translate an accepted application plan to its runtime request and source identity."""

from __future__ import annotations

import os
from typing import Any

from docwen_application.conversion_capabilities import (
    CapabilityBinding,
)
from docwen_application.conversion_contracts import (
    ConversionPlanRequest,
)
from docwen_core.models import (
    ConversionManifestContext,
    ConversionManifestInput,
    ConversionRequest,
    FileRef,
    OutputManifestPolicy,
    OutputPolicy,
)


def build_conversion_request(
    task_id: str,
    request: ConversionPlanRequest,
    binding: CapabilityBinding,
    effective_options: dict[str, Any],
) -> ConversionRequest:
    manifest_context = ConversionManifestContext(
        policy=OutputManifestPolicy(save_to_output=False, mask_input_path=True),
        inputs=tuple(
            ConversionManifestInput(
                path=handle.path,
                format=binding.input_format,
                category=binding.input_category,
            )
            for handle in request.inputs
            if handle.role in {"source", "neutral_document"}
        ),
    )
    runtime_options = dict(effective_options)
    public_properties = binding.options_schema.get("properties", {})
    if isinstance(public_properties, dict) and {
        "recognize_text",
        "preserve_resources",
    }.issubset(public_properties):
        runtime_options["to_md_enable_ocr"] = runtime_options.pop("recognize_text")
        runtime_options["to_md_keep_images"] = runtime_options.pop("preserve_resources")

    return ConversionRequest(
        request_id=task_id,
        input_refs=[
            FileRef(
                path=os.path.abspath(handle.path),
                format=(binding.input_format if handle.role in {"source", "neutral_document"} else "resource"),
                category=(binding.input_category if handle.role in {"source", "neutral_document"} else "other"),
                size_bytes=handle.size_bytes,
                input_kind=handle.kind,
                input_role=handle.role,
                logical_path=handle.logical_path,
                media_type=handle.media_type,
                metadata={
                    "machine_input_id": handle.input_id,
                    "machine_input_size_bytes": handle.size_bytes,
                    "machine_input_sha256": handle.sha256,
                },
            )
            for handle in request.inputs
        ],
        target_format=binding.target_format,
        action_name=binding.action_name,
        options=runtime_options,
        output_policy=OutputPolicy(
            output_dir=os.path.abspath(request.output.staging_root),
            overwrite_mode="error",
            write_artifacts=True,
            open_after_done=False,
        ),
        # The controller captures the complete request-scoped config snapshot.
        # Manifest policy is carried independently so Machine tasks never
        # publish the legacy sidecar manifest into their staging bundle.
        config_snapshot={},
        manifest_context=manifest_context,
    )
